# uia-common.ps1 — shared scaffold for the showcase UIA repro/verify scripts.
#
# Dot-source this from a sibling script:   . "$PSScriptRoot\uia-common.ps1"
#
# It provides the common plumbing every showcase UIA driver needs:
#   * XLL / build-dir / log path resolution         (Resolve-ShowcasePaths)
#   * Trust Center precondition gate                (Assert-ExcelTrustPreconditions,
#       Test-ExcelTrustedLocation, Clear-ExcelResiliency) -- call it FIRST; an
#       unmet lever fakes the very defects these scripts hunt
#   * "was the add-in even loaded?" ladder          (Get-XllLoadFailureDiagnosis,
#       Write-XllLoadDiagnosis, Get-ExcelDisabledItems) -- call it AT THE POINT
#       OF FAILURE; it separates "never loaded" (environment) from "a feature is
#       broken" (product), which otherwise look identical
#   * stale Excel/server process kill               (Stop-ShowcaseProcesses)
#   * UIA bootstrap ($AE/$TS) + Restart-Manager lock probe (Initialize-Uia, Probe-Lock)
#   * visible Excel window enumeration              (Get-ExcelWindows)
#   * MODAL dialog detection for polling loops       (Get-ExcelModalDialog,
#       Assert-NoExcelModal) -- a modal turns every poll into a timeout that
#       reads as "the feature under test is wedged"
#   * blank-workbook prep via a SHORT-LIVED COM client (New-ShowcaseWorkbook)
#   * launch real Excel + wait for the ready window (Start-ShowcaseExcel / Wait-ShowcaseWindow)
#   * select the showcase ribbon tab + click Build  (Invoke-RibbonButton / Invoke-BuildShowcase)
#   * FAITHFUL window close via WindowPattern.Close()+Don't-Save (Close-ExcelWindowFaithful)
#   * the project's mandatory TWO-TIER teardown      (Stop-ShowcaseProcesses for kill-only,
#       Stop-ShowcaseCom for COM graceful-Quit -> force-kill)
#
# Each consuming script keeps only its own sampling/verdict logic.
#
# SELF-TEST: tools\test-uia-common.ps1 covers everything here that can be checked
# without Excel, including the modal detector against a REAL dialog. Run it after
# touching this file -- it caught two stacked bugs in the detector on the day the
# detector was written, one of which made it silently never fire.
#
# ----------------------------------------------------------------------------
# ENCODING: this file is ASCII-ONLY on purpose. A .ps1 without a UTF-8 BOM is
# read by Windows PowerShell under the system ANSI codepage (e.g. CP949 on
# Korean Windows), which mangles any non-ASCII (Korean) literal into bytes that
# either break parsing or form an INVALID regex whose exception is swallowed by
# a surrounding try/catch -- which silently dropped the real Excel window and
# produced "FAIL: no ready window". So localized strings (e.g. the Korean
# "Don't Save" button) are built by codepoint, never as a source literal.
# ----------------------------------------------------------------------------

$xlexe = "C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE"

# --- path resolution --------------------------------------------------------
# Resolves the built XLL, the build dir, and the two server logs relative to
# the CALLING script's tools\ directory. Returns a hashtable; callers usually
# splat into locals:  $P = Resolve-ShowcasePaths; $xll = $P.Xll
function Resolve-ShowcasePaths {
    $buildDir  = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\build")
    [pscustomobject]@{
        Xlexe     = $xlexe
        Xll       = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\build\xll_showcase.xll")
        BuildDir  = $buildDir
        GoLog     = Join-Path $buildDir "xll_showcase_go.log"
        NativeLog = Join-Path $buildDir "xll_showcase_native.log"
    }
}

# --- Trust Center preconditions ---------------------------------------------
# WHY THIS EXISTS: every one of these scripts reports "the product is broken"
# using symptoms that an unmet Trust Center precondition produces IDENTICALLY.
# Three DIFFERENT levers govern three DIFFERENT things, and 2026-07-29 hit all
# three in sequence while blaming the XLL:
#
#   (1) Trusted Locations\LocationN (Path + AllowSubfolders)
#         governs whether the file/add-in is trusted at all. It ALSO exempts an
#         UNSIGNED XLL from the signature requirement below -- verified: with
#         RequireAddinSig=1 the showcase still loaded, because a trusted
#         location covered build\.
#   (2) DataConnectionWarnings
#         governs the RTD / external-data prompt. A trusted location does NOT
#         suppress it -- Excel treats RTD as a separate category. If it is not
#         0 a MODAL dialog appears, the driver polls forever behind it, and the
#         run is read as "RTD wedged" (same misdiagnosis shape as the .Text
#         trap).
#   (3) RequireAddinSig
#         if 1 AND the XLL is not covered by a trusted location, Excel blocks
#         the unsigned add-in with NO dialog and NO warning. This one is the
#         most dangerous: the only symptom is "Excel is up but there is no
#         server process", which reads exactly like a product defect.
#
# So: READ all three, decide against the ACTUAL xll path, and FAIL LOUDLY.
# This function deliberately does NOT write to the registry -- changing a
# user's security settings is theirs to do; we print the exact remediation.
#
# Returns nothing on success. Throws on an unmet precondition, so a caller that
# ignores the result still stops instead of proceeding into a bogus verdict.
#   -XllPath      : the .xll actually being loaded (its FOLDER is what matters).
#   -RequireRtd   : also enforce DataConnectionWarnings=0 (any script that waits
#                   on RTD topics; harmless to pass always).
$ExcelSecurityKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'

function Test-ExcelTrustedLocation {
    # -SecurityKey exists so the self-test can point the whole trust/resiliency
    # ladder at a SCRATCH registry subtree. Changing a real user's security
    # settings to test a detector is not acceptable, and a detector that is only
    # ever run against the machine's real settings is untested by definition.
    param([Parameter(Mandatory)][string]$Folder, [string]$SecurityKey = $ExcelSecurityKey)
    $root = Join-Path $SecurityKey 'Trusted Locations'
    if (-not (Test-Path $root)) { return $false }
    $target = [System.IO.Path]::GetFullPath($Folder).TrimEnd('\') + '\'
    foreach ($k in (Get-ChildItem $root -EA SilentlyContinue)) {
        $p = (Get-ItemProperty $k.PSPath -EA SilentlyContinue)
        if (-not $p -or -not $p.Path) { continue }
        # Path values routinely contain %USERPROFILE% and friends.
        $loc = [Environment]::ExpandEnvironmentVariables([string]$p.Path)
        try { $loc = [System.IO.Path]::GetFullPath($loc) } catch { continue }
        $loc = $loc.TrimEnd('\') + '\'
        if ($target -eq $loc) { return $true }
        if ($p.AllowSubfolders -eq 1 -and $target.StartsWith($loc, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Assert-ExcelTrustPreconditions {
    param([Parameter(Mandatory)][string]$XllPath, [switch]$RequireRtd)

    $folder  = Split-Path -Parent ([System.IO.Path]::GetFullPath($XllPath))
    $sec     = Get-ItemProperty $ExcelSecurityKey -EA SilentlyContinue
    $sig     = if ($sec -and $null -ne $sec.RequireAddinSig)         { [int]$sec.RequireAddinSig }         else { 0 }
    $dcw     = if ($sec -and $null -ne $sec.DataConnectionWarnings)  { [int]$sec.DataConnectionWarnings }  else { 1 }
    $trusted = Test-ExcelTrustedLocation -Folder $folder

    Write-Output ("PRECHECK: xll folder      = {0}" -f $folder)
    Write-Output ("PRECHECK: trusted location= {0}" -f $trusted)
    Write-Output ("PRECHECK: RequireAddinSig = {0}" -f $sig)
    Write-Output ("PRECHECK: DataConnWarnings= {0}" -f $dcw)

    $fail = @()
    if ($sig -eq 1 -and -not $trusted) {
        $fail += ("RequireAddinSig=1 and '{0}' is NOT a trusted location. Excel will block the UNSIGNED showcase XLL " -f $folder) +
                 "SILENTLY (no dialog, no log, no server process). Fix: add the folder as a Trusted Location " +
                 ("(New-Item '{0}\Trusted Locations\LocationXLLGen' -Force; " -f $ExcelSecurityKey) +
                 ("Set-ItemProperty '{0}\Trusted Locations\LocationXLLGen' Path '{1}'; " -f $ExcelSecurityKey, $folder) +
                 ("Set-ItemProperty '{0}\Trusted Locations\LocationXLLGen' AllowSubfolders 1)  -- or set RequireAddinSig=0." -f $ExcelSecurityKey)
    }
    if ($RequireRtd -and $dcw -ne 0) {
        $fail += ("DataConnectionWarnings={0} (need 0). Excel will raise a MODAL external-data prompt for the RTD " -f $dcw) +
                 "topics; the driver then polls behind it and the run is misread as 'RTD wedged'. A trusted location " +
                 ("does NOT cover this. Fix: Set-ItemProperty '{0}' DataConnectionWarnings 0 -Type DWord." -f $ExcelSecurityKey)
    }
    if (-not $trusted) {
        Write-Output ("PRECHECK: WARN - '{0}' is not a trusted location; the add-in may prompt or load degraded." -f $folder)
    }

    # Post-crash residue. DisabledItems is the other silent "add-in did not
    # load" cause, and a pending DocumentRecovery makes Excel open the recovery
    # pane, which steals the window the UIA driver is waiting for (3 of 5 runs
    # on 2026-07-29). Both are REPORTED, not cleared -- DocumentRecovery can
    # hold a real user's unsaved work, so clearing it is the caller's explicit
    # choice via Clear-ExcelResiliency.
    $res = Join-Path (Split-Path -Parent $ExcelSecurityKey) 'Resiliency'
    foreach ($sub in @('DisabledItems', 'StartupItems', 'DocumentRecovery')) {
        $p = Join-Path $res $sub
        if (Test-Path $p) {
            $k = Get-Item $p
            $n = $k.ValueCount + $k.SubKeyCount
            if ($n -gt 0) {
                Write-Output ("PRECHECK: WARN - Resiliency\{0} has {1} entries. It can block the add-in or hijack the window the driver waits for; run Clear-ExcelResiliency to drop them." -f $sub, $n)
            }
        }
    }

    if ($fail.Count -gt 0) {
        foreach ($f in $fail) { Write-Output ("PRECHECK FAIL: " + $f) }
        throw "Excel Trust Center preconditions not met -- refusing to run (see PRECHECK FAIL above). These are ENVIRONMENT problems; do not record them as product defects."
    }
    Write-Output "PRECHECK: OK"
}

# Drops post-crash Excel resiliency residue. Opt-in on purpose: DocumentRecovery
# may hold a real user's unsaved workbooks.
function Clear-ExcelResiliency {
    param([switch]$IncludeDocumentRecovery)
    $res = Join-Path (Split-Path -Parent $ExcelSecurityKey) 'Resiliency'
    $subs = @('DisabledItems', 'StartupItems')
    if ($IncludeDocumentRecovery) { $subs += 'DocumentRecovery' }
    foreach ($sub in $subs) {
        $p = Join-Path $res $sub
        if (Test-Path $p) { Remove-Item $p -Recurse -Force -EA SilentlyContinue; Write-Output ("cleared Resiliency\{0}" -f $sub) }
    }
}

# --- "was the add-in even LOADED?" exclusion ladder --------------------------
# WHY THIS EXISTS. Every driver in this folder can fail in two ways that produce
# byte-identical symptoms: no ready window, no Build button on the ribbon, a cell
# parked on the loading placeholder, a #NAME? formula, an empty log. Either
#   (a) THE ADD-IN WAS NEVER LOADED -- an ENVIRONMENT problem; every verdict in
#       the run is about code that never ran, or
#   (b) the add-in loaded and a feature is broken -- a real product defect.
# Reporting (a) as (b) is the most expensive mistake this harness can make: it
# files a bug against a binary Excel never touched.
#
# WHAT THE PRECONDITION GATE ALREADY SETTLES, STATED ACCURATELY (corrected
# 2026-08-03). This paragraph used to say the gate 'CANNOT settle this'. That is
# false for rung 1: Assert-ExcelTrustPreconditions tests the SAME
# RequireAddinSig=1 + not-a-trusted-location condition and THROWS before Excel is
# started, so in a run that called the gate, rung 1 can only fire if the registry
# changed after the gate ran -- or if the driver never called the gate at all.
# Rung 1 is kept for exactly those two cases, and because the ladder is a shared
# function with no way to know which. What the gate genuinely cannot settle is
# the rest: DisabledItems is only WARNED about (never fatal, because a stale
# entry that names something else must not block a run), the disabled-items
# decision is made by Excel at THIS load, and the native log does not exist until
# the DLL entry point runs. So the ladder is evaluated AT THE POINT OF FAILURE,
# and the rung order is not cosmetic -- each rung MASKS every rung below it, so
# the first one that fires is the answer:
#
#   0. the XLL file is not there at all
#        every rung below is about how Excel treated a file that does not exist.
#        Without this rung a leftover native log from an earlier build made the
#        ladder answer 'Loaded -- this belongs to the product' for a missing XLL.
#   1. RequireAddinSig=1 AND the folder is not a trusted location
#        Excel refuses the unsigned XLL with no dialog, no log, no process.
#        Nothing below can be observed, therefore nothing below can be believed.
#   2. Resiliency\DisabledItems  (values AND subkeys -- Excel uses both shapes)
#        a previous crash put the add-in on the disabled list; Excel skips it
#        silently on every subsequent start. Again nothing below is observable.
#   3. the native log
#        the first thing the DLL entry point writes. Missing or empty => the
#        add-in never ran. Present => it DID load, and the failure the caller is
#        holding belongs to the product.
#
# There is deliberately NO throwing variant. Every driver reaches this point
# holding live Excel and Go-server processes it still has to tear down; a throw
# would skip the two-tier teardown and leave them running, which corrupts the
# NEXT run -- the exact class of mistake this ladder exists to catch. Callers
# print the diagnosis and then continue into their own cleanup path.

# Enumerates Excel's disabled-items entries as {Name, Text}. The entries are
# REG_BINARY blobs; the path Excel disabled is carried inside as UTF-16, so the
# blob is decoded BOTH ways and reduced to printable ASCII for matching. Some
# entries carry no readable path at all, which is why the caller distinguishes
# "an entry names our file" from "there are entries" (different confidence).
#
# VALUES **AND SUBKEYS** (fixed 2026-08-03). This read only GetValueNames(), so
# a DisabledItems entry stored as a SUBKEY was invisible: the ladder skipped rung
# 2 and answered rung=Loaded -- "this failure belongs to the product" -- for an
# add-in Excel had disabled. The precheck at Assert-ExcelTrustPreconditions was
# already counting ValueCount + SubKeyCount, so the two disagreed about the same
# registry key: the run printed a residue WARN and then a Loaded verdict. A
# subkey contributes one entry per value it holds, plus one for the subkey itself
# when it holds none, so an entry cannot be counted as zero.
function Get-ExcelDisabledItems {
    param([string]$SecurityKey = $ExcelSecurityKey)
    $out = New-Object System.Collections.ArrayList
    $path = Join-Path (Split-Path -Parent $SecurityKey) 'Resiliency\DisabledItems'
    if (-not (Test-Path $path)) { return ,$out }
    $k = $null
    try { $k = Get-Item $path -EA Stop } catch { return ,$out }

    $decode = {
        param($v)
        if ($v -is [byte[]]) {
            # TWO decodes, because the UTF-16 one only lands when the path happens
            # to start on an even offset inside the blob. Dropping the NUL bytes
            # and reading what is left as ASCII recovers the path at EITHER
            # alignment, and also covers entries stored as plain bytes.
            $u = [Text.Encoding]::Unicode.GetString($v)
            $z = [Text.Encoding]::ASCII.GetString([byte[]]@($v | Where-Object { $_ -ne 0 }))
            # Non-printables become spaces; runs of spaces collapse. ASCII-only
            # source, so this regex is safe under any codepage (see ENCODING).
            (($u + ' ' + $z) -replace '[^\x20-\x7E]', ' ') -replace ' {2,}', ' '
        } else {
            [string]$v
        }
    }

    foreach ($n in $k.GetValueNames()) {
        $v = $null
        try { $v = $k.GetValue($n) } catch { continue }
        [void]$out.Add([pscustomobject]@{ Name = $n; Text = ([string](& $decode $v)).Trim() })
    }

    foreach ($sn in $k.GetSubKeyNames()) {
        $sk = $null
        try { $sk = Get-Item (Join-Path $path $sn) -EA Stop } catch { continue }
        $any = $false
        foreach ($n in $sk.GetValueNames()) {
            $v = $null
            try { $v = $sk.GetValue($n) } catch { continue }
            $any = $true
            # The subkey NAME is prepended: Excel sometimes carries the path
            # there and nowhere in the values.
            $txt = ($sn + ' ' + [string](& $decode $v)).Trim()
            [void]$out.Add([pscustomobject]@{ Name = ($sn + '\' + $n); Text = $txt })
        }
        if (-not $any) {
            [void]$out.Add([pscustomobject]@{ Name = $sn; Text = $sn })
        }
    }
    return ,$out
}

# Runs the ladder and returns a verdict object:
#   Rung        MissingXll | RequireAddinSig | DisabledItems | NoNativeLog | StaleNativeLog | Loaded
#   Loaded      $false for every rung except 'Loaded'
#   Confidence  high | low   ('low' = the rung fired on circumstantial evidence)
#   Cause / Remediation / Evidence / Message
# The Message is the printable form, and it reports the native-log state whichever
# rung fired (on its own line unless the evidence line already IS that state), so
# a low-confidence rung 2 can be weighed against it instead of taken on faith.
#   -NativeLog : defaults to the showcase native log beside the XLL.
#   -Since     : the moment this run launched Excel. A log older than that was
#                written by an EARLIER run, so it proves nothing about this one.
#   -SecurityKey : scratch-registry hook for the self-test.
function Get-XllLoadFailureDiagnosis {
    param(
        [Parameter(Mandatory)][string]$XllPath,
        [string]$NativeLog,
        [datetime]$Since,
        [string]$SecurityKey = $ExcelSecurityKey
    )
    $full   = [System.IO.Path]::GetFullPath($XllPath)
    $folder = Split-Path -Parent $full
    $leaf   = Split-Path -Leaf $full
    if (-not $NativeLog) { $NativeLog = Join-Path $folder 'xll_showcase_native.log' }
    $useSince = $PSBoundParameters.ContainsKey('Since')

    # --- observations (all three rungs are MEASURED before any is chosen, so the
    # message can report the others as context) -------------------------------
    $sec     = Get-ItemProperty $SecurityKey -EA SilentlyContinue
    $sig     = if ($sec -and $null -ne $sec.RequireAddinSig) { [int]$sec.RequireAddinSig } else { 0 }
    $trusted = Test-ExcelTrustedLocation -Folder $folder -SecurityKey $SecurityKey
    # NOT @(Get-ExcelDisabledItems ...): that helper returns a wrapped ArrayList
    # (',$out', the same stable-type idiom Get-ExcelWindows uses), so @() around
    # the call yields a ONE-element array holding the list itself -- $_.Text is
    # then $null and the ladder dies mid-diagnosis. Measured, not theorised.
    $items   = Get-ExcelDisabledItems -SecurityKey $SecurityKey
    $mine    = @($items | Where-Object { ("" + $_.Text).IndexOf($leaf, [StringComparison]::OrdinalIgnoreCase) -ge 0 })

    $logExists = Test-Path $NativeLog
    $logLen    = -1
    $logWrite  = $null
    if ($logExists) {
        try { $fi = Get-Item $NativeLog -EA Stop; $logLen = $fi.Length; $logWrite = $fi.LastWriteTime } catch { $logExists = $false }
    }
    $logNote =
        if (-not $logExists)  { "native log MISSING ({0})" -f $NativeLog }
        elseif ($logLen -eq 0) { "native log EMPTY ({0})" -f $NativeLog }
        else { "native log present ({0}, {1} bytes, last write {2:yyyy-MM-dd HH:mm:ss})" -f $NativeLog, $logLen, $logWrite }
    $logStale = ($logExists -and $logLen -gt 0 -and $useSince -and $logWrite -lt $Since)

    # The XLL's own existence. Measured with the same care as the log, because
    # this is the state that used to be answered "Loaded (measured)".
    $xllExists = Test-Path -LiteralPath $full -PathType Leaf
    $xllNote   = if ($xllExists) { "XLL present ({0})" -f $full } else { "XLL MISSING ({0})" -f $full }

    $rung = ''; $conf = 'high'; $cause = ''; $fix = ''; $ev = ''

    # --- rung 0: the file is not there at all ---------------------------------
    # Added 2026-08-03. Without it, a native log left behind by an EARLIER build
    # (probe P4) made every rung below fall through to 'Loaded' -- the ladder
    # reported "the add-in DID load, this belongs to the product" for an XLL that
    # does not exist. -Since only downgrades that to the low-confidence
    # StaleNativeLog rung, and only when the caller passes it.
    if (-not $xllExists) {
        $rung  = 'MissingXll'
        $cause = ("the add-in was NEVER LOADED: there is no file at '{0}'. Every rung below is about how Excel treated a file that is not there, and any native log present was written by an earlier build." -f $full)
        $ev    = $xllNote
        $fix   = "build the showcase XLL (task build-cpp-debug / build-cpp) or correct -XllPath, then RE-RUN; this run's verdict is void."
    }
    # --- rung 1: the signature requirement, which hides rungs 2 and 3 ---------
    elseif ($sig -eq 1 -and -not $trusted) {
        $rung  = 'RequireAddinSig'
        $cause = ("the add-in was NEVER LOADED: RequireAddinSig=1 and '{0}' is not a trusted location, so Excel refused the UNSIGNED XLL silently -- no dialog, no log, no server process." -f $folder)
        $ev    = ("RequireAddinSig=1, trusted location=False, folder={0}" -f $folder)
        $fix   = ("add the folder as a Trusted Location (New-Item '{0}\Trusted Locations\LocationXLLGen' -Force; " -f $SecurityKey) +
                 ("Set-ItemProperty '{0}\Trusted Locations\LocationXLLGen' Path '{1}'; " -f $SecurityKey, $folder) +
                 ("Set-ItemProperty '{0}\Trusted Locations\LocationXLLGen' AllowSubfolders 1) -- or set RequireAddinSig=0. Then RE-RUN; this run's verdict is void." -f $SecurityKey)
    }
    # --- rung 2: the disabled list, which hides rung 3 ------------------------
    elseif ($items.Count -gt 0) {
        $rung  = 'DisabledItems'
        if ($mine.Count -gt 0) {
            $cause = ("the add-in was NEVER LOADED: Excel's Resiliency\DisabledItems names '{0}'. A previous crash disabled it and Excel has skipped it silently at every start since." -f $leaf)
            $ev    = ("DisabledItems entry: {0}" -f ($mine[0].Text))
        } else {
            $conf  = 'low'
            $cause = ("the add-in may never have been LOADED: Resiliency\DisabledItems holds {0} entry/entries. None decodes to '{1}', but Excel also records entries with no readable path, so this cannot be cleared by inspection -- weigh it against the native-log line below." -f $items.Count, $leaf)
            $ev    = ("DisabledItems entries: {0} (none naming this file)" -f $items.Count)
        }
        $fix = "run Clear-ExcelResiliency (or Remove-Item '" + (Join-Path (Split-Path -Parent $SecurityKey) 'Resiliency\DisabledItems') + "' -Recurse) and RE-RUN; this run's verdict is void."
    }
    # --- rung 3: the native log ----------------------------------------------
    elseif (-not $logExists -or $logLen -eq 0) {
        $rung  = 'NoNativeLog'
        $cause = "the add-in was NEVER LOADED: the DLL entry point never wrote its native log, so nothing in the add-in ran. Signature policy and the disabled list are both clear, so the failure is upstream of them."
        $ev    = $logNote
        $fix   = ("confirm the XLL exists at '{0}', that its BITNESS matches this Excel (a 32-bit Excel cannot load a 64-bit XLL and says nothing), that the file is not mark-of-the-web blocked (Unblock-File '{0}'), and that the launch actually passed it to Excel." -f $full)
    }
    elseif ($logStale) {
        $rung  = 'StaleNativeLog'
        $conf  = 'low'
        $cause = ("the add-in probably did NOT load in THIS run: the native log's last write ({0:yyyy-MM-dd HH:mm:ss}) predates this run's launch ({1:yyyy-MM-dd HH:mm:ss}), so its contents are left over from an earlier run. Note the other reading: Windows can delay a timestamp update while the writer still holds the file open, so read the log TAIL before concluding." -f $logWrite, $Since)
        $ev    = $logNote
        $fix   = ("read the tail of '{0}'; if it does not cover this run, treat the add-in as not loaded and re-run after checking bitness and the launch command line." -f $NativeLog)
    }
    # --- no rung fired: the add-in DID load ----------------------------------
    else {
        $rung  = 'Loaded'
        $cause = "the add-in DID load (the entry point wrote the native log), so this failure is NOT a load failure -- it belongs to the product. Record it as a defect and read the log tail."
        $ev    = $logNote
        $fix   = ("read the tail of '{0}' and its Go-side sibling for the failing call." -f $NativeLog)
    }

    # The native-log line is always present EXCEPT where it would merely repeat the
    # evidence line: the rungs that fired ON the log already say it.
    $msg = ("XLL-LOAD DIAGNOSIS [rung={0} confidence={1}]: {2}" -f $rung, $conf, $cause) + [Environment]::NewLine +
           ("  evidence   : {0}" -f $ev) + [Environment]::NewLine +
           $(if ($ev -ne $logNote) { ("  native log : {0}" -f $logNote) + [Environment]::NewLine } else { "" }) +
           ("  next step  : {0}" -f $fix)

    [pscustomobject]@{
        Rung        = $rung
        Loaded      = ($rung -eq 'Loaded')
        Confidence  = $conf
        Cause       = $cause
        Evidence    = $ev
        Remediation = $fix
        NativeLog   = $NativeLog
        LogNote     = $logNote
        Message     = $msg
    }
}

# Prints the ladder's verdict. This is what the drivers call at the point where
# the add-in should be loaded by now and evidently is not; it never throws, so
# the caller's teardown still runs (see the note above).
function Write-XllLoadDiagnosis {
    param(
        [Parameter(Mandatory)][string]$XllPath,
        [string]$NativeLog,
        [datetime]$Since,
        [string]$SecurityKey = $ExcelSecurityKey
    )
    $a = @{ XllPath = $XllPath; SecurityKey = $SecurityKey }
    if ($NativeLog) { $a['NativeLog'] = $NativeLog }
    if ($PSBoundParameters.ContainsKey('Since')) { $a['Since'] = $Since }
    $d = $null
    try { $d = Get-XllLoadFailureDiagnosis @a } catch { Write-Output ("XLL-LOAD DIAGNOSIS unavailable: {0}" -f $_); return }
    Write-Output $d.Message
}

# --- UIA bootstrap ----------------------------------------------------------
# Loads the UIA assemblies and publishes $AE / $TS into the CALLER's scope
# (script scope) so callers can use them exactly as before. Idempotent.
function Initialize-Uia {
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
    Set-Variable -Name AE -Scope Script -Value ([Windows.Automation.AutomationElement])
    Set-Variable -Name TS -Scope Script -Value ([Windows.Automation.TreeScope])
}

# --- Restart Manager: which process(es) hold a lock on $path ----------------
Add-Type -Namespace RM -Name Locks -MemberDefinition @'
[StructLayout(LayoutKind.Sequential)]
struct RM_UNIQUE_PROCESS { public int dwProcessId; public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime; }
[StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
struct RM_PROCESS_INFO {
  public RM_UNIQUE_PROCESS Process;
  [MarshalAs(UnmanagedType.ByValTStr, SizeConst=256)] public string strAppName;
  [MarshalAs(UnmanagedType.ByValTStr, SizeConst=64)]  public string strServiceShortName;
  public int ApplicationType; public uint AppStatus; public uint TSSessionId; [MarshalAs(UnmanagedType.Bool)] public bool bRestartable;
}
[DllImport("rstrtmgr.dll", CharSet=CharSet.Unicode)] static extern int RmStartSession(out uint h, int flags, string key);
[DllImport("rstrtmgr.dll")] static extern int RmEndSession(uint h);
[DllImport("rstrtmgr.dll", CharSet=CharSet.Unicode)] static extern int RmRegisterResources(uint h, uint nf, string[] f, uint na, IntPtr a, uint ns, string[] s);
[DllImport("rstrtmgr.dll")] static extern int RmGetList(uint h, out uint needed, ref uint count, [In,Out] RM_PROCESS_INFO[] info, ref uint reasons);
public static string Holders(string path) {
  uint h; var key = Guid.NewGuid().ToString("N");
  if (RmStartSession(out h, 0, key) != 0) return "(RM start failed)";
  try {
    if (RmRegisterResources(h, 1, new[]{path}, 0, IntPtr.Zero, 0, null) != 0) return "(RM register failed)";
    uint needed = 0, count = 0, reasons = 0;
    int rc = RmGetList(h, out needed, ref count, null, ref reasons);
    if (needed == 0) return "(none)";
    var arr = new RM_PROCESS_INFO[needed]; count = needed;
    rc = RmGetList(h, out needed, ref count, arr, ref reasons);
    if (rc != 0) return "(RM getlist rc=" + rc + ")";
    var sb = new System.Text.StringBuilder();
    for (int i=0;i<count;i++) sb.Append(arr[i].strAppName + " (pid=" + arr[i].Process.dwProcessId + ")  ");
    return sb.ToString().Trim();
  } finally { RmEndSession(h); }
}
'@

function Probe-Lock([string]$path) {
    if (-not (Test-Path $path)) { return "MISSING" }
    try {
        $fs = [System.IO.File]::Open($path, 'Open', 'ReadWrite', 'None')  # exclusive
        $fs.Close(); return "FREE (deletable)"
    } catch {
        return "LOCKED by -> " + [RM.Locks]::Holders($path)
    }
}

# --- visible Excel window enumeration ---------------------------------------
# Returns a (possibly empty) [System.Collections.ArrayList] of AutomationElement
# top-level Excel windows. Using an explicit ArrayList (not PowerShell array
# unrolling) keeps the return type stable whether 0, 1, or N windows match.
#
# Identifies the transient "Opening..." splash by ASCII-only means (the file is
# ASCII-only -- see ENCODING note at top): the ready workbook window's title
# contains the workbook file extension '.xlsx', the splash does not.
#   -RequireWorkbook : match only the ready '<file>.xlsx - Excel' window
#                      (acquisition); excludes the splash + the Start screen.
#   (default)        : match ANY 'Excel' window EXCEPT the ASCII 'Opening'
#                      splash (post-close ghost detection -- a lingering
#                      windowless ghost has NO window, so any window counts).
function Get-ExcelWindows {
    param([switch]$RequireWorkbook)
    $cond = New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [Windows.Automation.ControlType]::Window)
    $all = $AE::RootElement.FindAll($TS::Children, $cond)
    $out = New-Object System.Collections.ArrayList
    foreach ($w in $all) {
        try {
            $nm = $w.Current.Name
            if ($nm -notmatch 'Excel') { continue }
            if ($RequireWorkbook) {
                if ($nm -match '\.xlsx') { [void]$out.Add($w) }   # ready workbook window only
            } else {
                if ($nm -notmatch 'Opening') { [void]$out.Add($w) }  # any Excel window, skip ASCII splash
            }
        } catch {}
    }
    return ,$out
}

# --- modal dialog detection -------------------------------------------------
# A driver that polls for a cell value or a window while Excel is sitting on a
# MODAL dialog polls until timeout and then reports the thing it was waiting
# for as broken. That misread has a name in this repo's history: the RTD
# external-data warning (DataConnectionWarnings != 0) put up a modal, the
# driver timed out, and the run was written up as "RTD is wedged". The Trust
# Center gate now prevents THAT particular modal, but a modal can appear for
# reasons the gate cannot enumerate (a repair prompt, a license notice, an
# unrelated add-in), so the polling loops need to be able to tell the two
# apart. "Excel is asking a question" and "the feature under test is broken"
# must not produce the same verdict.
#
# DETECTION PREDICATE, and why it is not just IsModal. The obvious test is
# WindowPattern.Current.IsModal -- the property Windows itself sets, needing no
# list of dialog titles. It was tried first and MEASURED WRONG (2026-08-03, this
# machine, via tools\test-uia-common.ps1):
#
#   a dialog with NO owner window   -> class '#32770', IsModal = FALSE
#   a dialog WITH an owner window   -> class '#32770', IsModal = TRUE
#
# So IsModal alone misses the unowned case entirely. The Win32 dialog CLASS
# '#32770' was true in both, and it is not localized, so the predicate is
#
#     ClassName -eq '#32770'  OR  WindowPattern.IsModal
#
# ...keeping IsModal for any dialog Office draws with a class of its own.
#
# SCOPE also matters and was also measured: an OWNED dialog is not a child of the
# desktop, it is a DESCENDANT of its owner frame, so a Children-only sweep of the
# desktop finds nothing at all in the Excel case. Each candidate frame is
# therefore searched to Descendants, and the frame itself is tested too (that is
# where an unowned dialog turns up). The sweep is rooted per-process rather than
# running Descendants from RootElement, which would walk the whole desktop tree.
#
# KNOWN over-trigger, accepted: a MODELESS '#32770' (Excel's Find dialog, say)
# also reports. None of these drivers opens one, and a stray dialog stealing
# focus would break them anyway, so failing is the right answer either way.
#
# Returns $null when no dialog is up, otherwise a pscustomobject with the dialog
# Name, its visible text (joined), its ProcessId, and which signal fired.
function Get-ExcelModalDialog {
    param(
        [int]$ProcessId = 0,       # 0 = any process that looks like Excel
        [switch]$AnyProcess        # test hook: do not filter by name/pid at all
    )
    # SELF-BOOTSTRAP. This must NOT depend on the caller having run Initialize-Uia:
    # two of the six drivers (verify-gridonce-error, verify-ydp-stranding) are pure
    # COM and never call it, yet they poll through Wait-Settled and so reach here.
    # Relying on $AE cost a real failure -- $AE::ControlTypeProperty on a null $AE
    # throws "Value cannot be null (Parameter 'property')" from inside the poll
    # loop, which surfaces as a DRIVER ERROR in the middle of a product test.
    # Add-Type is idempotent, so paying it here is free after the first call.
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
    $ae = [Windows.Automation.AutomationElement]
    $ts = [Windows.Automation.TreeScope]
    $winCond = New-Object Windows.Automation.PropertyCondition($ae::ControlTypeProperty, [Windows.Automation.ControlType]::Window)

    # 1. the frames worth searching
    $frames = New-Object System.Collections.ArrayList
    foreach ($w in $ae::RootElement.FindAll($ts::Children, $winCond)) {
        try {
            $keep = $false
            if ($AnyProcess) {
                $keep = $true
            } elseif ($ProcessId -gt 0) {
                $keep = ($w.Current.ProcessId -eq $ProcessId)
            } else {
                # A dialog's Name is its own title, not 'Excel', so fall back to
                # the owning process actually being an Excel process.
                $keep = ($w.Current.Name -match 'Excel')
                if (-not $keep) {
                    try { $keep = ((Get-Process -Id $w.Current.ProcessId -EA Stop).ProcessName -eq 'EXCEL') } catch {}
                }
            }
            if ($keep) { [void]$frames.Add($w) }
        } catch {}
    }

    # 2. the frame itself, then every Window descendant of it
    foreach ($f in $frames) {
        $cands = New-Object System.Collections.ArrayList
        [void]$cands.Add($f)
        try { foreach ($d in $f.FindAll($ts::Descendants, $winCond)) { [void]$cands.Add($d) } } catch {}

        foreach ($w in $cands) {
            try {
                $cls = ''
                try { $cls = $w.Current.ClassName } catch {}
                $isDlgClass = ($cls -eq '#32770')
                $isModal = $false
                try { $isModal = $w.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern).Current.IsModal } catch {}
                if (-not ($isDlgClass -or $isModal)) { continue }

                # The MESSAGE, which is the whole reason a caller can act on the
                # throw. Collected by WIN32 CLASS ('Static' = a label), not by
                # ControlType: a MessageBox's message line was measured to be
                # ControlType.Pane, so filtering on ControlType.Text -- the
                # obvious choice, and the first one tried -- returned nothing and
                # every thrown message came out with an empty text field.
                # Button labels are skipped: they are localized and say nothing
                # about what is being asked.
                $msg = New-Object System.Collections.ArrayList
                try {
                    foreach ($t in $w.FindAll($ts::Descendants, [Windows.Automation.Condition]::TrueCondition)) {
                        try {
                            $tc = ''
                            try { $tc = $t.Current.ClassName } catch {}
                            $isLabel = ($tc -eq 'Static') -or
                                       ($t.Current.ControlType -eq [Windows.Automation.ControlType]::Text)
                            if (-not $isLabel) { continue }
                            $n = $t.Current.Name
                            if ($n) { [void]$msg.Add($n) }
                        } catch {}
                    }
                } catch {}
                return [pscustomobject]@{
                    Name      = $w.Current.Name
                    Text      = ($msg -join ' | ')
                    ProcessId = $w.Current.ProcessId
                    Signal    = $(if ($isDlgClass -and $isModal) { 'class+modal' } elseif ($isDlgClass) { 'class' } else { 'modal' })
                }
            } catch {}
        }
    }
    return $null
}

# Throws when a modal dialog is blocking Excel. Call this from inside polling
# loops so a modal FAILS THE RUN with the dialog's own text instead of being
# laundered into a timeout on whatever the loop was sampling.
#
# NOT called from Close-ExcelWindowFaithful: that function's entire job is to
# put up and dismiss the "Save changes?" prompt, so a modal there is the
# EXPECTED state, not an error. Keep it that way -- wiring this into the close
# path would make every faithful close fail.
function Assert-NoExcelModal {
    param([int]$ProcessId = 0, [string]$Context = 'polling')
    $dlg = Get-ExcelModalDialog -ProcessId $ProcessId
    if ($dlg) {
        throw ("Excel is blocked on a MODAL dialog during {0}: '{1}'{2}. " -f
                   $Context, $dlg.Name, ($(if ($dlg.Text) { " -- text: '" + $dlg.Text + "'" } else { "" }))) +
              "This is an ENVIRONMENT/interaction problem, not a product defect: a driver that keeps " +
              "polling here would time out and the timeout would be misread as the feature under test " +
              "being wedged. Dismiss the dialog, remove its cause, then re-run."
    }
}

# --- stale / final process kill (tier-2 of the two-tier teardown for the
#     UIA-driven scripts that never hold a COM client) -----------------------
# Kills EXCEL + the showcase Go server (and the legacy 'go_server' name when
# -IncludeGoServer is set). Used both as the up-front stale-process sweep and
# as the final force-kill tier.
function Stop-ShowcaseProcesses {
    param([switch]$IncludeGoServer)
    $names = @('EXCEL', 'xll_showcase')
    if ($IncludeGoServer) { $names += 'go_server' }
    Get-Process $names -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
}

# --- blank-workbook prep ----------------------------------------------------
# A blank workbook is required: launching ONLY the .xll shows the Start screen,
# from which the ribbon/Build flow is not reachable. We mint it with a
# SHORT-LIVED COM client that is fully released + force-killed before the real
# (COM-free) Excel launch -- a held COM client would mask the very bugs these
# repros chase. Returns the temp .xlsx path.
function New-ShowcaseWorkbook {
    param([Parameter(Mandatory)][string]$Name)
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) $Name
    $c = New-Object -ComObject Excel.Application; $c.DisplayAlerts = $false
    $b = $c.Workbooks.Add(); $b.SaveAs($tmp, 51); $b.Close($false); $c.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($c)
    Start-Sleep 2; Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue; Start-Sleep 1
    return $tmp
}

# --- launch real Excel (no held COM object) ---------------------------------
# Opens the workbook + loads the XLL addin + ribbon, exactly as a user would by
# double-clicking. Returns the Process object (caller reads .Id).
function Start-ShowcaseExcel {
    param([Parameter(Mandatory)][string]$Workbook, [Parameter(Mandatory)][string]$Xll)
    return Start-Process $xlexe -ArgumentList "`"$Workbook`"", "`"$Xll`"" -PassThru
}

# --- wait for the ready Excel window ----------------------------------------
# Polls UIA for the ready workbook window. EXCEL.EXE can relaunch off a pid
# different from Start-Process's, so we match on the workbook title via
# Get-ExcelWindows -RequireWorkbook and report the window-owning pid back
# through [ref]$ResolvedPid (seeded with the launched pid as a fallback).
#   -SettleFirst <s> : sleep before the first poll (the COM-bounce variant
#                      makes Excel slow to materialize, and polling UIA too
#                      eagerly during init can transiently throw on Current.Name).
# Returns the AutomationElement window, or $null on timeout.
function Wait-ShowcaseWindow {
    param(
        [int]$LaunchedPid = 0,
        [ref]$ResolvedPid,
        [int]$SettleFirst = 0,
        [int]$Tries = 38,
        [int]$IntervalMs = 800
    )
    if ($SettleFirst -gt 0) { Start-Sleep $SettleFirst }
    if ($ResolvedPid) { $ResolvedPid.Value = $LaunchedPid }
    for ($i = 0; $i -lt $Tries; $i++) {
        $cands = Get-ExcelWindows -RequireWorkbook
        if ($cands.Count -gt 0) {
            $w = $cands[0]
            if ($ResolvedPid) { try { $ResolvedPid.Value = $w.Current.ProcessId } catch {} }
            return $w
        }
        # A modal dialog here (repair prompt, recovery pane, license notice) means
        # the ready workbook window will NEVER appear, so the remaining tries are
        # dead time that ends in "FAIL: no ready window" -- a verdict that points
        # at the add-in instead of at the dialog actually holding Excel.
        Assert-NoExcelModal -Context 'waiting for the ready workbook window'
        Start-Sleep -Milliseconds $IntervalMs
    }
    return $null
}

# --- select showcase ribbon tab + invoke a button ---------------------------
# Selecting the custom ribbon tab is what brings its buttons into the UIA tree;
# only then can the button be Invoked. Returns $true if a matching button was
# invoked. $NamePattern is matched (regex) against button names.
function Invoke-RibbonButton {
    param(
        [Parameter(Mandatory)]$Window,
        [Parameter(Mandatory)][string]$NamePattern,
        [string]$Tab = 'xll-gen Showcase'
    )
    $tabCond = New-Object Windows.Automation.AndCondition(@(
        (New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [Windows.Automation.ControlType]::TabItem)),
        (New-Object Windows.Automation.PropertyCondition($AE::NameProperty, $Tab)) ))
    $tabEl = $Window.FindFirst($TS::Descendants, $tabCond)
    if ($tabEl) { try { $tabEl.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select(); Start-Sleep 1 } catch {} }
    $btnCond = New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [Windows.Automation.ControlType]::Button)
    foreach ($bn in $Window.FindAll($TS::Descendants, $btnCond)) {
        try {
            if ($bn.Current.Name -match $NamePattern) {
                $bn.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
                return $true
            }
        } catch {}
    }
    return $false
}

# --- FAITHFUL window close --------------------------------------------------
# Drives the REAL window close via UIA WindowPattern.Close() (a user clicking X)
# and dismisses the "Save changes? -> Don't Save" prompt -- NOT a COM
# Application.Quit (which would mask the close-time bugs by tearing RTD down on
# a different code path). The Korean "Don't Save" button is matched by codepoint
# (U+C800 U+C7A5), never as a source literal (see ENCODING note at top); a
# keyboard 'n' / Alt+F4 escalation also dismisses lingering windows.
#   -Escalate : if a top-level Excel window still lingers after Close()+Don't-Save,
#               escalate to a FAITHFUL Alt+F4 on the foreground Excel window
#               (still a user "close window" gesture, NOT COM Quit). Some
#               environments leave the app frame running after the workbook
#               window closes (no quit -> no teardown -> inconclusive run).
function Close-ExcelWindowFaithful {
    param([Parameter(Mandatory)]$Window, [switch]$Escalate)
    $btnCond = New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [Windows.Automation.ControlType]::Button)
    $dontSaveKR = [string][char]0xC800 + [string][char]0xC7A5  # localized "Don't Save" prefix

    Write-Output "--- closing window via UIA (faithful WindowPattern.Close) ---"
    try { $Window.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern).Close() } catch { Write-Output "close pattern failed: $_" }

    # Dismiss the "Save changes? -> Don't Save" prompt.
    for ($i = 0; $i -lt 12; $i++) {
        Start-Sleep -Milliseconds 500
        $dlgBtn = $null
        foreach ($w in (Get-ExcelWindows)) {
            foreach ($bn in $w.FindAll($TS::Descendants, $btnCond)) {
                try {
                    $bnm = $bn.Current.Name
                    if (($bnm -match 'Do.{0,2}n.?t Save') -or ($bnm -like ('*'+$dontSaveKR+'*'))) { $dlgBtn = $bn; break }
                } catch {}
            }
            if ($dlgBtn) { break }
        }
        if ($dlgBtn) { try { $dlgBtn.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke(); Write-Output "clicked Don't-Save" } catch {}; break }
    }
    # Keyboard fallback: if a save dialog is still up, press 'N' (Don't Save accelerator).
    foreach ($w in (Get-ExcelWindows)) {
        $hasDialog = $false
        foreach ($bn in $w.FindAll($TS::Descendants, $btnCond)) {
            try { if ($bn.Current.Name -match 'Save|Cancel') { $hasDialog = $true; break } } catch {}
        }
        if ($hasDialog) {
            try { $shell = New-Object -ComObject WScript.Shell; $shell.SendKeys('n'); Write-Output "sent 'n' (Don't-Save accelerator)" } catch {}
            break
        }
    }

    if (-not $Escalate) { return }

    # Faithful-close fallback. In some environments WindowPattern.Close() on the
    # workbook window dismisses the document but leaves the Excel application
    # frame running (no quit -> no OnBeginShutdown/OnDisconnection -> teardown
    # never fires and the run is inconclusive). If a top-level Excel window is
    # still present a moment later, escalate to Alt+F4 on the foreground Excel
    # window. Alt+F4 is a FAITHFUL user "close the window" gesture (NOT COM
    # Application.Quit, which would mask the S1' bug). We only do this when a
    # window lingers, so genuine clean closes are unaffected.
    Start-Sleep -Milliseconds 800
    $lingering = Get-ExcelWindows
    if ($lingering.Count -gt 0) {
        Write-Output "--- window still present after Close(); escalating to faithful Alt+F4 ---"
        for ($k = 0; $k -lt 5; $k++) {
            $cur = Get-ExcelWindows
            if ($cur.Count -eq 0) { break }
            try {
                $w = $cur[0]
                try { $w.SetFocus() } catch {}
                Start-Sleep -Milliseconds 200
                $shell = New-Object -ComObject WScript.Shell
                $shell.SendKeys('%{F4}')
                Write-Output "sent Alt+F4"
            } catch { Write-Output "Alt+F4 failed: $_" }
            Start-Sleep -Milliseconds 600
            foreach ($w in (Get-ExcelWindows)) {
                foreach ($bn in $w.FindAll($TS::Descendants, $btnCond)) {
                    try {
                        $bnm = $bn.Current.Name
                        if (($bnm -match 'Do.{0,2}n.?t Save') -or ($bnm -like ('*'+$dontSaveKR+'*'))) {
                            $bn.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
                            Write-Output "clicked Don't-Save (post Alt+F4)"
                        }
                    } catch {}
                }
            }
            Start-Sleep -Milliseconds 600
        }
    }
}

# --- TWO-TIER teardown for COM-driven scripts -------------------------------
# PROJECT-MANDATORY two-tier cleanup: FIRST attempt a graceful exit (close the
# workbook + Application.Quit + release the COM RCWs so Excel can shut down on
# its own code path), THEN force-kill any survivor. Never replace this with a
# bare defer or a force-kill alone -- the graceful tier exercises the real
# teardown path the bugs live on; the force-kill tier only guarantees the test
# host is clean afterward.
#   -App / -Workbook / -Worksheet : the COM RCWs to release (any may be $null).
#   -IncludeGoServer              : also kill the legacy 'go_server' process.
function Stop-ShowcaseCom {
    param($App, $Workbook, $Worksheet, [switch]$IncludeGoServer)
    # tier 1: graceful
    if ($Workbook)  { try { $Workbook.Close($false) } catch {} }
    if ($App)       { try { $App.Quit() } catch {} }
    if ($Worksheet) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Worksheet) } catch {} }
    if ($Workbook)  { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Workbook) } catch {} }
    if ($App)       { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($App) } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep 2
    # tier 2: force-kill any survivor
    Stop-ShowcaseProcesses -IncludeGoServer:$IncludeGoServer
}

# --- entering a formula the way a USER would --------------------------------
# WHY THIS IS SHARED. On a dynamic-array Excel, writing a grid formula through
# .Formula applies IMPLICIT INTERSECTION -- an invisible leading @ -- so
# '=YDH("MSFT",30)' silently becomes '=@YDH("MSFT",30)' and the cell takes the
# SINGLE-CELL path. A script that does that while claiming to test the grid
# spill path is testing something else and cannot fail for the reason it says.
# verify-ydp-stranding-uia.ps1 did exactly that until 2026-08-02.
#
# The probe MUST itself write through .Formula2, and that is the whole subtlety:
# probing with .Formula makes '=SEQUENCE(2,2)' paint a lone 1, which is
# indistinguishable at a glance from a pre-DA Excel that cannot spill at all --
# and reading that as "no dynamic arrays" then hides whether OUR grid spills.
# (The tell that it is not really pre-DA: a pre-DA Excel has no SEQUENCE and
# answers #NAME?.) .Formula2 does not exist before the DA engine, so a throw is
# the honest pre-DA signal.
#
# Publishes $script:hasDA into the CALLER's scope, the way Initialize-Uia
# publishes $AE/$TS. Returns $true when dynamic arrays are available.
function Test-DynamicArrays {
    param([Parameter(Mandatory)]$Worksheet)
    $da = $false
    try {
        $Worksheet.Range("Z1").Formula2 = '=SEQUENCE(2,2)'
        Start-Sleep 1
        $da = ($null -ne $Worksheet.Range("AA1").Value2)
    } catch {
        Write-Output "  (.Formula2 unavailable: $($_.Exception.Message))"
    }
    try { $Worksheet.Range("Z1:AA2").ClearContents() } catch {}
    Set-Variable -Name hasDA -Scope Script -Value $da
    # Deliberately returns NOTHING. Write-Output shares the pipeline with the
    # return value, so a caller writing [void](Test-DynamicArrays ...) to discard
    # a bool would swallow this diagnostic line with it -- and the diagnostic is
    # the only way to see which path the grid actually took. Callers read $hasDA.
    Write-Output ("dynamic arrays: " + $(if ($da) { 'available - entering grids via .Formula2' } else { 'NOT available (pre-DA build) - grids need legacy CSE' }))
}

# Enters a formula through the API that matches this Excel. Call
# Test-DynamicArrays once first.
function Set-Formula {
    param([Parameter(Mandatory)]$Worksheet, [Parameter(Mandatory)][string]$Address, [Parameter(Mandatory)][string]$Formula)
    if ($script:hasDA) { $Worksheet.Range($Address).Formula2 = $Formula }
    else { $Worksheet.Range($Address).Formula = $Formula }
}
