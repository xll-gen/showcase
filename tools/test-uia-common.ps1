# test-uia-common.ps1 -- self-test for the parts of uia-common.ps1 that can be
# checked WITHOUT Excel. Run it after touching the shared scaffold:
#
#     pwsh -NoProfile -File tools\test-uia-common.ps1
#
# WHY THIS FILE EXISTS. The showcase harness is the only thing standing between
# "the product is broken" and "the environment is broken", and it has been wrong
# in that direction more than once: a regex comment stripper that silently ate
# assertions, a grid entered through .Formula so the spill path was never
# exercised, a dialog scan whose localized alternatives could not match because
# of the source encoding. Every one of those FAILED SILENTLY -- the run stayed
# green while covering nothing. A harness needs its own negative controls.
#
# ENCODING: ASCII-only, same rule as uia-common.ps1.

. "$PSScriptRoot\uia-common.ps1"

$fails = 0
$checks = 0
function Check([bool]$ok, [string]$what) {
    $script:checks++
    if ($ok) { Write-Output "  ok    $what" } else { Write-Output "  FAIL  $what"; $script:fails++ }
}

Write-Output "=== uia-common.ps1 self-test ==="

# ---------------------------------------------------------------------------
# 1. no non-ASCII inside a MATCHING operand
# ---------------------------------------------------------------------------
# uia-common.ps1 documents the mechanism: a .ps1 with no UTF-8 BOM is read under
# the system ANSI codepage, so a non-ASCII literal is mangled -- and when the
# mangled literal is a REGEX, the resulting exception gets swallowed by a
# surrounding try/catch and the match simply never happens. Not hypothetical:
# verify-rtd-stream-uia.ps1 carried Korean alternatives in its dialog-detection
# regex until 2026-08-03, so the localized half of the pattern was dead on
# exactly the machines it was added for, silently.
#
# The check is deliberately narrowed to lines using a matching operator rather
# than "any non-ASCII in code". A first draft banned all of it and flagged eight
# em-dashes inside Write-Output MESSAGES -- noise that says nothing about
# correctness, and noise is how a real finding gets scrolled past. Mangled prose
# in an output string is ugly; a mangled regex changes behavior. Only the second
# is a test.
Write-Output "[1] no non-ASCII inside -match/-like/-replace patterns (mangled regex fails silently)"
$matchOps = '-i?(match|notmatch|like|notlike|replace|csplit|split)\b'
$flagged = @()
foreach ($f in (Get-ChildItem "$PSScriptRoot\*.ps1")) {
    $lineNo = 0
    foreach ($line in (Get-Content $f.FullName -Encoding UTF8)) {
        $lineNo++
        $code = ($line -replace '#.*$', '')          # drop trailing/whole-line comments
        if ($code -notmatch $matchOps) { continue }
        foreach ($ch in $code.ToCharArray()) {
            if ([int]$ch -gt 127) {
                $flagged += ("{0}:{1} U+{2:X4}" -f $f.Name, $lineNo, [int]$ch)
                break
            }
        }
    }
}
Check ($flagged.Count -eq 0) ("no non-ASCII in matching operands" +
    $(if ($flagged.Count) { " -- found: " + ($flagged -join ', ') } else { "" }))

# ---------------------------------------------------------------------------
# 2. every driver clears resiliency residue before launching Excel
# ---------------------------------------------------------------------------
# A previous crashed run can leave the add-in on Excel's DisabledItems list. The
# next run then measures an add-in that never loaded, and every verdict in it is
# about nothing.
Write-Output "[2] every driver calls Clear-ExcelResiliency after the trust gate"
$drivers = @('diagnose-close-uia.ps1','ghost-check.ps1','repro-crash-uia.ps1',
             'verify-gridonce-error-uia.ps1','verify-rtd-stream-uia.ps1',
             'verify-ydp-stranding-uia.ps1')
foreach ($d in $drivers) {
    $src = Get-Content (Join-Path $PSScriptRoot $d) -Raw
    $gate = $src.IndexOf('Assert-ExcelTrustPreconditions')
    $clr  = $src.IndexOf('Clear-ExcelResiliency')
    Check (($gate -ge 0) -and ($clr -gt $gate)) ("$d gates then clears")
    # DocumentRecovery must NOT be auto-dropped: it can hold a real user's
    # unsaved workbooks. The gate WARNs; a human decides.
    Check ($src -notmatch 'Clear-ExcelResiliency\s+-IncludeDocumentRecovery') `
          ("$d does not auto-drop DocumentRecovery")
}

# ---------------------------------------------------------------------------
# 3. the modal detector, against a REAL modal dialog
# ---------------------------------------------------------------------------
# A detector that never fires leaves every polling loop exactly as it was, so it
# has to be shown BOTH firing and staying quiet, against a REAL dialog.
#
# BOTH SHAPES ARE EXERCISED, because they behave differently and the difference
# is what the detector got wrong on the first attempt (measured here):
#
#   unowned dialog -> class '#32770', IsModal FALSE, sits as a desktop child
#   owned dialog   -> class '#32770', IsModal TRUE,  sits as a DESCENDANT of its
#                     owner frame, NOT as a desktop child
#
# The first draft tested IsModal only and swept desktop Children only, so it
# missed the unowned case on the property and the owned case on the scope -- i.e.
# it would have found NOTHING in the Excel case while every wiring assertion in
# section [4] still passed. Excel's own dialogs are the owned shape.
#
# The child process is launched with -EncodedCommand, not -Command. With -Command
# the quotes inside the payload were eaten in transit, the child died on a parse
# error, no window ever appeared, and the detector "correctly" reported nothing --
# a passing-looking harness bug stacked on top of the real one. Base64 has no
# quoting to lose.
Write-Output "[3] modal detector fires on a real dialog (owned AND unowned) and only then"
Initialize-Uia

function Start-ModalProbe([bool]$Owned) {
    $body = if ($Owned) {
@'
Add-Type -AssemblyName System.Windows.Forms
$f = New-Object System.Windows.Forms.Form
$f.Text = "XllGenOwnerProbe"; $f.Show(); $f.Refresh()
[void][System.Windows.Forms.MessageBox]::Show($f, "blocking", "XllGenModalProbe")
'@
    } else {
@'
Add-Type -AssemblyName System.Windows.Forms
[void][System.Windows.Forms.MessageBox]::Show("blocking", "XllGenModalProbe")
'@
    }
    $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($body))
    return Start-Process pwsh -PassThru -ArgumentList @('-NoProfile','-STA','-EncodedCommand',$enc)
}

$dlg = Get-ExcelModalDialog -AnyProcess
Check ($null -eq $dlg) ("quiet when no dialog is up" + $(if ($dlg) { " (got '" + $dlg.Name + "')" } else { "" }))

foreach ($owned in @($false, $true)) {
    $label = $(if ($owned) { 'owned' } else { 'unowned' })
    $child = Start-ModalProbe $owned
    try {
        Start-Sleep 4
        Check (-not $child.HasExited) "$label probe process is still alive (it really put up a window)"
        $dlg = Get-ExcelModalDialog -AnyProcess
        Check ($null -ne $dlg) "a real $label dialog IS detected"
        if ($dlg) {
            Write-Output ("        Name='{0}' Signal='{1}' Text='{2}' pid={3}" -f $dlg.Name, $dlg.Signal, $dlg.Text, $dlg.ProcessId)
            Check ($dlg.ProcessId -eq $child.Id) "the detected $label dialog belongs to the process that raised it"
            Check ($dlg.Text -match 'blocking') "the $label dialog's visible TEXT is captured (that is what makes the thrown message useful)"
        }

        # The process filter must EXCLUDE unrelated processes, or a driver would
        # fail on whatever dialog happened to be open elsewhere on the desktop.
        Check ($null -eq (Get-ExcelModalDialog -ProcessId 999999)) `
              "a dialog in another process is NOT reported under -ProcessId ($label)"

        $threw = $false; $msg = ''
        try { Assert-NoExcelModal -ProcessId $child.Id -Context 'self-test' } catch { $threw = $true; $msg = "$_" }
        Check $threw "Assert-NoExcelModal THROWS while the $label dialog is up"
        Check ($msg -match 'XllGenModalProbe') "the thrown message names the $label dialog"
        Check ($msg -match 'ENVIRONMENT') "the thrown message says environment, not product defect ($label)"
    } finally {
        $child | Stop-Process -Force -EA SilentlyContinue
    }
    Start-Sleep 2
    Check ($null -eq (Get-ExcelModalDialog -AnyProcess)) "detector goes quiet once the $label dialog is gone"
    $threw = $false
    try { Assert-NoExcelModal -ProcessId $child.Id -Context 'self-test' } catch { $threw = $true }
    Check (-not $threw) "Assert-NoExcelModal does NOT throw with no dialog up ($label)"
}

# ---------------------------------------------------------------------------
# 3b. the detector must not need Initialize-Uia
# ---------------------------------------------------------------------------
# THIS CASE COST A REAL FAILURE. Assert-NoExcelModal is wired into Wait-Settled in
# verify-gridonce-error-uia.ps1, but that driver (and verify-ydp-stranding) is pure
# COM and never calls Initialize-Uia. The detector originally read $AE, so
# $AE::ControlTypeProperty on a null $AE threw "Value cannot be null (Parameter
# 'property')" from inside the poll loop -- surfacing as a DRIVER ERROR in the middle
# of a product test, on check [3] of a run whose checks [1] and [2] had passed.
#
# Section [3] above ran Initialize-Uia first, so it could never have caught this. The
# fix is that Get-ExcelModalDialog self-bootstraps; this pins it by NULLING the
# published variables before calling in.
Write-Output "[3b] the detector works without Initialize-Uia (two drivers never call it)"
$savedAE = $AE
$savedTS = $TS
try {
    Set-Variable -Name AE -Scope Script -Value $null
    Set-Variable -Name TS -Scope Script -Value $null
    $threw = $false
    $err = ''
    try { $null = Get-ExcelModalDialog -AnyProcess } catch { $threw = $true; $err = "$_" }
    Check (-not $threw) ("Get-ExcelModalDialog does not require Initialize-Uia" +
        $(if ($threw) { " -- threw: $err" } else { "" }))
    $threw = $false
    try { Assert-NoExcelModal -Context 'self-test, uninitialized UIA' } catch { $threw = $true; $err = "$_" }
    # No modal is up, so this must return quietly rather than throwing a UIA error.
    Check (-not $threw) ("Assert-NoExcelModal does not require Initialize-Uia" +
        $(if ($threw) { " -- threw: $err" } else { "" }))
} finally {
    Set-Variable -Name AE -Scope Script -Value $savedAE
    Set-Variable -Name TS -Scope Script -Value $savedTS
}

# ---------------------------------------------------------------------------
# 4. the modal check is wired into the polling loops -- and NOT into the close
# ---------------------------------------------------------------------------
# Close-ExcelWindowFaithful's job is to raise and dismiss the "Save changes?"
# prompt, so a modal there is the EXPECTED state. Wiring the assert into it
# would make every faithful close throw.
Write-Output "[4] wiring: asserted in polling loops, absent from the close path"
$common = Get-Content (Join-Path $PSScriptRoot 'uia-common.ps1') -Raw
$waitFn  = [regex]::Match($common, '(?s)function Wait-ShowcaseWindow.*?\n\}').Value
$closeFn = [regex]::Match($common, '(?s)function Close-ExcelWindowFaithful.*?\n\}').Value
Check ($waitFn -match 'Assert-NoExcelModal') "Wait-ShowcaseWindow asserts inside its retry loop"
Check ($closeFn.Length -gt 0 -and $closeFn -notmatch 'Assert-NoExcelModal') `
      "Close-ExcelWindowFaithful does NOT assert (it expects the save prompt)"
$grid = Get-Content (Join-Path $PSScriptRoot 'verify-gridonce-error-uia.ps1') -Raw
$settle = [regex]::Match($grid, '(?s)function Wait-Settled.*?\n\}').Value
Check ($settle -match 'Assert-NoExcelModal') "Wait-Settled asserts inside its poll loop"
$stream = Get-Content (Join-Path $PSScriptRoot 'verify-rtd-stream-uia.ps1') -Raw
Check ($stream -match 'Get-ExcelModalDialog') "the RTD stream sampler uses the shared detector"
Check ($stream -notmatch "match 'responding\|") "the old title-regex dialog scan is gone"

# ---------------------------------------------------------------------------
# 5. the "was the add-in even loaded?" ladder, against REAL conditions
# ---------------------------------------------------------------------------
# THE LESSON THIS SECTION IS PAYING FOR. The previous detector in this file passed
# every wiring assertion while detecting nothing, because the assertions checked
# that the call existed, not that the call answered correctly. So this section
# stands the conditions UP -- a scratch registry subtree carrying a real
# RequireAddinSig value, a real Trusted Locations entry, real REG_BINARY
# DisabledItems blobs, and a real log file present / empty / missing / backdated
# -- and asserts the ladder picks the RIGHT rung in each, including the ORDER
# (rung 1 must mask rungs 2 and 3, rung 2 must mask rung 3) and the negative
# control (a healthy environment must NOT be diagnosed as a load failure).
#
# NOTHING here touches the machine's real Excel security settings. Every read is
# redirected with -SecurityKey to HKCU:\Software\XllGenSelfTest, the whole subtree
# is deleted in the finally block, and the real subtree is snapshotted before and
# after and compared -- because "the test cleaned up after itself" is exactly the
# kind of claim that is worth an assertion rather than a comment.
Write-Output "[5] load-failure ladder picks the right rung (scratch registry + scratch log)"

$scratchRoot = 'HKCU:\Software\XllGenSelfTest'
$scratchSec  = "$scratchRoot\Office\16.0\Excel\Security"
$scratchTL   = "$scratchSec\Trusted Locations\Location0"
$scratchDis  = "$scratchRoot\Office\16.0\Excel\Resiliency\DisabledItems"
$scratchDir  = Join-Path ([System.IO.Path]::GetTempPath()) ("xllgen-loadladder-" + [guid]::NewGuid().ToString('N'))
$scratchXll  = Join-Path $scratchDir 'xll_showcase.xll'
$scratchLog  = Join-Path $scratchDir 'xll_showcase_native.log'

function Get-RealSecuritySnapshot {
    $s = New-Object System.Collections.ArrayList
    foreach ($p in @($ExcelSecurityKey, (Join-Path (Split-Path -Parent $ExcelSecurityKey) 'Resiliency\DisabledItems'))) {
        if (Test-Path $p) {
            $k = Get-Item $p
            foreach ($n in @($k.GetValueNames() | Sort-Object)) { [void]$s.Add(("{0}|{1}={2}" -f $p, $n, (($k.GetValue($n)) -join ','))) }
            [void]$s.Add(("{0}|subkeys={1}" -f $p, ((@($k.GetSubKeyNames()) | Sort-Object) -join ',')))
        } else {
            [void]$s.Add(("{0}|ABSENT" -f $p))
        }
    }
    return ($s -join '#')
}

function Reset-ScratchRegistry {
    if (Test-Path $scratchRoot) { Remove-Item $scratchRoot -Recurse -Force }
    New-Item $scratchSec -Force | Out-Null
}

# The DisabledItems payload Excel writes is a binary blob with the offending path
# inside it. Both alignments are exercised: with an even-length header the path
# decodes as UTF-16, with an odd-length one it does not, and a decoder that only
# tried UTF-16 would silently stop naming the file.
function Add-ScratchDisabledItem([string]$name, [string]$path, [int]$headerBytes) {
    New-Item $scratchDis -Force | Out-Null
    $hdr = New-Object byte[] $headerBytes
    $blob = [byte[]]($hdr + [Text.Encoding]::Unicode.GetBytes($path))
    New-ItemProperty -Path $scratchDis -Name $name -PropertyType Binary -Value $blob -Force | Out-Null
}

# Excel also records disabled items as SUBKEYS of DisabledItems, not only as
# values. Get-ExcelDisabledItems read GetValueNames() alone until 2026-08-03, so
# that shape was invisible to the ladder.
function Add-ScratchDisabledSubkey([string]$name, [string]$path, [int]$headerBytes) {
    $sub = Join-Path $scratchDis $name
    New-Item $sub -Force | Out-Null
    if ($null -ne $path) {
        $hdr = New-Object byte[] $headerBytes
        $blob = [byte[]]($hdr + [Text.Encoding]::Unicode.GetBytes($path))
        New-ItemProperty -Path $sub -Name 'Item' -PropertyType Binary -Value $blob -Force | Out-Null
    }
}

function Set-ScratchLog([string]$text) { [System.IO.File]::WriteAllText($scratchLog, $text) }

function Invoke-Ladder { param([hashtable]$extra = @{})
    $a = @{ XllPath = $scratchXll; NativeLog = $scratchLog; SecurityKey = $scratchSec }
    foreach ($k in $extra.Keys) { $a[$k] = $extra[$k] }
    return Get-XllLoadFailureDiagnosis @a
}

$realBefore = Get-RealSecuritySnapshot
# A section that DIES halfway records no failure and the run still ends in "ALL
# CHECKS PASSED" -- which is what happened the first time this section was run
# (a null-valued $_.Text threw at the first ladder call and 24 checks never
# executed while the suite reported green). Hence both guards below: the catch,
# and the count assertion after the finally.
$checks5 = $checks
try {
    New-Item -ItemType Directory -Path $scratchDir -Force | Out-Null
    [System.IO.File]::WriteAllText($scratchXll, 'not a real xll, only its name matters here')

    # --- negative control: a healthy environment must read as LOADED ---------
    Reset-ScratchRegistry
    Set-ScratchLog "xlAutoOpen: showcase loaded`r`n"
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'Loaded' -and $d.Loaded) ("clean registry + non-empty native log -> Loaded (got rung=" + $d.Rung + ")")
    Check ($d.Message -match 'NOT a load failure') "the Loaded verdict says the failure belongs to the product"

    # --- rung 3: the native log ----------------------------------------------
    Remove-Item $scratchLog -Force
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'NoNativeLog' -and -not $d.Loaded) ("missing native log -> NoNativeLog (got rung=" + $d.Rung + ")")
    Check ($d.Remediation -match 'BITNESS') "the NoNativeLog remediation names the bitness trap"

    Set-ScratchLog ''
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'NoNativeLog') ("zero-length native log -> NoNativeLog (got rung=" + $d.Rung + ")")

    # --- rung 2: the disabled list, which must MASK a present log ------------
    Set-ScratchLog "xlAutoOpen: showcase loaded`r`n"
    Add-ScratchDisabledItem 'ITEM1' $scratchXll 4
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'DisabledItems' -and -not $d.Loaded) ("DisabledItems naming the xll beats a PRESENT log (got rung=" + $d.Rung + ")")
    Check ($d.Confidence -eq 'high') "a name-matching DisabledItems entry is high confidence"
    Check ($d.Message -match 'xll_showcase\.xll') "the DisabledItems verdict names the file it found"

    Reset-ScratchRegistry
    Add-ScratchDisabledItem 'ITEM2' $scratchXll 3    # odd header -> path is NOT UTF-16 aligned
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'DisabledItems' -and $d.Confidence -eq 'high') `
          ("an odd-aligned DisabledItems blob is still decoded and matched (got rung=" + $d.Rung + "/" + $d.Confidence + ")")

    Reset-ScratchRegistry
    Add-ScratchDisabledItem 'ITEM3' 'C:\somewhere\other_addin.xll' 4
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'DisabledItems' -and $d.Confidence -eq 'low') `
          ("a DisabledItems entry naming ANOTHER file is reported at LOW confidence (got rung=" + $d.Rung + "/" + $d.Confidence + ")")
    Check ($d.Message -match 'native log present') "the low-confidence verdict still reports the log state to weigh against it"

    # --- rung 2, shape B: the entry is a SUBKEY, not a value -----------------
    # Get-ExcelDisabledItems read GetValueNames() only, so this shape counted as
    # ZERO entries and the ladder fell through to rung=Loaded ("this belongs to
    # the product") for an add-in Excel had disabled -- while the precheck, which
    # counts ValueCount + SubKeyCount on the same key, had already warned.
    Reset-ScratchRegistry
    Set-ScratchLog "xlAutoOpen: showcase loaded`r`n"
    Add-ScratchDisabledSubkey 'SUB1' $scratchXll 4
    $items = Get-ExcelDisabledItems -SecurityKey $scratchSec
    Check ($items.Count -ge 1) ("a DisabledItems entry stored as a SUBKEY is enumerated (got " + $items.Count + ")")
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'DisabledItems' -and -not $d.Loaded) `
          ("a SUBKEY DisabledItems entry beats a PRESENT log (got rung=" + $d.Rung + ")")
    Check ($d.Confidence -eq 'high') "a subkey entry whose blob names the xll is high confidence"

    # A subkey holding no values at all must still count as an entry: it is
    # residue, and reporting zero would resurrect the same silent fall-through.
    Reset-ScratchRegistry
    Add-ScratchDisabledSubkey 'SUB2' $null 0
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'DisabledItems' -and $d.Confidence -eq 'low') `
          ("an EMPTY DisabledItems subkey is still an entry, at low confidence (got rung=" + $d.Rung + "/" + $d.Confidence + ")")

    # --- rung 0: the XLL file is not there at all ----------------------------
    # Probe P4: with a native log left by an earlier build, every rung fell
    # through and the ladder answered "Loaded (measured)" for a file that does
    # not exist.
    Reset-ScratchRegistry
    Set-ScratchLog "xlAutoOpen: showcase loaded`r`n"
    Remove-Item $scratchXll -Force
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'MissingXll' -and -not $d.Loaded) `
          ("a MISSING xll file is rung 0, not 'Loaded', even with a non-empty native log (got rung=" + $d.Rung + ")")
    Check ($d.Message -match 'XLL MISSING') "the missing-XLL verdict names the path it looked at"
    [System.IO.File]::WriteAllText($scratchXll, 'not a real xll, only its name matters here')
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'Loaded') ("control: putting the file back returns the ladder to Loaded (got rung=" + $d.Rung + ")")

    # --- rung 1: the signature requirement, which must mask rungs 2 and 3 ----
    Reset-ScratchRegistry
    Add-ScratchDisabledItem 'ITEM4' $scratchXll 4
    Set-ItemProperty $scratchSec RequireAddinSig 1 -Type DWord
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'RequireAddinSig' -and -not $d.Loaded) `
          ("RequireAddinSig=1 + untrusted folder masks DisabledItems AND a present log (got rung=" + $d.Rung + ")")
    Check ($d.Remediation -match 'Trusted Location') "the RequireAddinSig remediation is the trusted-location fix"

    # the trusted-location EXEMPTION is real and must be honoured: same sig=1, but
    # the folder is trusted, so rung 1 must not fire.
    Reset-ScratchRegistry
    Set-ItemProperty $scratchSec RequireAddinSig 1 -Type DWord
    New-Item $scratchTL -Force | Out-Null
    Set-ItemProperty $scratchTL Path $scratchDir
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'Loaded') ("RequireAddinSig=1 but the folder IS trusted -> no rung 1 (got rung=" + $d.Rung + ")")

    # AllowSubfolders decides the PARENT case, and getting it backwards would
    # either invent load failures or hide them.
    Reset-ScratchRegistry
    Set-ItemProperty $scratchSec RequireAddinSig 1 -Type DWord
    New-Item $scratchTL -Force | Out-Null
    Set-ItemProperty $scratchTL Path (Split-Path -Parent $scratchDir)
    Set-ItemProperty $scratchTL AllowSubfolders 0 -Type DWord
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'RequireAddinSig') ("parent trusted WITHOUT AllowSubfolders does not cover the folder (got rung=" + $d.Rung + ")")
    Set-ItemProperty $scratchTL AllowSubfolders 1 -Type DWord
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'Loaded') ("parent trusted WITH AllowSubfolders covers the folder (got rung=" + $d.Rung + ")")

    # --- -Since: a log from an EARLIER run must not count as proof -----------
    Reset-ScratchRegistry
    Set-ScratchLog "xlAutoOpen: showcase loaded`r`n"
    (Get-Item $scratchLog).LastWriteTime = (Get-Date).AddHours(-1)
    $d = Invoke-Ladder
    Check ($d.Rung -eq 'Loaded') ("with no -Since, an old log is taken at face value (got rung=" + $d.Rung + ")")
    $d = Invoke-Ladder @{ Since = (Get-Date) }
    Check ($d.Rung -eq 'StaleNativeLog' -and -not $d.Loaded) `
          ("a log written BEFORE this run's launch is not proof of loading (got rung=" + $d.Rung + ")")
    Check ($d.Confidence -eq 'low') "the stale-log verdict is low confidence (an open handle can hold the timestamp back)"
    $d = Invoke-Ladder @{ Since = (Get-Date).AddHours(-2) }
    Check ($d.Rung -eq 'Loaded') ("a log written AFTER this run's launch counts as loaded (got rung=" + $d.Rung + ")")

    # --- the decoder and the printer -----------------------------------------
    Reset-ScratchRegistry
    $items = Get-ExcelDisabledItems -SecurityKey $scratchSec
    Check ($items.Count -eq 0) ("no DisabledItems key -> empty list (got " + $items.Count + ")")
    Add-ScratchDisabledItem 'ITEM5' $scratchXll 4
    $items = Get-ExcelDisabledItems -SecurityKey $scratchSec
    Check ($items.Count -eq 1 -and $items[0].Text -match 'xll_showcase\.xll') "a REG_BINARY entry decodes to text carrying the path"

    Reset-ScratchRegistry
    $printed = (Write-XllLoadDiagnosis -XllPath $scratchXll -NativeLog $scratchLog -SecurityKey $scratchSec | Out-String)
    Check ($printed -match 'XLL-LOAD DIAGNOSIS') "Write-XllLoadDiagnosis prints the verdict"
    Check ($printed -match 'next step') "the printed verdict carries a remediation line"
} catch {
    Check $false ("section [5] ABORTED before finishing: $_")
} finally {
    if (Test-Path $scratchRoot) { Remove-Item $scratchRoot -Recurse -Force -EA SilentlyContinue }
    if (Test-Path $scratchDir)  { Remove-Item $scratchDir -Recurse -Force -EA SilentlyContinue }
}
Check (($checks - $checks5) -ge 31) ("section [5] ran its full complement of ladder checks (ran " + ($checks - $checks5) + ", expected at least 31)")
Check (-not (Test-Path $scratchRoot)) "the scratch registry subtree is gone"
Check ((Get-RealSecuritySnapshot) -eq $realBefore) "the machine's REAL Excel security settings are untouched"

# ---------------------------------------------------------------------------
# 6. the ladder is wired into every driver, at a failure point
# ---------------------------------------------------------------------------
# Wiring assertions prove NOTHING about whether the ladder works -- section [5]
# is what does that. These pin that each driver still contains a LIVE call to it
# and that the call cannot throw the script out of its own teardown.
#
# PARSED, NOT GREPPED (corrected 2026-08-03). These used to be
# `$src -match 'Write-XllLoadDiagnosis'` over the raw file text, and text does not
# know what a comment is: commenting out all three call sites in ghost-check.ps1
# left this suite at ALL CHECKS PASSED / exit 0. A wiring check that cannot see
# un-wiring is the exact failure class this file exists for. They now parse the
# driver and look for a real CommandAst.
#
# The companion check is new too. The old second assertion was
# `$src -notmatch 'Assert-XllLoaded'` -- a name that appears NOWHERE in this
# repository, so it could not fail under any edit. What it was trying to say is
# now checked structurally: no `throw` may follow a ladder call in the ladder
# call's own statement block. Every driver reaches that point holding live Excel
# and Go-server processes it still has to tear down (see the "deliberately NO
# throwing variant" note in uia-common.ps1).
function Get-DriverAst([string]$file) {
    $tok = $null; $perr = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$tok, [ref]$perr)
    if ($perr -and @($perr).Count -gt 0) { throw ("parse errors in {0}: {1}" -f $file, $perr[0].Message) }
    return $ast
}
function Get-LadderCalls($ast) {
    return @($ast.FindAll({
        param($n)
        ($n -is [System.Management.Automation.Language.CommandAst]) -and
        ($null -ne $n.GetCommandName()) -and
        ($n.GetCommandName() -ieq 'Write-XllLoadDiagnosis')
    }, $true))
}
# Walks up from a node to the statement that sits directly in a statement block.
function Get-EnclosingStatement($node) {
    $s = $node
    while ($null -ne $s -and $null -ne $s.Parent -and
           -not ($s.Parent -is [System.Management.Automation.Language.StatementBlockAst] -or
                 $s.Parent -is [System.Management.Automation.Language.NamedBlockAst])) {
        $s = $s.Parent
    }
    return $s
}
Write-Output "[6] wiring: every driver runs the ladder where the add-in should be loaded and is not"
foreach ($d in $drivers) {
    $ast   = Get-DriverAst (Join-Path $PSScriptRoot $d)
    $calls = Get-LadderCalls $ast
    Check ($calls.Count -ge 1) "$d runs the load-failure ladder (a parsed command, not a commented-out line)"
    $throwsAfter = 0
    foreach ($c in $calls) {
        $stmt = Get-EnclosingStatement $c
        if ($null -eq $stmt -or $null -eq $stmt.Parent) { continue }
        foreach ($s in @($stmt.Parent.Statements)) {
            if (($s.Extent.StartOffset -gt $stmt.Extent.StartOffset) -and
                ($s -is [System.Management.Automation.Language.ThrowStatementAst])) { $throwsAfter++ }
        }
    }
    Check ($throwsAfter -eq 0) "$d does not throw out of the ladder's own block (teardown must still run)"
}
# The drivers that never clear the native log must pass -Since, or a log left by
# an earlier run reads as proof that THIS run loaded the add-in. Checked on the
# parsed parameter list, so a -Since inside a comment does not count.
foreach ($d in @('diagnose-close-uia.ps1','ghost-check.ps1','repro-crash-uia.ps1','verify-rtd-stream-uia.ps1','verify-ydp-stranding-uia.ps1')) {
    $calls = Get-LadderCalls (Get-DriverAst (Join-Path $PSScriptRoot $d))
    $withSince = 0
    foreach ($c in $calls) {
        foreach ($e in @($c.CommandElements)) {
            if (($e -is [System.Management.Automation.Language.CommandParameterAst]) -and ($e.ParameterName -ieq 'Since')) { $withSince++ }
        }
    }
    Check (($calls.Count -ge 1) -and ($withSince -eq $calls.Count)) `
          ("$d passes -Since on every ladder call (it does not clear the native log); " + $withSince + " of " + $calls.Count)
}
$commonAst = Get-DriverAst (Join-Path $PSScriptRoot 'uia-common.ps1')
$ladderFns = @($commonAst.FindAll({
    param($n)
    ($n -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and
    ($n.Name -ieq 'Get-XllLoadFailureDiagnosis')
}, $true))
Check ($ladderFns.Count -eq 1) "the ladder lives in the shared scaffold as a real function definition"

Write-Output ""
if ($fails -eq 0) { Write-Output "ALL CHECKS PASSED" } else { Write-Output ("{0} CHECK(S) FAILED" -f $fails) }
exit $fails
