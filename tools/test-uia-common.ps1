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
function Check([bool]$ok, [string]$what) {
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

Write-Output ""
if ($fails -eq 0) { Write-Output "ALL CHECKS PASSED" } else { Write-Output ("{0} CHECK(S) FAILED" -f $fails) }
exit $fails
