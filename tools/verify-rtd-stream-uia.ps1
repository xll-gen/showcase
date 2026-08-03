# verify-rtd-stream-uia.ps1 — verify streaming RTD keeps updating over time
# (regression check for the v0.8.8 fix: streaming RTD handler ctx was being
# cancelled on connect-return, freezing the stream after one value, which Excel
# then flagged as "RTD server not responding").
#
# Method: load the showcase XLL into a VISIBLE Excel via RegisterXLL, enter
# =Clock() (1s tick) and =StockTick("AAPL") (750ms tick), then sample both cells
# for ~30s. PASS = many distinct Clock values (stream alive). FAIL = frozen at
# one value (the pre-fix bug). Also watches for a modal "not responding" dialog
# and Process.Responding. Two-tier cleanup at the end.

. "$PSScriptRoot\uia-common.ps1"
$ErrorActionPreference = 'Continue'
$xll = (Resolve-ShowcasePaths).Xll
if (-not (Test-Path $xll)) { Write-Output "MISSING XLL: $xll"; exit 2 }
# Environment before product: an unmet Trust Center lever produces this
# script's exact failure symptoms. See Assert-ExcelTrustPreconditions.
Assert-ExcelTrustPreconditions -XllPath $xll -RequireRtd
# Drop post-crash resiliency residue BEFORE launching. DisabledItems and
# StartupItems are pure harness debris -- a previous crashed run can leave the
# add-in on Excel's disabled list, and then this script measures an add-in that
# was never loaded and blames the product. DocumentRecovery is deliberately NOT
# touched (it can hold a real user's unsaved workbooks); the gate above WARNs
# about it and leaves the choice to a human.
Clear-ExcelResiliency

Initialize-Uia

Stop-ShowcaseProcesses
Start-Sleep 1

$app = New-Object -ComObject Excel.Application
$app.Visible = $true
$app.DisplayAlerts = $false
$pidExcel = (Get-Process EXCEL -EA SilentlyContinue | Sort-Object StartTime | Select-Object -Last 1).Id
Write-Output "Excel COM up (pid=$pidExcel). Loading XLL: $xll"

$reg = $app.RegisterXLL($xll)
Write-Output "RegisterXLL returned: $reg"
Start-Sleep 3  # let xlAutoOpen launch the Go server + register RTD

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
# Calculation is only settable with a workbook open -- setting it right after
# CreateObject threw "cannot set the Calculation property" on every single run.
# It was non-fatal ($ErrorActionPreference = 'Continue') and the stream verdict
# was unaffected, which is exactly why it survived: a red error block printed on
# every run is one an operator learns to scroll past, and the next real one goes
# with it. (verify-ydp-stranding-uia.ps1 already had this in the right place.)
try { $app.Calculation = -4105 } catch { Write-Output "(could not set automatic calculation: $_)" }
$ws.Range("A1").Formula = "=Clock()"
$ws.Range("B1").Formula = "=StockTick(""AAPL"")"
Write-Output "Entered =Clock() in A1, =StockTick(AAPL) in B1. Sampling 30s..."

$clockVals = New-Object System.Collections.Generic.HashSet[string]
$stockVals = New-Object System.Collections.Generic.HashSet[string]
$dialog = $null
$samples = @()
for ($i = 0; $i -lt 15; $i++) {
    Start-Sleep 2
    try { $a = [string]$ws.Range("A1").Value2 } catch { $a = "<err>" }
    try { $b = [string]$ws.Range("B1").Value2 } catch { $b = "<err>" }
    if ($a) { [void]$clockVals.Add($a) }
    if ($b) { [void]$stockVals.Add($b) }
    $samples += ("t={0,2}s  A1='{1}'  B1='{2}'" -f (($i+1)*2), $a, $b)

    # Modal dialog scan. Uses the shared IsModal-based detector rather than the
    # title regex this used to carry, for two reasons found on 2026-08-03:
    #   1. That regex contained Korean alternatives as SOURCE LITERALS in a
    #      .ps1 with NO UTF-8 BOM -- precisely the encoding trap uia-common.ps1
    #      documents at the top of the file. Under Windows PowerShell's ANSI
    #      codepage those bytes are mangled, so the localized half of the
    #      pattern could never match on the very machines it was added for, and
    #      nothing said so.
    #   2. It matched on the dialog's TITLE, so it only found dialogs whose
    #      names someone had thought of. WindowPattern.IsModal is the property
    #      Windows itself sets, so it needs no such list.
    $dlg = Get-ExcelModalDialog -ProcessId $pidExcel
    if ($dlg) { $dialog = $dlg.Name + $(if ($dlg.Text) { " -- " + $dlg.Text } else { "" }) }
    $proc = Get-Process -Id $pidExcel -EA SilentlyContinue
    if ($proc -and -not $proc.Responding) { $samples += "  (Excel NOT responding at t=$((($i+1)*2))s)" }
}

$samples | ForEach-Object { Write-Output $_ }
Write-Output "distinct Clock values: $($clockVals.Count) -> [$([string]::Join(', ', $clockVals))]"
Write-Output "distinct StockTick values: $($stockVals.Count)"
if ($dialog) { Write-Output "DIALOG DETECTED: $dialog" }

$verdict = if ($dialog) { "FAIL (dialog: $dialog)" }
           elseif ($clockVals.Count -ge 4) { "PASS (Clock advanced $($clockVals.Count) distinct values)" }
           elseif ($clockVals.Count -le 1) { "FAIL (Clock FROZEN at $($clockVals.Count) value — stream died, pre-fix bug)" }
           else { "WEAK ($($clockVals.Count) distinct — inconclusive)" }
Write-Output "VERDICT: $verdict"

# Two-tier cleanup: graceful quit, then force-kill.
Stop-ShowcaseCom -App $app -Workbook $wb -Worksheet $ws
Write-Output "cleaned up"
