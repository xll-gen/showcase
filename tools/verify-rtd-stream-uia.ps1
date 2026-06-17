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

$ErrorActionPreference = 'Continue'
$xll = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\build\xll_showcase.xll")
if (-not (Test-Path $xll)) { Write-Output "MISSING XLL: $xll"; exit 2 }

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [Windows.Automation.AutomationElement]; $TS = [Windows.Automation.TreeScope]

Get-Process EXCEL, xll_showcase -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Start-Sleep 1

$app = New-Object -ComObject Excel.Application
$app.Visible = $true
$app.DisplayAlerts = $false
$app.Calculation = -4105  # xlCalculationAutomatic
$pidExcel = (Get-Process EXCEL -EA SilentlyContinue | Sort-Object StartTime | Select-Object -Last 1).Id
Write-Output "Excel COM up (pid=$pidExcel). Loading XLL: $xll"

$reg = $app.RegisterXLL($xll)
Write-Output "RegisterXLL returned: $reg"
Start-Sleep 3  # let xlAutoOpen launch the Go server + register RTD

$wb = $app.Workbooks.Add()
$ws = $wb.Worksheets.Item(1)
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

    # Scan for an Excel modal dialog (e.g. "...not responding...")
    $proc = Get-Process -Id $pidExcel -EA SilentlyContinue
    if ($proc) {
        $cond = New-Object Windows.Automation.PropertyCondition($AE::ProcessIdProperty, $pidExcel)
        foreach ($w in $AE::RootElement.FindAll($TS::Children, $cond)) {
            $nm = $w.Current.Name
            if ($nm -match 'responding|응답|RTD|연결') { $dialog = $nm }
        }
        if (-not $proc.Responding) { $samples += "  (Excel NOT responding at t=$((($i+1)*2))s)" }
    }
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
try { $wb.Close($false) } catch {}
try { $app.Quit() } catch {}
try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws) } catch {}
try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch {}
try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {}
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Start-Sleep 2
Get-Process EXCEL, xll_showcase -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Write-Output "cleaned up"
