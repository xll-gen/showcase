# repro-crash-uia.ps1 — FAITHFUL repro of the xll-gen-showcase post-BuildShowcaseSheet
# Excel crash, WITHOUT a competing COM client (which causes false hangs).
#
# Why UIA and not COM automation: invoking BuildShowcaseSheet over a held
# Excel.Application COM object deadlocks Excel's STA against the command's OWN
# Go-side sugar COM automation (a TEST ARTIFACT). Driving the real ribbon button
# via UI Automation leaves the Go server as the SOLE COM client = faithful.
#
# Detection: polls Process.Responding (HANG) + process exit (CRASH). Headless /
# Visible=$false does NOT reproduce — the window must be real & visible.
#
# Usage:  pwsh -File tools\repro-crash-uia.ps1
# Build the (debug) XLL first:  cmake --build build/cpp --config Debug
# Then read build\xll_showcase_native.log / xll_showcase_go.log for the tail.

$xlexe = "C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE"
$xll   = "$PSScriptRoot\..\build\xll_showcase.xll"
$xll   = [System.IO.Path]::GetFullPath($xll)

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [Windows.Automation.AutomationElement]; $TS = [Windows.Automation.TreeScope]

Get-Process EXCEL, xll_showcase -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Start-Sleep 1

# A blank workbook is required (launching only the .xll shows the Start screen).
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) "uia_showcase_repro.xlsx"
$c = New-Object -ComObject Excel.Application; $c.DisplayAlerts = $false
$b = $c.Workbooks.Add(); $b.SaveAs($tmp, 51); $b.Close($false); $c.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($c)
Start-Sleep 2; Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue; Start-Sleep 1

# Launch Excel NORMALLY (no held COM object): opens workbook + loads XLL addin + ribbon.
$p = Start-Process $xlexe -ArgumentList "`"$tmp`"", "`"$xll`"" -PassThru
$pid2 = $p.Id; Write-Output "launched EXCEL pid=$pid2"

$win = $null
for ($i = 0; $i -lt 30; $i++) {
    Start-Sleep -Milliseconds 800
    $cond = New-Object Windows.Automation.PropertyCondition($AE::ProcessIdProperty, $pid2)
    $w = $AE::RootElement.FindFirst($TS::Children, $cond)
    if ($w -and ($w.Current.Name -match 'Excel') -and ($w.Current.Name -notmatch '여는')) { $win = $w; break }
}
if (-not $win) { Write-Output "no ready window"; Get-Process EXCEL, xll_showcase -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue; return }

# Select the custom ribbon tab so its buttons enter the UIA tree, then Invoke the button.
$tabCond = New-Object Windows.Automation.AndCondition(@(
    (New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [Windows.Automation.ControlType]::TabItem)),
    (New-Object Windows.Automation.PropertyCondition($AE::NameProperty, 'xll-gen Showcase')) ))
$tab = $win.FindFirst($TS::Descendants, $tabCond)
if ($tab) { $tab.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select(); Start-Sleep 1 }

$target = $null
$btnCond = New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [Windows.Automation.ControlType]::Button)
foreach ($bn in $win.FindAll($TS::Descendants, $btnCond)) { if ($bn.Current.Name -match 'Build') { $target = $bn; break } }
if (-not $target) { Write-Output "no Build button"; Get-Process EXCEL, xll_showcase -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue; return }
$target.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
Write-Output "invoked 'Build Showcase Sheet' — monitoring 32s for crash/hang..."

$verdict = "ALIVE+RESPONSIVE"; $hung = 0
for ($i = 0; $i -lt 16; $i++) {
    Start-Sleep 2
    $proc = Get-Process -Id $pid2 -EA SilentlyContinue
    if (-not $proc) { $verdict = "CRASHED after ~$((($i+1)*2))s"; break }
    if (-not $proc.Responding) { $hung++; if ($hung -ge 3) { $verdict = "HUNG after ~$((($i+1)*2))s"; break } } else { $hung = 0 }
}
Write-Output "VERDICT: $verdict"
Start-Sleep 1; Get-Process EXCEL, xll_showcase -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Write-Output "cleaned up"
