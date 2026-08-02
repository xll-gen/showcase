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

. "$PSScriptRoot\uia-common.ps1"
$P = Resolve-ShowcasePaths; $xll = $P.Xll
# Environment before product: an unmet Trust Center lever produces this
# script's exact failure symptoms. See Assert-ExcelTrustPreconditions.
Assert-ExcelTrustPreconditions -XllPath $xll -RequireRtd
Initialize-Uia

Stop-ShowcaseProcesses
Start-Sleep 1

# A blank workbook is required (launching only the .xll shows the Start screen).
$tmp = New-ShowcaseWorkbook -Name "uia_showcase_repro.xlsx"

# Launch Excel NORMALLY (no held COM object): opens workbook + loads XLL addin + ribbon.
$p = Start-ShowcaseExcel -Workbook $tmp -Xll $xll
$pid2 = $p.Id; Write-Output "launched EXCEL pid=$pid2"

$win = Wait-ShowcaseWindow -LaunchedPid $pid2 -Tries 30
if (-not $win) { Write-Output "no ready window"; Stop-ShowcaseProcesses; return }

# Select the custom ribbon tab so its buttons enter the UIA tree, then Invoke the button.
if (-not (Invoke-RibbonButton -Window $win -NamePattern 'Build')) { Write-Output "no Build button"; Stop-ShowcaseProcesses; return }
Write-Output "invoked 'Build Showcase Sheet' — monitoring 32s for crash/hang..."

$verdict = "ALIVE+RESPONSIVE"; $hung = 0
for ($i = 0; $i -lt 16; $i++) {
    Start-Sleep 2
    $proc = Get-Process -Id $pid2 -EA SilentlyContinue
    if (-not $proc) { $verdict = "CRASHED after ~$((($i+1)*2))s"; break }
    if (-not $proc.Responding) { $hung++; if ($hung -ge 3) { $verdict = "HUNG after ~$((($i+1)*2))s"; break } } else { $hung = 0 }
}
Write-Output "VERDICT: $verdict"
Start-Sleep 1; Stop-ShowcaseProcesses
Write-Output "cleaned up"
