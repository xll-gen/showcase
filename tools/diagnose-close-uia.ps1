# diagnose-close-uia.ps1 — UIA repro + evidence capture for the two close-time bugs:
#   (S1) "창을 다 닫았는데 창이 하나 다시 열린다"  -> uncancelled xlcOnTime reopen
#   (S2) "native/go log를 못 지운다, Excel 프로세스는 없다" -> orphaned Go server holding inherited log handle
#
# Faithful close: drives the REAL window close via UIA WindowPattern.Close() (a user
# clicking X), handling the Korean/English save dialog — NOT a COM Application.Quit
# (which would mask the bug). After close it watches for:
#   * REOPEN  : any Excel top-level window / new EXCEL pid reappearing post-close
#   * ORPHAN  : xll_showcase server process still alive after EXCEL is gone
#   * LOCK    : _go.log / _native.log still locked, with the holding process named
#               via the Restart Manager API (rstrtmgr.dll) — no Sysinternals needed.
#
# Usage:  pwsh -File tools\diagnose-close-uia.ps1
# Requires the built XLL at ..\build\xll_showcase.xll and an interactive desktop session.

param([switch]$KillExcelOnClose)   # simulate user force-killing the ghost Excel (Task Manager / crash)
$ErrorActionPreference = 'Continue'
$xlexe   = "C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE"
$xll     = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\build\xll_showcase.xll")
$buildDir= [System.IO.Path]::GetFullPath("$PSScriptRoot\..\build")
$goLog     = Join-Path $buildDir "xll_showcase_go.log"
$nativeLog = Join-Path $buildDir "xll_showcase_native.log"

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [Windows.Automation.AutomationElement]; $TS = [Windows.Automation.TreeScope]

# --- Restart Manager: which process(es) hold a lock on $path -------------------
Add-Type -Namespace RM -Name Locks -UsingNamespace System.Collections.Generic -MemberDefinition @'
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
        $who = [RM.Locks]::Holders($path)
        return "LOCKED by -> $who"
    }
}

function Excel-Windows {
    # visible top-level windows whose name looks like an Excel window
    $cond = New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [Windows.Automation.ControlType]::Window)
    $all = $AE::RootElement.FindAll($TS::Children, $cond)
    $out = @()
    foreach ($w in $all) {
        try { if ($w.Current.Name -match 'Excel' -and $w.Current.Name -notmatch '여는|Opening') { $out += $w } } catch {}
    }
    return $out
}

Write-Output "=== diagnose-close-uia : XLL=$xll ==="
Get-Process EXCEL, xll_showcase -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Start-Sleep 1

# blank workbook (XLL alone shows Start screen)
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) "uia_close_diag.xlsx"
$c = New-Object -ComObject Excel.Application; $c.DisplayAlerts = $false
$b = $c.Workbooks.Add(); $b.SaveAs($tmp, 51); $b.Close($false); $c.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($c)
Start-Sleep 2; Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue; Start-Sleep 1

# launch Excel normally (no held COM client)
$p = Start-Process $xlexe -ArgumentList "`"$tmp`"", "`"$xll`"" -PassThru
$origPid = $p.Id
Write-Output "launched EXCEL pid=$origPid"

$win = $null
for ($i=0; $i -lt 30; $i++) {
    Start-Sleep -Milliseconds 800
    $cond = New-Object Windows.Automation.PropertyCondition($AE::ProcessIdProperty, $origPid)
    $w = $AE::RootElement.FindFirst($TS::Children, $cond)
    if ($w -and ($w.Current.Name -match 'Excel') -and ($w.Current.Name -notmatch '여는|Opening')) { $win = $w; break }
}
if (-not $win) { Write-Output "FAIL: no ready window"; Get-Process EXCEL, xll_showcase -EA SilentlyContinue | Stop-Process -Force; return }
Start-Sleep 2

$srvBefore = Get-Process xll_showcase -EA SilentlyContinue
Write-Output ("server pids after load: " + (($srvBefore.Id) -join ',' ))

# select custom ribbon tab + invoke Build (heavy calc -> CalculationEnded -> arms xlcOnTime)
$tabCond = New-Object Windows.Automation.AndCondition(@(
    (New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [Windows.Automation.ControlType]::TabItem)),
    (New-Object Windows.Automation.PropertyCondition($AE::NameProperty, 'xll-gen Showcase')) ))
$tab = $win.FindFirst($TS::Descendants, $tabCond)
if ($tab) { try { $tab.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select(); Start-Sleep 1 } catch {} }
$btnCond = New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [Windows.Automation.ControlType]::Button)
foreach ($bn in $win.FindAll($TS::Descendants, $btnCond)) { if ($bn.Current.Name -match 'Build') {
    try { $bn.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke(); Write-Output "invoked Build" } catch {}; break } }
Start-Sleep 5   # let the build/calc run and arm the deferred runner

# === FAITHFUL CLOSE via WindowPattern.Close() (user clicking X) ===============
Write-Output "--- closing window via UIA (faithful) ---"
try { $win.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern).Close() } catch { Write-Output "close pattern failed: $_" }

# handle "변경 내용을 저장하시겠습니까?" -> click 저장 안 함 / Don't Save
for ($i=0; $i -lt 10; $i++) {
    Start-Sleep -Milliseconds 500
    $dlgBtn = $null
    foreach ($w in (Excel-Windows)) {
        foreach ($bn in $w.FindAll($TS::Descendants, $btnCond)) {
            if ($bn.Current.Name -match '저장 안|안 함|Don.t Save|Do.?n.t') { $dlgBtn = $bn; break }
        }
        if ($dlgBtn) { break }
    }
    if ($dlgBtn) { try { $dlgBtn.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke(); Write-Output "clicked Don't-Save" } catch {}; break }
}

# simulate the user killing the ghost Excel right after close (Task Manager / crash):
# this races the graceful teardown's CloseHandle(hJob) and exposes server orphaning.
if ($KillExcelOnClose) {
    Write-Output "--- KILL: force-terminating ghost EXCEL immediately (simulating user/crash) ---"
    Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
}

# === OBSERVE NATURAL FATE for 40s — NO force-kill until the very end =========
# Sample every 1s: EXCEL pids, visible Excel windows, server pids. Record the
# moment the window disappears, the moment it REAPPEARS (S1), whether EXCEL ever
# fully exits, and whether the server outlives EXCEL (S2 orphan).
Write-Output "--- observing natural fate for 40s (no kill) ---"
$windowGoneAt = $null; $reopenAt = $null; $reopenPid = $null
$excelExitAt  = $null; $sawWindowGone = $false
for ($t=0; $t -lt 40; $t++) {
    Start-Sleep 1
    $wins  = Excel-Windows
    $exPid = (Get-Process EXCEL -EA SilentlyContinue).Id
    $svPid = (Get-Process xll_showcase -EA SilentlyContinue).Id
    if (-not $sawWindowGone -and $wins.Count -eq 0) { $windowGoneAt = $t; $sawWindowGone = $true }
    if ($sawWindowGone -and $wins.Count -gt 0 -and -not $reopenAt) {
        $reopenAt = $t; try { $reopenPid = $wins[0].Current.ProcessId } catch {}
    }
    if (-not $exPid -and -not $excelExitAt) { $excelExitAt = $t }
    Write-Output ("  t={0,2}s  win={1}  EXCEL=[{2}]  server=[{3}]" -f $t, $wins.Count, ($exPid -join ','), ($svPid -join ','))
    # stop early once fully settled: no window, no excel, no server
    if ($wins.Count -eq 0 -and -not $exPid -and -not $svPid) { Write-Output "  (fully settled)"; break }
}

Write-Output "--- final state ---"
$excelNow = (Get-Process EXCEL -EA SilentlyContinue).Id
$srvNow   = (Get-Process xll_showcase -EA SilentlyContinue).Id
$orphan   = ($srvNow -and -not $excelNow)
$ghost    = ($excelNow -and (Excel-Windows).Count -eq 0)   # alive process, no window
Write-Output ("EXCEL now: [" + ($excelNow -join ',') + "]   server now: [" + ($srvNow -join ',') + "]")
Write-Output ("window gone at: " + $windowGoneAt + "s   reopened at: " + $reopenAt + "s (pid=$reopenPid, orig=$origPid)   EXCEL exited at: " + $excelExitAt + "s")
Write-Output ("LOCK _go.log     : " + (Probe-Lock $goLog))
Write-Output ("LOCK _native.log : " + (Probe-Lock $nativeLog))

Write-Output "=== VERDICT ==="
Write-Output ("  S1 window reopened          : " + $(if($reopenAt -ne $null){"YES at ${reopenAt}s (pid=$reopenPid)"}else{"no"}))
Write-Output ("  S1' ghost EXCEL (no window) : " + $(if($ghost){"YES (pid=$excelNow lingering windowless)"}else{"no"}))
Write-Output ("  S2 orphan server (no EXCEL) : " + $(if($orphan){"YES (server=$srvNow)"}else{"no"}))

# cleanup (only now)
Get-Process EXCEL, xll_showcase -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Write-Output "cleaned up"
