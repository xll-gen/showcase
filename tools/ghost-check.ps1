# ghost-check.ps1 -- focused S1' ghost-Excel repro + evidence capture.
#
# Launches Excel with the showcase XLL, waits robustly for the ready window,
# invokes the Build ribbon command (heavy calc -> CalculationEnded -> arms the
# xlcOnTime deferred runner), then faithfully closes via WindowPattern.Close()
# (a user clicking X -- NOT COM Application.Quit, which would mask the bug).
# After close it observes for 30s WITHOUT force-killing and reports:
#   * whether EXCEL.EXE actually EXITS (the S1' ghost question)
#   * whether the Go server is reaped
#   * whether _go.log / _native.log become FREE
#   * whether a window reopens (S1)
#
# -KillExcelOnClose : simulate the user force-killing the ghost right at close
#                     (races graceful teardown's CloseHandle(hJob) -> S2 orphan).
#
# This is a more robust sibling of diagnose-close-uia.ps1: it settles before the
# first window poll (the COM-bounce variant makes Excel slow to show its window)
# and tracks the window-owning pid (EXCEL.EXE can relaunch off the Start-Process pid).
param([switch]$KillExcelOnClose, [switch]$FastClose)
$ErrorActionPreference = 'Continue'
$xlexe   = "C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE"
$xll     = [System.IO.Path]::GetFullPath("$PSScriptRoot\..\build\xll_showcase.xll")
$buildDir= [System.IO.Path]::GetFullPath("$PSScriptRoot\..\build")
$goLog     = Join-Path $buildDir "xll_showcase_go.log"
$nativeLog = Join-Path $buildDir "xll_showcase_native.log"

Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [Windows.Automation.AutomationElement]; $TS = [Windows.Automation.TreeScope]

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
    try { $fs=[System.IO.File]::Open($path,'Open','ReadWrite','None'); $fs.Close(); return "FREE (deletable)" }
    catch { return "LOCKED by -> " + [RM.Locks]::Holders($path) }
}
# Returns a (possibly empty) [System.Collections.ArrayList] of AutomationElement
# top-level Excel windows. Using an explicit ArrayList (not PowerShell array
# unrolling) keeps the return type stable whether 0, 1, or N windows match.
#
# IMPORTANT (encoding): this file must NOT embed non-ASCII (Korean) regex
# literals -- when the .ps1 is read under a non-UTF-8 codepage the bytes get
# mangled into an INVALID regex, and the resulting parse exception (thrown by
# -match/-notmatch) is swallowed by the surrounding try/catch, silently
# dropping the real Excel window ("FAIL: no ready window"). So we identify the
# transient "Opening..." splash by ASCII-only means: the ready workbook window's
# title contains the workbook file extension '.xlsx', the splash does not.
#   -RequireWorkbook : match only the ready '<file>.xlsx - Excel' window
#                      (acquisition); excludes the splash + the Start screen.
#   (default)        : match ANY 'Excel' window EXCEPT the ASCII 'Opening'
#                      splash (post-close ghost detection -- a lingering
#                      windowless ghost has NO window, so any window counts).
function Excel-Windows {
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

Write-Output "=== ghost-check : XLL=$xll ==="
Get-Process EXCEL, xll_showcase -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Start-Sleep 1

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) "ghost_check.xlsx"
$c = New-Object -ComObject Excel.Application; $c.DisplayAlerts = $false
$b = $c.Workbooks.Add(); $b.SaveAs($tmp, 51); $b.Close($false); $c.Quit()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($c)
Start-Sleep 2; Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue; Start-Sleep 2

$p = Start-Process $xlexe -ArgumentList "`"$tmp`"", "`"$xll`"" -PassThru
Write-Output "launched EXCEL pid=$($p.Id)"

# Robust wait: settle first (the COM-bounce makes Excel slow to materialize, and
# polling UIA too eagerly during init can transiently throw on Current.Name),
# then poll up to ~30s.
Start-Sleep 6
$win = $null; $excelPid = $p.Id
for ($i=0; $i -lt 38; $i++) {
    $cands = Excel-Windows -RequireWorkbook
    if ($cands.Count -gt 0) { $win = $cands[0]; try { $excelPid = $win.Current.ProcessId } catch {}; break }
    Start-Sleep -Milliseconds 800
}
if (-not $win) { Write-Output "FAIL: no ready window"; Get-Process EXCEL, xll_showcase -EA SilentlyContinue | Stop-Process -Force; return }
Write-Output "ready Excel window pid=$excelPid (launched pid=$($p.Id))"
Start-Sleep 2

$srvBefore = Get-Process xll_showcase -EA SilentlyContinue
Write-Output ("server pids after load: " + (($srvBefore.Id) -join ','))

# select custom ribbon tab + invoke Build (heavy calc -> CalculationEnded -> arms xlcOnTime)
$tabCond = New-Object Windows.Automation.AndCondition(@(
    (New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [Windows.Automation.ControlType]::TabItem)),
    (New-Object Windows.Automation.PropertyCondition($AE::NameProperty, 'xll-gen Showcase')) ))
$tab = $win.FindFirst($TS::Descendants, $tabCond)
if ($tab) { try { $tab.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select(); Start-Sleep 1; Write-Output "selected showcase ribbon tab" } catch { Write-Output "tab select failed: $_" } }
else { Write-Output "WARN: showcase ribbon tab not found" }
$btnCond = New-Object Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [Windows.Automation.ControlType]::Button)
$invoked = $false
foreach ($bn in $win.FindAll($TS::Descendants, $btnCond)) {
    try { if ($bn.Current.Name -match 'Build') { $bn.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke(); Write-Output "invoked Build"; $invoked=$true; break } } catch {}
}
if (-not $invoked) { Write-Output "WARN: Build button not found/invoked" }
# -FastClose: close almost immediately after Build to catch the xlcOnTime
# deferred runner while it is still ARMED (it arms at CalculationEnded and Excel
# dispatches the OnTime macro at the next idle, which a normal 5s wait lets drain).
# This is the HIGH #2 verification: does CancelDeferredRunner de-queue an armed runner?
if ($FastClose) { Start-Sleep -Milliseconds 250; Write-Output "FAST CLOSE (250ms after Build)" } else { Start-Sleep 5 }

Write-Output "--- closing window via UIA (faithful WindowPattern.Close) ---"
try { $win.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern).Close() } catch { Write-Output "close pattern failed: $_" }

# Dismiss the "Save changes? -> Don't Save" prompt. This whole file is ASCII-ONLY
# on purpose: a .ps1 without a UTF-8 BOM is read by Windows PowerShell under the
# system ANSI codepage (e.g. CP949 on Korean Windows), which mangles any non-ASCII
# (Korean) literal into bytes that either break parsing or form an invalid regex
# whose exception is swallowed by a try/catch -- the latter silently dropped the
# real Excel window and produced "FAIL: no ready window". So the Korean "Don't
# Save" button is built by codepoint (U+C800 U+C7A5 = "Don't Save"), never as a
# source literal; a keyboard 'n' fallback below also dismisses it.
$dontSaveKR = [string][char]0xC800 + [string][char]0xC7A5  # localized "Don't Save" prefix
for ($i=0; $i -lt 12; $i++) {
    Start-Sleep -Milliseconds 500
    $dlgBtn = $null
    foreach ($w in (Excel-Windows)) {
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
foreach ($w in (Excel-Windows)) {
    $hasDialog = $false
    foreach ($bn in $w.FindAll($TS::Descendants, $btnCond)) {
        try { if ($bn.Current.Name -match 'Save|Cancel') { $hasDialog = $true; break } } catch {}
    }
    if ($hasDialog) {
        try { $shell = New-Object -ComObject WScript.Shell; $shell.SendKeys('n'); Write-Output "sent 'n' (Don't-Save accelerator)" } catch {}
        break
    }
}

# Faithful-close fallback. In some environments WindowPattern.Close() on the
# workbook window dismisses the document but leaves the Excel application frame
# running (no quit -> no OnBeginShutdown/OnDisconnection -> teardown never fires
# and the run is inconclusive). If a top-level Excel window is still present a
# moment after the Close()+Don't-Save, escalate to Alt+F4 on the foreground Excel
# window. Alt+F4 is a FAITHFUL user "close the window" gesture (NOT COM
# Application.Quit, which would mask the S1' bug by tearing RTD down on a
# different code path). We only do this when a window lingers, so genuine clean
# closes are unaffected.
Start-Sleep -Milliseconds 800
$lingering = Excel-Windows
if ($lingering.Count -gt 0 -and -not $KillExcelOnClose) {
    Write-Output "--- window still present after Close(); escalating to faithful Alt+F4 ---"
    for ($k=0; $k -lt 5; $k++) {
        $cur = Excel-Windows
        if ($cur.Count -eq 0) { break }
        try {
            $w = $cur[0]
            # Bring to foreground, then Alt+F4.
            try { $w.SetFocus() } catch {}
            Start-Sleep -Milliseconds 200
            $shell = New-Object -ComObject WScript.Shell
            $shell.SendKeys('%{F4}')
            Write-Output "sent Alt+F4"
        } catch { Write-Output "Alt+F4 failed: $_" }
        # Re-dismiss any Save prompt the Alt+F4 raised.
        Start-Sleep -Milliseconds 600
        foreach ($w in (Excel-Windows)) {
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

if ($KillExcelOnClose) {
    Write-Output "--- KILL: force-terminating EXCEL immediately (simulating user/crash) ---"
    Get-Process EXCEL -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
}

Write-Output "--- observing natural fate for 30s (no kill) ---"
$windowGoneAt=$null; $reopenAt=$null; $reopenPid=$null; $excelExitAt=$null; $sawWindowGone=$false
for ($t=0; $t -lt 30; $t++) {
    Start-Sleep 1
    $wins = Excel-Windows
    $exPid = (Get-Process EXCEL -EA SilentlyContinue).Id
    $svPid = (Get-Process xll_showcase -EA SilentlyContinue).Id
    if (-not $sawWindowGone -and $wins.Count -eq 0) { $windowGoneAt=$t; $sawWindowGone=$true }
    if ($sawWindowGone -and $wins.Count -gt 0 -and -not $reopenAt) { $reopenAt=$t; try { $reopenPid=$wins[0].Current.ProcessId } catch {} }
    if (-not $exPid -and -not $excelExitAt) { $excelExitAt=$t }
    Write-Output ("  t={0,2}s  win={1}  EXCEL=[{2}]  server=[{3}]" -f $t, $wins.Count, ($exPid -join ','), ($svPid -join ','))
    if ($wins.Count -eq 0 -and -not $exPid -and -not $svPid) { Write-Output "  (fully settled)"; break }
}

Write-Output "--- final state ---"
$excelNow=(Get-Process EXCEL -EA SilentlyContinue).Id
$srvNow=(Get-Process xll_showcase -EA SilentlyContinue).Id
$orphan=($srvNow -and -not $excelNow)
$ghost=($excelNow -and (Excel-Windows).Count -eq 0)
Write-Output ("EXCEL now: [" + ($excelNow -join ',') + "]   server now: [" + ($srvNow -join ',') + "]")
Write-Output ("window gone at: ${windowGoneAt}s   reopened at: ${reopenAt}s (pid=$reopenPid)   EXCEL exited at: ${excelExitAt}s")
Write-Output ("LOCK _go.log     : " + (Probe-Lock $goLog))
Write-Output ("LOCK _native.log : " + (Probe-Lock $nativeLog))
Write-Output "=== VERDICT ==="
Write-Output ("  S1 window reopened          : " + $(if($reopenAt -ne $null){"YES at ${reopenAt}s (pid=$reopenPid)"}else{"no"}))
Write-Output ("  S1' ghost EXCEL (no window) : " + $(if($ghost){"YES (pid=$excelNow lingering windowless)"}else{"no"}))
Write-Output ("  S2 orphan server (no EXCEL) : " + $(if($orphan){"YES (server=$srvNow)"}else{"no"}))
Get-Process EXCEL, xll_showcase -EA SilentlyContinue | Stop-Process -Force -EA SilentlyContinue
Write-Output "cleaned up"
