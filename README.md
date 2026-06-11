# xll-gen-showcase

A single, self-contained add-in that exercises **every xll-gen feature** and
doubles as the manual end-to-end (E2E) checklist for the ribbon/command
feature.

It declares one of each worksheet-function type and execution mode, two RTD
streams, five ribbon commands driven by [sugar](https://github.com/xll-gen/sugar)
Excel automation, and a `CalculationEnded` event handler that uses the
command-scheduling path. Everything is hand-authored (no `xll-gen init`) and
curated so the layout reads as documentation.

- Project name: `xll_showcase`
- Go module: `xll_showcase`
- RTD ProgID: `XllShowcase.Rtd`; Ribbon ProgID: `XllShowcase.Ribbon`
- Build mode: `singlefile: ""` (**standalone** — the `.xll` and the Go server
  `.exe` are separate files). Standalone is the simplest shape for a demo: the
  XLL launches the server, and you can rebuild the Go side without re-linking
  C++. Use `singlefile: "xll"` for single-file distribution.

## Prerequisites

| Tool | Why | Notes |
|------|-----|-------|
| Go (matching `go.mod`, ≥ 1.24.3) | builds the Go server | required |
| CMake ≥ 3.20 | configures the C++ XLL | required for the `.xll` |
| A C++ toolchain (MinGW-w64 **or** MSVC) | compiles the XLL | required for the `.xll` |
| [Task](https://taskfile.dev) | convenience build runner | optional |
| Microsoft Excel (Windows) | runtime / E2E checklist | required to *run* it |

The `xll-gen` CLI itself is built from the local checkout (see below). All
module dependencies are public tagged releases (`xll-gen v0.4.0`,
`sugar v0.8.0`, `types v0.2.8`, `shm v0.7.5`) — no `replace` directives.

## Build

```sh
# 1. Build the xll-gen CLI from the local checkout (one time).
cd ../xll-gen && go build -o xll-gen.exe . && cd -

# 2. Generate Go + C++ sources from xll.yaml (also runs `go mod tidy`).
../xll-gen/xll-gen.exe generate

# 3. Build the Go server (MANDATORY gate — this must compile).
go build ./...                          # verify everything compiles
go build -ldflags="-H=windowsgui -s -w" -o build/xll_showcase.exe .

# 4. Build the C++ XLL.
#    MinGW:
cmake -S generated/cpp -B build/cpp -G "MinGW Makefiles" -DCMAKE_BUILD_TYPE=Release
cmake --build build/cpp --config Release
#    MSVC: drop the -G flag (CMake picks the Visual Studio generator).
```

Or, with Task: `task build-go && task build` (the generated `Taskfile.yml`
covers the standalone flow).

The resulting `.xll` is under `build/cpp/`. Place it next to
`build/xll_showcase.exe` (or point the launch config at the server) and open
the `.xll` in Excel.

### Build status (this checkout)

- **Go server: compiles cleanly.** `go build ./...` and `go vet ./...` both
  pass, and `go build -o build/xll_showcase.exe .` produces the server. This
  is the mandatory gate and it is green.
- **C++ XLL: blocked by two xll-gen v0.4.0 generator bugs.** CMake configures
  and the build reaches 96% (every static dep + all sources compile) before
  `xll_main.cpp` fails in exactly two generated functions:
  - `SumGrid` (a `grid` argument) — the generator emits a call to
    `GridToFlatBuffer(...)`, but the types library only provides
    `ConvertGrid(...)`. Undeclared identifier.
  - `WhoAmI` (`caller: true`) — the generated caller-resolution code passes
    `ScopedXLOPER12*` to `xll::CallExcel` / `ConvertRange` (which expect
    `LPXLOPER12`) and dereferences a `ScopedXLOPER12` with `->`. Type
    mismatches.

  Both are defects in xll-gen v0.4.0's **C++** templates for the `grid`-arg and
  `caller`-aware paths; the Go side of those same functions compiles fine. No
  other feature errors — sync/volatile/async/RTD/commands/ribbon/event all
  compile. To produce a runnable `.xll` from this checkout you must either fix
  those two C++ templates in xll-gen or drop `SumGrid`/`WhoAmI` from
  `xll.yaml`.

## Feature coverage

Every declared feature, and where to observe it. After clicking **Build
Showcase Sheet** the live cells appear in column B of the demo worksheet.

| Feature | Declared as | Observe via |
|---------|-------------|-------------|
| sync int | `Add(a,b int) -> int` | cell `=Add(2,3)` → 5 |
| sync float | `Multiply(a,b float) -> float` | cell `=Multiply(1.5,4)` → 6 |
| sync string | `Greet(name string) -> string` | cell `=Greet("Excel")` |
| sync bool | `IsEven(n int) -> bool` | cell `=IsEven(10)` → TRUE |
| `grid` arg (2-D `any` cells) | `SumGrid(g grid) -> float` | cell `=SumGrid(E4:F5)` → 10 |
| caller-aware | `WhoAmI() -> string` (`caller: true`) | cell `=WhoAmI()` reports its own address |
| volatile | `RandomLine() -> string` (`volatile: true`) | cell `=RandomLine()`; changes on F9 |
| async | `SlowSquare(n int) -> int` (`mode: async`) | cell `=SlowSquare(9)`; ~1.5s then 81 |
| RTD (`any` return) | `Clock() -> any` (`mode: rtd`) | cell `=Clock()`; ticks 1/s |
| RTD with arg | `StockTick(symbol string) -> any` | cell `=StockTick("AAPL")`; wandering price |
| command (sugar) | `BuildShowcaseSheet` | ribbon **Demo → Build Showcase Sheet** (large) |
| command (sugar) | `ClearShowcase` | ribbon **Demo → Clear Showcase** |
| command + shortcut | `WriteTimestamp` (`shortcut: T`) | ribbon **Commands** / **Ctrl+Shift+T** → writes A1 |
| command + shortcut | `SlowFill` (`shortcut: S`) | ribbon **Commands** / **Ctrl+Shift+S** → 5s then "done" |
| command (CommandContext) | `ShowContext` | ribbon **Commands → Show Context** → name/control/PID into cells |
| ribbon (structured) | `tab: "xll-gen Showcase"`, groups Demo + Commands | custom ribbon tab |
| event + scheduling | `CalculationEnded` → `OnRecalc` | press **F9**; "Last recalc" cell stamps a time |

### Intentional omission

A **sync** function with `return: "any"` (e.g. `Echo(v any) -> any`) is
**omitted**. xll-gen v0.4.0's Go generator emits
`ipc.EchoResponseAddResult(b, res)`, passing a `*protocol.Any` where the
FlatBuffers setter expects a `flatbuffers.UOffsetT` — that generated code does
not compile. The `any` type is still fully covered: as a worksheet-function
**argument** (grid cells are `any`-typed scalars) and as an **RTD return**
(`Clock`/`StockTick` stream `any`). This is a genuine generator limitation
surfaced by the showcase, not a gap in the demo.

## sugar API used by the command handlers

The command handlers attach to the **running** Excel instance and drive it
through sugar. APIs actually used (verified against `sugar/excel/*.go`):

- `excel.GetApplication(ctx)` → `Application`
- `app.Books().Active()` → `Workbook`; `wb.Sheets().Active()` → `Worksheet`
- `sheet.Range(addr)`, `sheet.UsedRange()`, `sheet.Clear()`, `sheet.AutoFit()`
- `Range.SetValue`, `Range.SetFormula`, `Range.SetColor`, `Range.Find`,
  `Range.Row`, `Range.Err`
- `Range.Font()` → `Font.SetBold` / `SetSize` / `SetColor`
- `excel.RGB(r,g,b)` for fills and font colors
- Arena lifecycle via `sugar.Do(func(ctx sugar.Context) error { ... })`, which
  COM-initializes the thread (STA-correct) and **auto-releases every tracked
  COM object** when the callback returns — the real equivalent of
  "`defer ctx.Release()`". Errors from `GetApplication` are surfaced and logged
  (graceful degradation), never panicked.

`ShowContext` reads `cmd.CommandName`, `cmd.ControlID`, and `cmd.ExcelPID` from
the `server.CommandContext` passed to every command handler.

## End-to-end checklist (run on a real Excel machine)

Run before any release that touches ribbon/command code paths. Folds in the
7-point ribbon checklist and expands it to cover functions, async, and RTD.
Each item maps to a button or cell.

**Ribbon / command (the 7 original points):**

1. **Tab appears.** Open the `.xll`; the **xll-gen Showcase** tab shows the
   **Demo** group (Build Showcase Sheet [large], Clear Showcase) and the
   **Commands** group (Write Timestamp, Slow Fill (5s), Show Context). Any
   non-ASCII labels render correctly (embedded as XML numeric character refs).
2. **Button → handler.** Click **Write Timestamp** → A1 shows
   `WriteTimestamp @ HH:MM:SS` (the Go handler ran and drove Excel via sugar).
3. **Slow handler stays responsive.** Click **Slow Fill (5s)**; immediately
   type in other cells / scroll. Excel stays responsive for the full 5s, then
   A1 shows `SlowFill done @ …` — proves the fire-and-forget STA contract.
4. **Shortcut.** Press **Ctrl+Shift+T** (no ribbon) → A1 updates. Press
   **Ctrl+Shift+S** → 5s later "done".
5. **Alt+F8.** Open the macro dialog, type `ShowContext`, **Run** → A1:B3 fill
   with CommandName / ControlID (empty for Alt+F8) / ExcelPID. (XLL commands
   are runnable but not listed.)
6. **Graceful degradation.** Make the HKCU add-in keys unwritable (read-only
   ACL or a registry-restricted account) → a **warning** appears in the native
   log, **no** error dialog, and worksheet functions + the Ctrl+Shift
   shortcuts still work.
7. **Clean quit.** Quit Excel → clean exit (no crash; explicit COMAddIns
   disconnect in `xlAutoClose`); no orphan `xll_showcase.exe`; no leaked shared
   memory; HKCU ribbon keys removed.

**Worksheet functions, async, RTD (expansion):**

8. **Build the sheet.** Click **Build Showcase Sheet**. A formatted worksheet
   appears: bold title, dark-blue section headers with white text (sugar
   formatting), the function table with **live** results in column B, an inline
   numeric block (E4:F5) feeding `SumGrid`, and an instructions block.
9. **Sync types.** Verify `Add`→5, `Multiply`→6, `Greet`→`Hello, Excel!`,
   `IsEven`→TRUE, `SumGrid(E4:F5)`→10.
10. **Caller-aware.** `WhoAmI()` reports the address of its own cell.
11. **Volatile.** Note `RandomLine()`'s value, press **F9** → it changes (and
    the "Last recalc" cell updates from the `CalculationEnded` → `OnRecalc`
    handler, exercising the event + scheduling path).
12. **Async.** `SlowSquare(9)` shows a pending/`#N/A`-style value for ~1.5s,
    then resolves to **81** without blocking the sheet.
13. **RTD.** `Clock()` updates once per second; `StockTick("AAPL")` updates a
    wandering price several times per second — both live, streamed from the Go
    server over the in-proc RTD COM path.
14. **Clear.** Click **Clear Showcase** → the demo content is removed.

Excel-spawning automated tests MUST follow the two-tier cleanup rule: graceful
exit first, then force-kill on timeout — never just `defer`.
