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

The `xll-gen` CLI itself is built from the local checkout (see below). This
showcase **builds end-to-end with xll-gen ≥ v0.4.1**, which fixed the two C++
codegen defects (`grid` argument + `caller`-aware) that previously blocked the
`.xll`. Module dependencies are public tagged releases (`xll-gen v0.4.2`,
`sugar v0.8.0`, `types v0.2.8`, `shm v0.7.3`) — no `replace` directives.

> **Note:** the ribbon **PNG file icons** (`image: ./icons/*.png`) need a CLI
> built from xll-gen **after v0.4.2** (PR #331, on `main`). An older CLI emits
> the path as a (broken) `imageMso` name — the add-in still builds and runs,
> the buttons just lose their icons. The Go *module* dependency is unaffected.

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
- **C++ XLL: builds end-to-end with xll-gen ≥ v0.4.1.** CMake configures and
  the build reaches 100% — every static dep, all sources, and `xll_main.cpp`
  compile, and the linker emits `build/xll_showcase.xll`. The two functions
  that v0.4.0 mis-generated are now correct:
  - `SumGrid` (a `grid` argument) — the generator now emits
    `ConvertGrid(g, builder)` (the function the types library actually
    provides), not the non-existent `GridToFlatBuffer(...)`.
  - `WhoAmI` (`caller: true`) — the caller-resolution code now uses
    `ScopedXLOPER12` correctly: `.get()` is passed to `xll::CallExcel` /
    `ConvertRange` (which take `LPXLOPER12`) and members are read via
    `.get()->`, so the type mismatches are gone.

  Both were defects in xll-gen v0.4.0's **C++** templates for the `grid`-arg
  and `caller`-aware paths; xll-gen #328 (v0.4.1) fixed them. No feature errors
  remain — sync/volatile/async/RTD/commands/ribbon/event all compile and the
  add-in links to a runnable `.xll`.

## Feature coverage

Every declared feature, and where to observe it. After clicking **Build
Showcase Sheet** the live cells appear in column B of the demo worksheet.

| Feature | Declared as | Observe via |
|---------|-------------|-------------|
| sync int | `Add(a,b int) -> int` | cell `=Add(2,3)` → 5 |
| sync float | `Multiply(a,b float) -> float` | cell `=Multiply(1.5,4)` → 6 |
| sync string | `Greet(name string) -> string` | cell `=Greet("Excel")` |
| sync bool | `IsEvenInt(n int) -> bool` | cell `=IsEvenInt(10)` → TRUE (not `IsEven` — the built-in `ISEVEN` silently shadows an XLL function of the same name) |
| sync `any` round-trip | `EchoAny(v any) -> any` | cell `=EchoAny("dynamic!")` echoes any scalar (not `Echo` — `ECHO` is an XLM macro command; Excel rejects the formula outright) |
| `grid` arg (2-D `any` cells) | `SumGrid(g grid) -> float` | cell `=SumGrid(E4:F5)` → 10 |
| spill `numgrid` (dense FP12) | `TimesTable(n int) -> numgrid` | cell `=TimesTable(5)` → spills a 5×5 multiplication table |
| spill `grid` (mixed cells) | `StatsGrid(g grid) -> grid` | cell `=StatsGrid(E4:F5)` → spills a 2-col Count/Sum/Mean/Min/Max summary |
| spill async `numgrid` | `SlowMatrix(rows,cols int) -> numgrid` (`mode: async`) | cell `=SlowMatrix(3,3)`; ~1.5s then spills a 3×3 matrix |
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
| ribbon PNG file icons | `image: ./icons/build.png` (32×32, large) and `./icons/clear.png` (16×16) | **Demo** buttons render embedded PNG icons with transparent corners (GDI+ loadImage path; requires xll-gen > v0.4.2 — regenerate icons via `go run tools/gen_icons.go`) |
| event + scheduling | `CalculationEnded` → `OnRecalc` | press **F9**; "Last recalc" cell stamps a time |

### Formerly omitted: sync `return: "any"`

Earlier revisions omitted a **sync** function with `return: "any"`: xll-gen
v0.4.0–v0.4.1's Go generator emitted `ipc.EchoResponseAddResult(b, res)`,
passing a `*protocol.Any` where the FlatBuffers setter expects a
`flatbuffers.UOffsetT` — non-compiling generated code. The generator now maps
`return: "any"` to a plain Go `any` handler return and serializes it through
the same canonical Go-value→`protocol.Any` mapping the RTD path uses, so
`EchoAny(v any) -> any` above covers the last `any` position (argument, RTD
return, and now sync return).

### Dynamic arrays (spill)

`TimesTable`, `StatsGrid`, and `SlowMatrix` demo xll-gen's **spill** support: a
sync or async function may return a 2-D array that Excel spills across the cells
below/right of the formula cell.

- **`numgrid`** (`[][]float64`, dense, serialized as FP12): `TimesTable(n)` and
  the async `SlowMatrix(rows,cols)`.
- **`grid`** (`[][]any`, mixed `nil`/`bool`/`string`/int/float cells):
  `StatsGrid(g)` consumes a `grid` argument (same `*protocol.Grid` read view as
  `SumGrid`) and returns a 2-column label/value summary mixing string labels
  with numbers.

Type `=TimesTable(5)` in one cell and watch it **spill** into a 5×5 block
(`(r+1)*(c+1)`). `=SlowMatrix(3,3)` shows that an **async** result spills the
same way — the XLL ACKs immediately and the array streams back through the
identical conversion path.

> **Excel-version behavior:** on **dynamic-array Excel (2021+/365)** these
> spill automatically — no registration flag is required (return code `Q` for
> `grid`, `K%`/FP12 for `numgrid`). On **pre-dynamic-array Excel** the formula
> returns only the **top-left cell**; to see the whole array, pre-select a
> range and enter the formula as a legacy **CSE array** (`Ctrl+Shift+Enter`).
> A jagged or empty grid is reported as the function's error in the cell. The
> `range` type is **not** supported as a return (`U`-coded returns break Excel
> registration — use `grid`/`numgrid` for spilling arrays).

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
   **PNG icons render:** Build Showcase Sheet shows a green ⊕ (32×32) and
   Clear Showcase a red ⊗ (16×16), both with *transparent* corners — the
   circle edge must blend into the ribbon background, not sit in a white/black
   square (verifies the alpha-preserving GDI+ `loadImage` path).
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
   `IsEvenInt`→TRUE, `EchoAny`→`dynamic!`, `SumGrid(E4:F5)`→10.
9b. **Spill (dynamic arrays).** On Excel 2021+/365, `=TimesTable(5)` spills a
    5×5 multiplication table and `=StatsGrid(E4:F5)` spills a 2-column
    Count/Sum/Mean/Min/Max summary into the cells below/right of the formula
    cell (a blue spill border on selection). `=SlowMatrix(3,3)` spills a 3×3
    matrix ~1.5s after entry — proving async results spill too. On pre-DA Excel
    each shows only its top-left cell unless entered as a CSE array.
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
