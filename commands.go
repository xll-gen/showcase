//go:build windows

package main

import (
	"context"
	"fmt"
	"time"

	"github.com/xll-gen/sugar"
	"github.com/xll-gen/sugar/excel"
	"github.com/xll-gen/xll-gen/pkg/server"
)

// ---------------------------------------------------------------------------
// Commands (sugar-driven Excel automation)
//
// Each handler runs on the Go server's worker pool, fire-and-forget from
// Excel's point of view. sugar.Do sets up an STA-correct COM apartment and an
// arena that releases every tracked COM object (the Application and all
// objects reached through it) when the callback returns — the real
// equivalent of "defer ctx.Release()".
// ---------------------------------------------------------------------------

const showcaseAnchor = "A1" // top-left of the demo block

// BuildShowcaseSheet is the centerpiece: it lays out a worksheet that exercises
// every worksheet function and lists the ribbon/shortcut instructions.
//
// Speed: the build is dominated by cross-process COM round-trips, so it (1)
// suspends ScreenUpdating + switches Calculation to manual for the whole build
// (restored, with a single recalc, in a guaranteed cleanup path), and (2)
// writes each contiguous area as ONE block — the A:C table values in a single
// SetValue, the B-column formulas in a single SetFormula2Array, the E:F input
// block, the H labels, and the instructions/recalc rows — instead of cell by
// cell. The three spill anchors stay single-cell Formula2 writes (a spill
// formula lives in one anchor). The layout is byte-for-byte the same as the
// per-cell version.
func (s *Service) BuildShowcaseSheet(ctx context.Context, cmd server.CommandContext) error {
	return runExcelApp("BuildShowcaseSheet", cmd.ExcelPID, func(app excel.Application, sheet excel.Worksheet) (err error) {
		header := excel.RGB(0x1F, 0x4E, 0x79) // dark blue fill
		white := excel.RGB(0xFF, 0xFF, 0xFF)

		// --- Suspend UI + recalc for the whole build ---------------------
		// Capture the live settings so we restore exactly what the user had,
		// then force ScreenUpdating off and Calculation to manual. Without
		// this, every formula write triggers a recalc/repaint mid-build.
		// The cleanup runs on EVERY return path (including errors and panics
		// in the writes below) so the user's Excel is never left in manual
		// calc / no-repaint — that is a real support nightmare. A single
		// Application.Calculate at the end refreshes all the formulas we wrote.
		prevCalc, calcErr := app.Calculation()
		prevScreen, screenErr := app.ScreenUpdating()
		if calcErr == nil {
			app.SetCalculation(excel.CalculationManual)
		}
		if screenErr == nil {
			app.SetScreenUpdating(false)
		}
		defer func() {
			// Restore in reverse order; recalc once so written formulas
			// evaluate. Restore errors do not mask a build error.
			if calcErr == nil {
				app.SetCalculation(prevCalc)
			} else {
				app.SetCalculation(excel.CalculationAutomatic)
			}
			app.Call("Calculate")
			if screenErr == nil {
				app.SetScreenUpdating(prevScreen)
			} else {
				app.SetScreenUpdating(true)
			}
		}()

		// --- Title -------------------------------------------------------
		title := sheet.Range("A1")
		if err = title.SetValue("xll-gen Showcase — every feature, live").Err(); err != nil {
			return err
		}
		title.Font().SetBold(true).SetSize(14)

		// --- Worksheet Functions section --------------------------------
		// Header rows A3:C4 written as one 2x3 block (col B/C of the title
		// row left blank), then the fill/bold formatting applied per banner.
		if err = sheet.Range("A3:C4").SetValue([][]any{
			{"Worksheet Functions", nil, nil},
			{"Function", "Live Result", "Notes"},
		}).Err(); err != nil {
			return err
		}
		fnHdr := sheet.Range("A3:C3")
		fnHdr.SetColor(header)
		fnHdr.Font().SetBold(true).SetColor(white)
		sheet.Range("A4:C4").Font().SetBold(true)

		// Inline numeric block for SumGrid/StatsGrid (E4:F5), one block write.
		// Columns E-F sit clear of the A-C table and the H+ spill area.
		if err = sheet.Range("E4:F5").SetValue([][]any{
			{1.0, 2.0},
			{3.0, 4.0},
		}).Err(); err != nil {
			return err
		}

		// Scalar-result functions only. The three array-returning (spill)
		// functions live in their own area (columns H+) so their spill ranges
		// never collide with this single-column-result table. The whole A:C
		// table is one block SetValue (labels in A, blanks in B, notes in C);
		// the B-column formulas are then one SetFormula2Array — Formula2-native
		// so dynamic-array Excel does not rewrite UDF calls into the
		// implicit-intersection `=@Fn(...)` form.
		rows := []struct{ label, formula, note string }{
			{"Add(2,3)", "=Add(2,3)", "sync int"},
			{"Multiply(1.5,4)", "=Multiply(1.5,4)", "sync float"},
			{`Greet("Excel")`, `=Greet("Excel")`, "sync string"},
			{"IsEvenInt(10)", "=IsEvenInt(10)", "sync bool"},
			{`EchoAny("dynamic!")`, `=EchoAny("dynamic!")`, "sync any -> any"},
			{"SumGrid(E4:F5)", "=SumGrid(E4:F5)", "grid -> 10"},
			{"WhoAmI()", "=WhoAmI()", "caller-aware"},
			{"RandomLine()", "=RandomLine()", "volatile (F9)"},
			{"SlowSquare(9)", "=SlowSquare(9)", "async, ~1.5s -> 81"},
			{"Clock()", "=Clock()", "RTD, ticks 1/s"},
			{`StockTick("AAPL")`, `=StockTick("AAPL")`, "RTD, wandering price"},
		}
		const firstRow = 5
		tableBlock := make([][]any, len(rows)) // A:C values, col B blank
		formulaCol := make([][]any, len(rows)) // B column, one formula/cell
		for i, r := range rows {
			tableBlock[i] = []any{r.label, nil, r.note}
			formulaCol[i] = []any{r.formula}
		}
		lastRow := firstRow + len(rows) - 1 // 15
		if err = sheet.Range(fmt.Sprintf("A%d:C%d", firstRow, lastRow)).
			SetValue(tableBlock).Err(); err != nil {
			return err
		}
		if err = sheet.Range(fmt.Sprintf("B%d:B%d", firstRow, lastRow)).
			SetFormula2Array(formulaCol).Err(); err != nil {
			return err
		}

		// --- Dynamic Arrays (spill) section -----------------------------
		//
		// Each spill demo gets its FULL spill extent reserved plus blank-row
		// margins, all in columns H..M so nothing here can trample the A-C
		// function table or the E-F SumGrid block. Layout (1 blank margin row
		// between demos; the anchor is one row below its label):
		//
		//   H3:M3   section header
		//   H5      label  "TimesTable(5) — numgrid, spills 5x5"
		//   H6      anchor =TimesTable(5)   -> spills H6:L10   (5 rows x 5 cols)
		//   (row 11 blank margin)
		//   H12     label  "StatsGrid(E4:F5) — grid, spills 6x2"
		//   H13     anchor =StatsGrid(E4:F5) -> spills H13:I18 (6 rows x 2 cols)
		//   (row 19 blank margin)
		//   H20     label  "SlowMatrix(3,3) — async numgrid, spills 3x3"
		//   H21     anchor =SlowMatrix(3,3)  -> spills H21:J23 (3 rows x 3 cols)
		//   (row 24 blank margin)
		if err = sheet.Range("H3").SetValue("Dynamic Arrays (spill)").Err(); err != nil {
			return err
		}
		spHdr := sheet.Range("H3:M3")
		spHdr.SetColor(header)
		spHdr.Font().SetBold(true).SetColor(white)

		spills := []struct {
			labelRow, anchorRow int
			label, formula      string
		}{
			{5, 6, "TimesTable(5) — numgrid, spills 5x5", "=TimesTable(5)"},
			{12, 13, "StatsGrid(E4:F5) — grid, spills 6x2", "=StatsGrid(E4:F5)"},
			{20, 21, "SlowMatrix(3,3) — async numgrid, spills 3x3", "=SlowMatrix(3,3)"},
		}
		// Labels at H5, H12, H20 written as ONE sparse column block H5:H20
		// (intervening rows nil); the anchor formulas (H6/H13/H21) are written
		// separately below, so nilling them here is harmless. Bold is applied
		// per-label since the anchors between them must stay non-bold.
		labelBlock := make([][]any, spills[len(spills)-1].labelRow-spills[0].labelRow+1)
		for i := range labelBlock {
			labelBlock[i] = []any{nil}
		}
		for _, sp := range spills {
			labelBlock[sp.labelRow-spills[0].labelRow][0] = sp.label
		}
		if err = sheet.Range(fmt.Sprintf("H%d:H%d", spills[0].labelRow, spills[len(spills)-1].labelRow)).
			SetValue(labelBlock).Err(); err != nil {
			return err
		}
		for _, sp := range spills {
			sheet.Range(fmt.Sprintf("H%d", sp.labelRow)).Font().SetBold(true)
			// One spill formula per anchor cell — stays a single-cell Formula2
			// write (a spill anchor is one cell, not a block).
			if err = sheet.Range(fmt.Sprintf("H%d", sp.anchorRow)).
				SetFormulaSpill(sp.formula).Err(); err != nil {
				return err
			}
		}

		// --- Instructions section ---------------------------------------
		// Banner + 5 lines written as one A-column block (A18:A23).
		instrRow := firstRow + len(rows) + 2 // 18
		instructions := []string{
			"Ribbon tab 'xll-gen Showcase' -> Demo: Build Showcase Sheet / Clear Showcase",
			"Ribbon tab -> Commands: Write Timestamp / Slow Fill (5s) / Show Context",
			"Ctrl+Shift+T -> WriteTimestamp; Ctrl+Shift+S -> SlowFill",
			"Alt+F8 -> type a command name (e.g. ShowContext) -> Run",
			"Press F9 to recalc; RandomLine and the 'last recalc' cell update.",
		}
		instrBlock := make([][]any, len(instructions)+1)
		instrBlock[0] = []any{"Ribbon & Shortcuts"}
		for i, line := range instructions {
			instrBlock[i+1] = []any{line}
		}
		instrLast := instrRow + len(instructions) // 23
		if err = sheet.Range(fmt.Sprintf("A%d:A%d", instrRow, instrLast)).
			SetValue(instrBlock).Err(); err != nil {
			return err
		}
		ih := sheet.Range(fmt.Sprintf("A%d:C%d", instrRow, instrRow))
		ih.SetColor(header)
		ih.Font().SetBold(true).SetColor(white)

		// 'last recalc' marker + value, one A:B block.
		recalcRow := instrRow + 1 + len(instructions) + 1 // 25
		if err = sheet.Range(fmt.Sprintf("A%d:B%d", recalcRow, recalcRow)).
			SetValue([][]any{{"Last recalc (set by event handler):", "press F9"}}).Err(); err != nil {
			return err
		}

		// Capture the marker location for OnRecalc. Reading sheet.Name() here is
		// safe: BuildShowcaseSheet runs as a COMMAND (STA is free), unlike the
		// event handler, which must never touch COM (see OnRecalc). OnRecalc
		// stamps the timestamp into column B at this row via generated.ScheduleSet,
		// so it never needs to Find the marker at runtime.
		sheetName, err := sheet.Name()
		if err != nil {
			return err
		}
		s.recalcMu.Lock()
		s.recalcSheet = sheetName
		s.recalcRow = recalcRow
		s.recalcMu.Unlock()

		// --- Live Market Data — Yahoo Finance (rtd-once) -----------------
		//
		// Placed well below everything else so the YDH spill (an OHLCV table,
		// up to ~30 rows x 6 cols) has clear runway and never collides with the
		// A:C table, the E:F SumGrid block, or the H:M spill demos above.
		//
		//   A28:F28   section header
		//   A30       label  "YDP — one live quote field (rtd-once, settles from #GETTING_DATA)"
		//   A31:C33   YDP rows: label | =YDP(...) | note   (scalar rtd-once -> Formula2)
		//   A36       label  "YDH('MSFT',30) — rtd-once GRID: spills an OHLCV table from #GETTING_DATA"
		//   A37       anchor =YDH("MSFT",30)  -> spills A37:F~58 (header + ~21 bars)
		const finHdrRow = 28
		if err = sheet.Range(fmt.Sprintf("A%d:F%d", finHdrRow, finHdrRow)).
			SetValue([][]any{{"Live Market Data — Yahoo Finance (rtd-once)", nil, nil, nil, nil, nil}}).Err(); err != nil {
			return err
		}
		finHdr := sheet.Range(fmt.Sprintf("A%d:F%d", finHdrRow, finHdrRow))
		finHdr.SetColor(header)
		finHdr.Font().SetBold(true).SetColor(white)

		// YDP scalar demos. Each is a normal single-cell Formula2 write (YDP is
		// a scalar rtd-once: it settles in place, no spill).
		const bdpLabelRow = 30
		if err = sheet.Range(fmt.Sprintf("A%d", bdpLabelRow)).
			SetValue("YDP — one live quote field (rtd-once; settles from #GETTING_DATA)").Err(); err != nil {
			return err
		}
		sheet.Range(fmt.Sprintf("A%d", bdpLabelRow)).Font().SetBold(true)

		bdpRows := []struct{ label, formula, note string }{
			{`YDP("AAPL","price")`, `=YDP("AAPL","price")`, "live price (float)"},
			{`YDP("AAPL","change%")`, `=YDP("AAPL","change%")`, "% change vs prev close"},
			{`YDP("AAPL","currency")`, `=YDP("AAPL","currency")`, "currency (string)"},
		}
		const bdpFirst = 31
		bdpBlock := make([][]any, len(bdpRows))   // A:C labels/notes, col B blank
		bdpFormula := make([][]any, len(bdpRows)) // B column formulas
		for i, r := range bdpRows {
			bdpBlock[i] = []any{r.label, nil, r.note}
			bdpFormula[i] = []any{r.formula}
		}
		bdpLast := bdpFirst + len(bdpRows) - 1 // 33
		if err = sheet.Range(fmt.Sprintf("A%d:C%d", bdpFirst, bdpLast)).
			SetValue(bdpBlock).Err(); err != nil {
			return err
		}
		if err = sheet.Range(fmt.Sprintf("B%d:B%d", bdpFirst, bdpLast)).
			SetFormula2Array(bdpFormula).Err(); err != nil {
			return err
		}

		// YDH grid demo — the headline rtd-once-grid spill. Single-cell spill
		// anchor; on entry the cell shows #GETTING_DATA, then spills the OHLCV
		// table downward once the off-thread fetch returns.
		const bdhLabelRow = 36
		const bdhAnchorRow = 37
		if err = sheet.Range(fmt.Sprintf("A%d", bdhLabelRow)).
			SetValue(`YDH("MSFT",30) — rtd-once GRID: shows #GETTING_DATA, then spills a Date/OHLCV table`).Err(); err != nil {
			return err
		}
		sheet.Range(fmt.Sprintf("A%d", bdhLabelRow)).Font().SetBold(true)
		if err = sheet.Range(fmt.Sprintf("A%d", bdhAnchorRow)).
			SetFormulaSpill(`=YDH("MSFT",30)`).Err(); err != nil {
			return err
		}

		return sheet.AutoFit()
	})
}

// ClearShowcase clears everything the demo wrote.
func (s *Service) ClearShowcase(ctx context.Context, cmd server.CommandContext) error {
	return runExcel("ClearShowcase", cmd.ExcelPID, func(sheet excel.Worksheet) error {
		return sheet.Clear()
	})
}

// WriteTimestamp writes the current time to A1 (also bound to Ctrl+Shift+T).
func (s *Service) WriteTimestamp(ctx context.Context, cmd server.CommandContext) error {
	return runExcel("WriteTimestamp", cmd.ExcelPID, func(sheet excel.Worksheet) error {
		return sheet.Range(showcaseAnchor).
			SetValue("WriteTimestamp @ " + time.Now().Format("15:04:05")).Err()
	})
}

// SlowFill sleeps 5s then writes "done". Excel stays fully responsive while it
// runs, proving the fire-and-forget STA contract (also bound to Ctrl+Shift+S).
func (s *Service) SlowFill(ctx context.Context, cmd server.CommandContext) error {
	time.Sleep(5 * time.Second)
	return runExcel("SlowFill", cmd.ExcelPID, func(sheet excel.Worksheet) error {
		return sheet.Range(showcaseAnchor).
			SetValue("SlowFill done @ " + time.Now().Format("15:04:05")).Err()
	})
}

// ShowContext surfaces the invoking CommandContext into cells A1:B3, proving
// the ribbon control id / command name / Excel PID round-trip.
func (s *Service) ShowContext(ctx context.Context, cmd server.CommandContext) error {
	return runExcel("ShowContext", cmd.ExcelPID, func(sheet excel.Worksheet) error {
		sheet.Range("A1").SetValue("CommandName")
		sheet.Range("B1").SetValue(cmd.CommandName)
		sheet.Range("A2").SetValue("ControlID")
		sheet.Range("B2").SetValue(cmd.ControlID)
		sheet.Range("A3").SetValue("ExcelPID")
		sheet.Range("B3").SetValue(int32(cmd.ExcelPID))
		return sheet.Range("A1:B3").Err()
	})
}

// runExcel is the shared scaffold for command handlers. It attaches to the
// running Excel instance, resolves the active worksheet, and runs fn — all
// inside a sugar arena that releases every COM object on return. attachExcel
// can legitimately fail (no Excel, COM busy); we surface and log the error
// rather than panicking the worker.
func runExcel(name string, excelPID uint32, fn func(sheet excel.Worksheet) error) error {
	return runExcelApp(name, excelPID, func(_ excel.Application, sheet excel.Worksheet) error {
		return fn(sheet)
	})
}

// attachExcel attaches to the EXACT Excel instance that invoked the command.
// The Go server runs as Excel's child process, so the Running Object Table is
// not reachable from its COM apartment — the ROT-based excel.GetApplication
// fails with "작업을 사용할 수 없습니다" (MK_E_UNAVAILABLE). The command carries the
// hosting Excel's PID (CommandContext.ExcelPID), so we attach by PID via the
// XLMAIN->XLDESK->EXCEL7 window walk. We fall back to the ROT path only when no
// PID is available (PID 0), so shortcut/Alt+F8 paths that lack one still work
// in single-instance setups.
func attachExcel(sctx sugar.Context, excelPID uint32) excel.Application {
	if excelPID != 0 {
		app := excel.GetApplicationByPID(sctx, excelPID)
		if app.Err() == nil {
			return app
		}
		// Fall through to the ROT path as a best-effort backup (e.g. window
		// chain transiently unavailable); the original error is preserved if
		// that also fails.
	}
	return excel.GetApplication(sctx)
}

// runExcelApp is runExcel's app-aware variant: fn also receives the
// Application so a handler can toggle application-wide settings
// (ScreenUpdating / Calculation) for the duration of its work. Everything
// still runs inside the same sugar arena that releases every COM object on
// return.
func runExcelApp(name string, excelPID uint32, fn func(app excel.Application, sheet excel.Worksheet) error) error {
	return sugar.Do(func(sctx sugar.Context) error {
		app := attachExcel(sctx, excelPID)
		if err := app.Err(); err != nil {
			return fmt.Errorf("%s: cannot attach to Excel: %w", name, err)
		}
		sheet := app.Books().Active().Sheets().Active()
		if err := sheet.Err(); err != nil {
			return fmt.Errorf("%s: no active sheet: %w", name, err)
		}
		if err := fn(app, sheet); err != nil {
			return fmt.Errorf("%s: %w", name, err)
		}
		return nil
	})
}
