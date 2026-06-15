//go:build windows

// Command xll_showcase is the Go server backing the xll-gen-showcase add-in.
// It implements every worksheet function, every RTD stream, and every ribbon
// command declared in xll.yaml. Worksheet functions are pure Go; command
// handlers drive Excel through the sugar COM-automation library.
package main

import (
	"context"
	"fmt"
	"math/rand"
	"net/url"
	"os"
	"strconv"
	"strings"
	"time"

	"github.com/xll-gen/sugar"
	"github.com/xll-gen/sugar/excel"
	"github.com/xll-gen/types/go/protocol"
	"github.com/xll-gen/xll-gen/pkg/server"
	flatbuffers "github.com/google/flatbuffers/go"

	"xll_showcase/generated"
)

// Service implements the generated XllService interface.
type Service struct{}

// ---------------------------------------------------------------------------
// Worksheet functions
// ---------------------------------------------------------------------------

// Add returns a + b (sync int).
func (s *Service) Add(ctx context.Context, a, b int32) (int32, error) {
	return a + b, nil
}

// Multiply returns a * b (sync float).
func (s *Service) Multiply(ctx context.Context, a, b float64) (float64, error) {
	return a * b, nil
}

// Greet returns a greeting (sync string).
func (s *Service) Greet(ctx context.Context, name string) (string, error) {
	if name == "" {
		name = "world"
	}
	return "Hello, " + name + "!", nil
}

// IsEvenInt reports whether n is even (sync bool). Named IsEvenInt, not
// IsEven: ISEVEN is a built-in Excel worksheet function that silently shadows
// an XLL registration of the same name.
func (s *Service) IsEvenInt(ctx context.Context, n int32) (bool, error) {
	return n%2 == 0, nil
}

// EchoAny returns its argument unchanged (sync any -> any). The cell value
// arrives as a *protocol.Any read view; the handler returns a plain Go value
// and the generated server serializes it back (string/int32/float/bool keep
// their type; an empty or unreadable cell echoes as an empty cell).
// Named EchoAny, not Echo: ECHO is a built-in Excel 4.0 (XLM) macro command,
// and Excel rejects worksheet formulas that use an XLM command name.
func (s *Service) EchoAny(ctx context.Context, v *protocol.Any) (any, error) {
	sv, ok := server.ToScalar(v)
	if !ok {
		return nil, nil // empty / missing cell -> empty cell
	}
	switch sv.Type {
	case protocol.AnyValueInt:
		return sv.Int, nil
	case protocol.AnyValueNum:
		return sv.Num, nil
	case protocol.AnyValueBool:
		return sv.Bool, nil
	case protocol.AnyValueStr:
		return sv.Str, nil
	case protocol.AnyValueErr:
		return fmt.Sprintf("#ERR(%d)", sv.Err), nil
	}
	return nil, nil
}

// SumGrid sums every numeric cell in the incoming grid (grid -> float).
func (s *Service) SumGrid(ctx context.Context, g *protocol.Grid) (float64, error) {
	if g == nil {
		return 0, nil
	}
	var total float64
	var cell protocol.Scalar
	for i := 0; i < g.DataLength(); i++ {
		if !g.Data(&cell, i) {
			continue
		}
		var tbl flatbuffers.Table
		if !cell.Val(&tbl) {
			continue
		}
		switch cell.ValType() {
		case protocol.ScalarValueNum:
			var num protocol.Num
			num.Init(tbl.Bytes, tbl.Pos)
			total += num.Val()
		case protocol.ScalarValueInt:
			var iv protocol.Int
			iv.Init(tbl.Bytes, tbl.Pos)
			total += float64(iv.Val())
		}
	}
	return total, nil
}

// ---------------------------------------------------------------------------
// Dynamic-array (spill) functions
//
// A handler that returns a 2-D array spills across the sheet on dynamic-array
// Excel (2021+/365): `numgrid` returns [][]float64 (dense, serialized as FP12)
// and `grid` returns [][]any (mixed nil/bool/string/int/float cells). The grid
// must be rectangular and non-empty; a jagged/empty grid surfaces as the
// function's error in the cell. On pre-DA Excel only the top-left cell appears
// unless the formula is entered as a CSE array over a selected range.
// ---------------------------------------------------------------------------

// TimesTable returns an n x n multiplication table (numgrid -> dense FP12).
// =TimesTable(5) spills a 5x5 grid where cell (r,c) = (r+1)*(c+1).
func (s *Service) TimesTable(ctx context.Context, n int32) ([][]float64, error) {
	if n < 1 {
		return nil, fmt.Errorf("n must be >= 1, got %d", n)
	}
	out := make([][]float64, n)
	for r := int32(0); r < n; r++ {
		row := make([]float64, n)
		for c := int32(0); c < n; c++ {
			row[c] = float64((r + 1) * (c + 1))
		}
		out[r] = row
	}
	return out, nil
}

// StatsGrid summarizes the numeric cells of the incoming grid and spills a
// 2-column label/value table (grid -> mixed string+number grid). It reads the
// grid argument through the same *protocol.Grid read view as SumGrid, then
// builds a plain [][]any for the return.
func (s *Service) StatsGrid(ctx context.Context, g *protocol.Grid) ([][]any, error) {
	var count int
	var sum, min, max float64
	if g != nil {
		var cell protocol.Scalar
		for i := 0; i < g.DataLength(); i++ {
			if !g.Data(&cell, i) {
				continue
			}
			var tbl flatbuffers.Table
			if !cell.Val(&tbl) {
				continue
			}
			var v float64
			switch cell.ValType() {
			case protocol.ScalarValueNum:
				var num protocol.Num
				num.Init(tbl.Bytes, tbl.Pos)
				v = num.Val()
			case protocol.ScalarValueInt:
				var iv protocol.Int
				iv.Init(tbl.Bytes, tbl.Pos)
				v = float64(iv.Val())
			default:
				continue
			}
			if count == 0 {
				min, max = v, v
			} else {
				if v < min {
					min = v
				}
				if v > max {
					max = v
				}
			}
			sum += v
			count++
		}
	}
	var mean float64
	if count > 0 {
		mean = sum / float64(count)
	}
	// Header row plus one row per statistic; mixes string labels with numbers.
	return [][]any{
		{"Statistic", "Value"},
		{"Count", count},
		{"Sum", sum},
		{"Mean", mean},
		{"Min", min},
		{"Max", max},
	}, nil
}

// SlowMatrix returns a rows x cols matrix after a ~1.5s delay (async numgrid).
// Proves an async result spills exactly like a sync grid: the XLL ACKs at once
// and the array streams back through the same conversion path. Cell (r,c) is
// encoded as r*cols + c so the spill layout is easy to eyeball.
func (s *Service) SlowMatrix(ctx context.Context, rows, cols int32) ([][]float64, error) {
	if rows < 1 || cols < 1 {
		return nil, fmt.Errorf("rows and cols must be >= 1, got %dx%d", rows, cols)
	}
	select {
	case <-ctx.Done():
		return nil, ctx.Err()
	case <-time.After(1500 * time.Millisecond):
	}
	out := make([][]float64, rows)
	for r := int32(0); r < rows; r++ {
		row := make([]float64, cols)
		for c := int32(0); c < cols; c++ {
			row[c] = float64(r*cols + c)
		}
		out[r] = row
	}
	return out, nil
}

// WhoAmI returns the address of the calling cell (caller-aware). The caller's
// reference arrives as a *protocol.Range whose first Rect carries the 0-based
// row/column span; we render it as an A1-style address.
func (s *Service) WhoAmI(ctx context.Context, caller *protocol.Range) (string, error) {
	if caller == nil || caller.RefsLength() == 0 {
		return "I was called from an unknown cell", nil
	}
	var rect protocol.Rect
	if !caller.Refs(&rect, 0) {
		return "I was called from an unknown cell", nil
	}
	addr := a1(rect.RowFirst(), rect.ColFirst())
	if rect.RowFirst() != rect.RowLast() || rect.ColFirst() != rect.ColLast() {
		addr += ":" + a1(rect.RowLast(), rect.ColLast())
	}
	sheet := string(caller.SheetName())
	if sheet != "" {
		return fmt.Sprintf("I was called from %s!%s", sheet, addr), nil
	}
	return "I was called from " + addr, nil
}

// a1 converts 0-based (row, col) to an A1-style address (e.g. 0,0 -> A1).
func a1(row, col int32) string {
	name := ""
	for c := col; ; c = c/26 - 1 {
		name = string(rune('A'+c%26)) + name
		if c < 26 {
			break
		}
	}
	return fmt.Sprintf("%s%d", name, row+1)
}

var quotes = []string{
	"Simplicity is the soul of efficiency.",
	"Make it work, make it right, make it fast.",
	"Talk is cheap. Show me the code.",
	"Premature optimization is the root of all evil.",
	"Programs must be written for people to read.",
}

// RandomLine returns a random quote. Declared volatile, so Excel recalculates
// it on every change / F9.
func (s *Service) RandomLine(ctx context.Context) (string, error) {
	return quotes[rand.Intn(len(quotes))], nil
}

// SlowSquare squares n after a ~1.5s delay (async). The XLL ACKs immediately
// and the result streams back when this returns.
func (s *Service) SlowSquare(ctx context.Context, n int32) (int32, error) {
	select {
	case <-ctx.Done():
		return 0, ctx.Err()
	case <-time.After(1500 * time.Millisecond):
	}
	return n * n, nil
}

// ---------------------------------------------------------------------------
// RTD streaming functions
// ---------------------------------------------------------------------------

// Clock_RTD streams the current time once per second until Excel disconnects.
func (s *Service) Clock_RTD(ctx context.Context, topicID int32) error {
	go func() {
		ticker := time.NewTicker(1 * time.Second)
		defer ticker.Stop()
		// Push an immediate first value so the cell isn't blank.
		_ = generated.PushRtdUpdate(topicID, time.Now().Format("15:04:05"))
		for {
			select {
			case <-ctx.Done():
				return
			case <-ticker.C:
				_ = generated.PushRtdUpdate(topicID, time.Now().Format("15:04:05"))
			}
		}
	}()
	return nil
}

// StockTick_RTD streams a wandering mock price for the given symbol.
func (s *Service) StockTick_RTD(ctx context.Context, topicID int32, symbol string) error {
	go func() {
		price := 100.0 + rand.Float64()*50.0
		ticker := time.NewTicker(750 * time.Millisecond)
		defer ticker.Stop()
		_ = generated.PushRtdUpdate(topicID, fmt.Sprintf("%s %.2f", symbol, price))
		for {
			select {
			case <-ctx.Done():
				return
			case <-ticker.C:
				price += (rand.Float64() - 0.5) * 2.0
				_ = generated.PushRtdUpdate(topicID, fmt.Sprintf("%s %.2f", symbol, price))
			}
		}
	}()
	return nil
}

// ---------------------------------------------------------------------------
// Live Market Data — Yahoo Finance (mode:"rtd-once")
//
// YDP and YDH are written as ORDINARY sync-shaped handlers — take a ctx, fetch,
// return a value — yet they run mode:"rtd-once": the cell shows #GETTING_DATA,
// the handler runs once OFF the calc thread (so Excel never blocks on the HTTP
// round-trip), and the cell settles when the fetch returns. YDH is the headline
// rtd-once-GRID demo: it returns a [][]any that SPILLS an OHLCV table, exercising
// the new grid guest->host + RTD-triggered re-spill machinery end to end.
// ---------------------------------------------------------------------------

// YDP — "Yahoo Data Point": one live quote field for a ticker
// (rtd-once any, memoize_ttl 60s). field is matched case-insensitively; numeric
// fields return float64, text fields (currency/exchange/time) return string.
func (s *Service) YDP(ctx context.Context, ticker, field string) (any, error) {
	res, err := fetchChart(ctx, ticker, url.Values{
		"interval": {"1d"},
		"range":    {"1d"},
	})
	if err != nil {
		return nil, err
	}
	m := &res.Meta

	// prevClose: chartPreviousClose is the series-relative previous close;
	// fall back to previousClose if Yahoo omits it.
	prevClose := m.ChartPreviousClose
	if prevClose == 0 {
		prevClose = m.PreviousClose
	}

	switch strings.ToLower(field) {
	case "price":
		return m.RegularMarketPrice, nil
	case "prevclose", "previousclose":
		return prevClose, nil
	case "change":
		return m.RegularMarketPrice - prevClose, nil
	case "change%", "changepct":
		if prevClose == 0 {
			return 0.0, nil
		}
		return (m.RegularMarketPrice - prevClose) / prevClose * 100, nil
	case "currency":
		return m.Currency, nil
	case "exchange":
		return m.ExchangeName, nil
	case "high":
		return m.RegularMarketDayHigh, nil
	case "low":
		return m.RegularMarketDayLow, nil
	case "volume":
		return m.RegularMarketVolume, nil
	case "time":
		if m.RegularMarketTime == 0 {
			return "", nil
		}
		t := time.Unix(m.RegularMarketTime, 0)
		// A same-day quote reads best as a wall-clock time; anything older
		// keeps a full RFC3339 stamp so the date is not lost.
		if time.Since(t) < 24*time.Hour {
			return t.Format("15:04:05"), nil
		}
		return t.Format(time.RFC3339), nil
	case "high52":
		return m.FiftyTwoWeekHigh, nil
	case "low52":
		return m.FiftyTwoWeekLow, nil
	default:
		return nil, fmt.Errorf("unknown field %q", field)
	}
}

// YDH — "Yahoo Data History": spills a historical OHLCV table for the last
// `days` calendar days (rtd-once GRID, memoize_ttl 60s). The returned [][]any is
// a header row {Date,Open,High,Low,Close,Volume} followed by one row per trading
// bar; Yahoo emits nulls for non-trading days (holidays) which are skipped. This
// is the end-to-end demo of the rtd-once grid-spill feature: a non-blocking
// network fetch whose result spills via the new guest->host grid path.
//
// The Date column is returned as a real time.Time carrying the bar's market
// timestamp (the time-of-day is kept intentionally, not normalized to
// midnight), so xll-gen serializes it to an Excel date-time serial and
// auto-formats the data cells as yyyy-mm-dd hh:mm:ss at calc-end —
// value-driven, per-cell formatting (the header string and the numeric OHLCV
// columns are left untouched). The values are real Excel dates, sortable and
// usable in date arithmetic, not text.
func (s *Service) YDH(ctx context.Context, ticker string, days int32) ([][]any, error) {
	if days < 1 {
		return nil, fmt.Errorf("days must be >= 1, got %d", days)
	}
	if days > 3650 {
		days = 3650 // clamp to ~10y so an accidental huge value stays sane
	}

	now := time.Now()
	p2 := now.Unix()
	p1 := now.Add(-time.Duration(days) * 24 * time.Hour).Unix()

	res, err := fetchChart(ctx, ticker, url.Values{
		"period1":  {strconv.FormatInt(p1, 10)},
		"period2":  {strconv.FormatInt(p2, 10)},
		"interval": {"1d"},
	})
	if err != nil {
		return nil, err
	}
	if len(res.Indicators.Quote) == 0 {
		return nil, fmt.Errorf("no OHLCV data for %q", ticker)
	}
	q := res.Indicators.Quote[0]

	grid := [][]any{{"Date", "Open", "High", "Low", "Close", "Volume"}}
	for i, ts := range res.Timestamp {
		// Guard against ragged arrays and skip null (holiday/no-trade) bars:
		// Yahoo nulls every OHLCV field together on a non-trading day.
		if i >= len(q.Open) || i >= len(q.High) || i >= len(q.Low) ||
			i >= len(q.Close) || i >= len(q.Volume) {
			break
		}
		if q.Open[i] == nil || q.High[i] == nil || q.Low[i] == nil ||
			q.Close[i] == nil || q.Volume[i] == nil {
			continue
		}
		// Emit the bar date as a real time.Time, keeping the market time-of-day
		// (not normalized to midnight): xll-gen serializes it to an Excel
		// date-time serial and auto-formats this column as yyyy-mm-dd hh:mm:ss
		// at calc-end (value-driven — only this date column is formatted; the
		// numeric OHLCV columns and the "Date" header string are left as-is).
		date := time.Unix(ts, 0)
		grid = append(grid, []any{
			date,
			*q.Open[i], *q.High[i], *q.Low[i], *q.Close[i], *q.Volume[i],
		})
	}
	if len(grid) == 1 {
		return nil, fmt.Errorf("no trading bars for %q in the last %d days", ticker, days)
	}
	return grid, nil
}

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
		tableBlock := make([][]any, len(rows))   // A:C values, col B blank
		formulaCol := make([][]any, len(rows))   // B column, one formula/cell
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
		bdpBlock := make([][]any, len(bdpRows))  // A:C labels/notes, col B blank
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

// ---------------------------------------------------------------------------
// Events
// ---------------------------------------------------------------------------

// OnRecalc fires after every Excel calculation cycle (CalculationEnded). It
// schedules a write of the recalc timestamp via the command-scheduling path so
// the write happens on Excel's thread at a safe point.
func (s *Service) OnRecalc(ctx context.Context) error {
	_ = sugar.Do(func(sctx sugar.Context) error {
		// Events carry no CommandContext, but the Go server is launched as
		// Excel's child, so the hosting Excel's PID is our parent PID — attach
		// by it (same window-walk reason as the command handlers; the ROT is
		// unreachable from this process).
		app := attachExcel(sctx, uint32(os.Getppid()))
		if err := app.Err(); err != nil {
			return err
		}
		sheet := app.Books().Active().Sheets().Active()
		// Find the marker cell written by BuildShowcaseSheet and stamp B next to it.
		cell, found, err := sheet.UsedRange().Find("Last recalc (set by event handler):")
		if err != nil || !found {
			return err
		}
		row, err := cell.Row()
		if err != nil {
			return err
		}
		return sheet.Range(fmt.Sprintf("B%d", row)).
			SetValue("recalc @ " + time.Now().Format("15:04:05")).Err()
	})
	return nil
}

func (s *Service) OnCalculationCanceled(ctx context.Context) error { return nil }

// ---------------------------------------------------------------------------
// RTD lifecycle
// ---------------------------------------------------------------------------

func (s *Service) OnRtdConnect(ctx context.Context, topicID int32, strings []string, newValues bool) error {
	return nil
}

func (s *Service) OnRtdDisconnect(ctx context.Context, topicID int32) error { return nil }

func main() {
	generated.Serve(&Service{})
}
