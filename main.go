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

// IsEven reports whether n is even (sync bool).
func (s *Service) IsEven(ctx context.Context, n int32) (bool, error) {
	return n%2 == 0, nil
}

// Echo returns its argument unchanged (sync any -> any). The cell value
// arrives as a *protocol.Any read view; the handler returns a plain Go value
// and the generated server serializes it back (string/int32/float/bool keep
// their type; an empty or unreadable cell echoes as an empty cell).
func (s *Service) Echo(ctx context.Context, v *protocol.Any) (any, error) {
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
func (s *Service) BuildShowcaseSheet(ctx context.Context, cmd server.CommandContext) error {
	return runExcel("BuildShowcaseSheet", func(sheet excel.Worksheet) error {
		header := excel.RGB(0x1F, 0x4E, 0x79) // dark blue fill
		white := excel.RGB(0xFF, 0xFF, 0xFF)

		// Title
		title := sheet.Range("A1")
		title.SetValue("xll-gen Showcase — every feature, live")
		title.Font().SetBold(true).SetSize(14)

		// --- Worksheet Functions section ---
		sheet.Range("A3").SetValue("Worksheet Functions")
		fnHdr := sheet.Range("A3:C3")
		fnHdr.SetColor(header)
		fnHdr.Font().SetBold(true).SetColor(white)
		sheet.Range("A4").SetValue("Function")
		sheet.Range("B4").SetValue("Live Result")
		sheet.Range("C4").SetValue("Notes")
		sheet.Range("A4:C4").Font().SetBold(true)

		// Populate an inline numeric block for SumGrid to consume (E4:F5).
		sheet.Range("E4").SetValue(1.0)
		sheet.Range("F4").SetValue(2.0)
		sheet.Range("E5").SetValue(3.0)
		sheet.Range("F5").SetValue(4.0)

		rows := []struct{ label, formula, note string }{
			{"Add(2,3)", "=Add(2,3)", "sync int"},
			{"Multiply(1.5,4)", "=Multiply(1.5,4)", "sync float"},
			{`Greet("Excel")`, `=Greet("Excel")`, "sync string"},
			{"IsEven(10)", "=IsEven(10)", "sync bool"},
			{`Echo("dynamic!")`, `=Echo("dynamic!")`, "sync any -> any"},
			{"SumGrid(E4:F5)", "=SumGrid(E4:F5)", "grid -> 10"},
			{"WhoAmI()", "=WhoAmI()", "caller-aware"},
			{"RandomLine()", "=RandomLine()", "volatile (F9)"},
			{"SlowSquare(9)", "=SlowSquare(9)", "async, ~1.5s -> 81"},
			{"Clock()", "=Clock()", "RTD, ticks 1/s"},
			{`StockTick("AAPL")`, `=StockTick("AAPL")`, "RTD, wandering price"},
		}
		for i, r := range rows {
			row := 5 + i
			sheet.Range(fmt.Sprintf("A%d", row)).SetValue(r.label)
			sheet.Range(fmt.Sprintf("B%d", row)).SetFormula(r.formula)
			sheet.Range(fmt.Sprintf("C%d", row)).SetValue(r.note)
		}

		// --- Instructions section ---
		instrRow := 5 + len(rows) + 2
		sheet.Range(fmt.Sprintf("A%d", instrRow)).SetValue("Ribbon & Shortcuts")
		ih := sheet.Range(fmt.Sprintf("A%d:C%d", instrRow, instrRow))
		ih.SetColor(header)
		ih.Font().SetBold(true).SetColor(white)
		instructions := []string{
			"Ribbon tab 'xll-gen Showcase' -> Demo: Build Showcase Sheet / Clear Showcase",
			"Ribbon tab -> Commands: Write Timestamp / Slow Fill (5s) / Show Context",
			"Ctrl+Shift+T -> WriteTimestamp; Ctrl+Shift+S -> SlowFill",
			"Alt+F8 -> type a command name (e.g. ShowContext) -> Run",
			"Press F9 to recalc; RandomLine and the 'last recalc' cell update.",
		}
		for i, line := range instructions {
			sheet.Range(fmt.Sprintf("A%d", instrRow+1+i)).SetValue(line)
		}

		// 'last recalc' cell that the CalculationEnded handler updates.
		recalcRow := instrRow + 1 + len(instructions) + 1
		sheet.Range(fmt.Sprintf("A%d", recalcRow)).SetValue("Last recalc (set by event handler):")
		sheet.Range(fmt.Sprintf("B%d", recalcRow)).SetValue("press F9")

		return sheet.AutoFit()
	})
}

// ClearShowcase clears everything the demo wrote.
func (s *Service) ClearShowcase(ctx context.Context, cmd server.CommandContext) error {
	return runExcel("ClearShowcase", func(sheet excel.Worksheet) error {
		return sheet.Clear()
	})
}

// WriteTimestamp writes the current time to A1 (also bound to Ctrl+Shift+T).
func (s *Service) WriteTimestamp(ctx context.Context, cmd server.CommandContext) error {
	return runExcel("WriteTimestamp", func(sheet excel.Worksheet) error {
		return sheet.Range(showcaseAnchor).
			SetValue("WriteTimestamp @ " + time.Now().Format("15:04:05")).Err()
	})
}

// SlowFill sleeps 5s then writes "done". Excel stays fully responsive while it
// runs, proving the fire-and-forget STA contract (also bound to Ctrl+Shift+S).
func (s *Service) SlowFill(ctx context.Context, cmd server.CommandContext) error {
	time.Sleep(5 * time.Second)
	return runExcel("SlowFill", func(sheet excel.Worksheet) error {
		return sheet.Range(showcaseAnchor).
			SetValue("SlowFill done @ " + time.Now().Format("15:04:05")).Err()
	})
}

// ShowContext surfaces the invoking CommandContext into cells A1:B3, proving
// the ribbon control id / command name / Excel PID round-trip.
func (s *Service) ShowContext(ctx context.Context, cmd server.CommandContext) error {
	return runExcel("ShowContext", func(sheet excel.Worksheet) error {
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
// inside a sugar arena that releases every COM object on return. GetApplication
// can legitimately fail (no Excel, COM busy); we surface and log the error
// rather than panicking the worker.
func runExcel(name string, fn func(sheet excel.Worksheet) error) error {
	return sugar.Do(func(sctx sugar.Context) error {
		app := excel.GetApplication(sctx)
		if err := app.Err(); err != nil {
			return fmt.Errorf("%s: cannot attach to Excel: %w", name, err)
		}
		sheet := app.Books().Active().Sheets().Active()
		if err := sheet.Err(); err != nil {
			return fmt.Errorf("%s: no active sheet: %w", name, err)
		}
		if err := fn(sheet); err != nil {
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
		app := excel.GetApplication(sctx)
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
