//go:build windows

package main

import (
	"context"
	"fmt"
	"math/rand"
	"time"

	flatbuffers "github.com/google/flatbuffers/go"
	"github.com/xll-gen/types/go/protocol"
)

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
