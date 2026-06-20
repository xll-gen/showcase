//go:build windows

package main

import (
	"context"
	"fmt"

	flatbuffers "github.com/google/flatbuffers/go"
	"github.com/xll-gen/types/go/protocol"
	"github.com/xll-gen/xll-gen/pkg/server"
)

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

// rangeNumericCells invokes fn for every numeric (Int or Num) cell of g,
// decoding the raw FlatBuffers Scalar union once. Non-numeric, empty, or
// unreadable cells are skipped; a nil grid yields no calls. This is the one
// place the showcase touches the on-wire grid encoding — SumGrid and StatsGrid
// both consume grids through it rather than re-walking the union.
func rangeNumericCells(g *protocol.Grid, fn func(v float64)) {
	if g == nil {
		return
	}
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
			fn(num.Val())
		case protocol.ScalarValueInt:
			var iv protocol.Int
			iv.Init(tbl.Bytes, tbl.Pos)
			fn(float64(iv.Val()))
		}
	}
}

// SumGrid sums every numeric cell in the incoming grid (grid -> float).
func (s *Service) SumGrid(ctx context.Context, g *protocol.Grid) (float64, error) {
	var total float64
	rangeNumericCells(g, func(v float64) { total += v })
	return total, nil
}
