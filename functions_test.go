//go:build windows

package main

import (
	"context"
	"math"
	"testing"

	flatbuffers "github.com/google/flatbuffers/go"
	"github.com/xll-gen/types/go/protocol"
)

// buildScalarCell serializes one Int/Num/Str/nil cell into a protocol.Scalar,
// mirroring the on-wire encoding xll-gen produces for grid arguments. Only the
// cell kinds the grid-decode path cares about are covered.
func buildScalarCell(b *flatbuffers.Builder, c any) flatbuffers.UOffsetT {
	var valType protocol.ScalarValue
	var uOff flatbuffers.UOffsetT
	switch v := c.(type) {
	case nil:
		protocol.NilStart(b)
		uOff = protocol.NilEnd(b)
		valType = protocol.ScalarValueNil
	case int:
		protocol.IntStart(b)
		protocol.IntAddVal(b, int32(v))
		uOff = protocol.IntEnd(b)
		valType = protocol.ScalarValueInt
	case float64:
		protocol.NumStart(b)
		protocol.NumAddVal(b, v)
		uOff = protocol.NumEnd(b)
		valType = protocol.ScalarValueNum
	case string:
		sOff := b.CreateString(v)
		protocol.StrStart(b)
		protocol.StrAddVal(b, sOff)
		uOff = protocol.StrEnd(b)
		valType = protocol.ScalarValueStr
	default:
		panic("buildScalarCell: unsupported cell type")
	}
	protocol.ScalarStart(b)
	protocol.ScalarAddValType(b, valType)
	protocol.ScalarAddVal(b, uOff)
	return protocol.ScalarEnd(b)
}

// buildGrid constructs a rows x cols protocol.Grid (row-major) from cells.
func buildGrid(rows, cols int, cells ...any) *protocol.Grid {
	b := flatbuffers.NewBuilder(0)
	offs := make([]flatbuffers.UOffsetT, len(cells))
	for i, c := range cells {
		offs[i] = buildScalarCell(b, c)
	}
	protocol.GridStartDataVector(b, len(offs))
	for i := len(offs) - 1; i >= 0; i-- {
		b.PrependUOffsetT(offs[i])
	}
	vec := b.EndVector(len(offs))
	protocol.GridStart(b)
	protocol.GridAddRows(b, int32(rows))
	protocol.GridAddCols(b, int32(cols))
	protocol.GridAddData(b, vec)
	g := protocol.GridEnd(b)
	b.Finish(g)
	return protocol.GetRootAsGrid(b.FinishedBytes(), 0)
}

// TestRangeNumericCells verifies the shared decoder visits only numeric cells
// (Int and Num), skips strings/nils, and ignores a nil grid.
func TestRangeNumericCells(t *testing.T) {
	// 2x3 grid: ints, floats, a string and a nil that must be skipped.
	g := buildGrid(2, 3, 1, 2.5, "skip", nil, 3, 4.0)

	var got []float64
	rangeNumericCells(g, func(v float64) { got = append(got, v) })

	want := []float64{1, 2.5, 3, 4}
	if len(got) != len(want) {
		t.Fatalf("visited %v, want %v", got, want)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Errorf("cell %d = %v, want %v", i, got[i], want[i])
		}
	}

	// nil grid: no calls.
	calls := 0
	rangeNumericCells(nil, func(float64) { calls++ })
	if calls != 0 {
		t.Errorf("nil grid produced %d calls, want 0", calls)
	}
}

func TestSumGrid(t *testing.T) {
	s := &Service{}
	g := buildGrid(2, 3, 1, 2.5, "skip", nil, 3, 4.0)
	got, err := s.SumGrid(context.Background(), g)
	if err != nil {
		t.Fatal(err)
	}
	if want := 10.5; got != want {
		t.Errorf("SumGrid = %v, want %v", got, want)
	}

	zero, err := s.SumGrid(context.Background(), nil)
	if err != nil || zero != 0 {
		t.Errorf("SumGrid(nil) = %v, %v; want 0, nil", zero, err)
	}
}

func TestStatsGrid(t *testing.T) {
	s := &Service{}
	g := buildGrid(2, 3, 1, 2.5, "skip", nil, 3, 4.0)
	rows, err := s.StatsGrid(context.Background(), g)
	if err != nil {
		t.Fatal(err)
	}
	// Header + 5 stats rows.
	if len(rows) != 6 {
		t.Fatalf("StatsGrid returned %d rows, want 6", len(rows))
	}
	want := map[string]float64{"Count": 4, "Sum": 10.5, "Mean": 2.625, "Min": 1, "Max": 4}
	for _, r := range rows[1:] {
		label := r[0].(string)
		exp, ok := want[label]
		if !ok {
			t.Errorf("unexpected stat row %q", label)
			continue
		}
		var act float64
		switch v := r[1].(type) {
		case int:
			act = float64(v)
		case float64:
			act = v
		default:
			t.Errorf("%s: unexpected value type %T", label, r[1])
			continue
		}
		if math.Abs(act-exp) > 1e-9 {
			t.Errorf("%s = %v, want %v", label, act, exp)
		}
	}
}
