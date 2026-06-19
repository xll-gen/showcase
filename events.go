//go:build windows

package main

import (
	"context"
	"time"

	flatbuffers "github.com/google/flatbuffers/go"
	"github.com/xll-gen/types/go/protocol"

	"xll_showcase/generated"
)

// ---------------------------------------------------------------------------
// Events
// ---------------------------------------------------------------------------

// OnRecalc fires after every Excel calculation cycle (CalculationEnded). The
// XLL invokes this handler via a SYNCHRONOUS shm round-trip that BLOCKS Excel's
// STA (main) thread until this returns. It must therefore do NO Excel COM
// automation: any COM call (attach/Find/SetValue) needs the STA thread — which
// is blocked waiting on this very handler — so it would deadlock until the
// round-trip times out (~2s freeze on every recalc).
//
// Instead we enqueue a DEFERRED command via generated.ScheduleSet: the write is
// returned in the round-trip response and executed by the XLL on the STA thread
// AFTER this handler returns — no re-entrant COM, no deadlock. The target cell
// (column B at the marker row on the showcase sheet) was captured by
// BuildShowcaseSheet, so we never have to touch Excel to locate it here.
func (s *Service) OnRecalc(ctx context.Context) error {
	s.recalcMu.Lock()
	sheetName := s.recalcSheet
	recalcRow := s.recalcRow
	s.recalcMu.Unlock()

	// BuildShowcaseSheet hasn't run yet (no marker cell): nothing to stamp.
	if recalcRow == 0 {
		return nil
	}

	// Target cell B{recalcRow}: 0-based row recalcRow-1, 0-based col 1.
	b := flatbuffers.NewBuilder(0)
	sOff := b.CreateString(sheetName)
	protocol.RangeStartRefsVector(b, 1)
	protocol.CreateRect(b, int32(recalcRow-1), int32(recalcRow-1), 1, 1)
	refsOff := b.EndVector(1)
	protocol.RangeStart(b)
	protocol.RangeAddSheetName(b, sOff)
	protocol.RangeAddRefs(b, refsOff)
	rOff := protocol.RangeEnd(b)
	b.Finish(rOff)
	r := protocol.GetRootAsRange(b.FinishedBytes(), 0)

	// Value: a STRING "recalc @ HH:MM:SS".
	b2 := flatbuffers.NewBuilder(0)
	strOff := b2.CreateString("recalc @ " + time.Now().Format("15:04:05"))
	protocol.StrStart(b2)
	protocol.StrAddVal(b2, strOff)
	sValOff := protocol.StrEnd(b2)
	protocol.AnyStart(b2)
	protocol.AnyAddValType(b2, protocol.AnyValueStr)
	protocol.AnyAddVal(b2, sValOff)
	aOff := protocol.AnyEnd(b2)
	b2.Finish(aOff)
	v := protocol.GetRootAsAny(b2.FinishedBytes(), 0)

	generated.ScheduleSet(r, v)
	return nil
}

func (s *Service) OnCalculationCanceled(ctx context.Context) error { return nil }
