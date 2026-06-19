//go:build windows

// Command xll_showcase is the Go server backing the xll-gen-showcase add-in.
// It implements every worksheet function, every RTD stream, and every ribbon
// command declared in xll.yaml. Worksheet functions are pure Go; command
// handlers drive Excel through the sugar COM-automation library.
//
// The handlers are split by concern across the package (all package main):
//   - functions.go   — scalar worksheet functions (Add, Greet, SumGrid, ...)
//   - spill.go       — dynamic-array (spill) functions (TimesTable, StatsGrid, ...)
//   - rtd.go         — RTD streaming functions + RTD connect/disconnect lifecycle
//   - marketdata.go  — Yahoo Finance rtd-once handlers (YDP, YDH); HTTP client in yahoo.go
//   - commands.go    — sugar-driven ribbon/command handlers + the Excel-attach scaffolds
//   - events.go      — calculation event handlers (OnRecalc, OnCalculationCanceled)
package main

import (
	"sync"

	"xll_showcase/generated"
)

// Service implements the generated XllService interface.
type Service struct {
	// recalcMu guards the recalc-target location captured by BuildShowcaseSheet
	// and read by OnRecalc. BuildShowcaseSheet runs on a command worker while
	// OnRecalc runs on a server worker goroutine, so they race without it.
	recalcMu    sync.Mutex
	recalcSheet string
	recalcRow   int
}

func main() {
	generated.Serve(&Service{})
}
