//go:build windows

package main

import (
	"context"
	"fmt"
	"math/rand"
	"time"

	"xll_showcase/generated"
)

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
// RTD lifecycle
// ---------------------------------------------------------------------------

func (s *Service) OnRtdConnect(ctx context.Context, topicID int32, strings []string, newValues bool) error {
	return nil
}

func (s *Service) OnRtdDisconnect(ctx context.Context, topicID int32) error { return nil }
