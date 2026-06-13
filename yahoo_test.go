//go:build windows

package main

import (
	"encoding/json"
	"testing"
	"time"
)

// A captured-shape v8/finance/chart response (trimmed to the fields YDP/YDH
// read). The third history bar is a holiday: Yahoo emits a timestamp but nulls
// every OHLCV field, which YDH must skip. No live network is used in tests.
const fixtureChartJSON = `{
  "chart": {
    "result": [
      {
        "meta": {
          "currency": "USD",
          "symbol": "AAPL",
          "exchangeName": "NMS",
          "regularMarketPrice": 195.50,
          "chartPreviousClose": 190.00,
          "previousClose": 190.00,
          "regularMarketTime": 1718294400,
          "regularMarketDayHigh": 196.10,
          "regularMarketDayLow": 193.20,
          "regularMarketVolume": 51234567,
          "fiftyTwoWeekHigh": 199.62,
          "fiftyTwoWeekLow": 164.08
        },
        "timestamp": [1718000000, 1718086400, 1718172800],
        "indicators": {
          "quote": [
            {
              "open":   [100.0, 101.0, null],
              "high":   [102.0, 103.0, null],
              "low":    [ 99.0, 100.5, null],
              "close":  [101.5, 102.5, null],
              "volume": [1000000, 1100000, null]
            }
          ]
        }
      }
    ],
    "error": null
  }
}`

func parseFixture(t *testing.T) *chartResult {
	t.Helper()
	var env chartEnvelope
	if err := json.Unmarshal([]byte(fixtureChartJSON), &env); err != nil {
		t.Fatalf("unmarshal fixture: %v", err)
	}
	if len(env.Chart.Result) == 0 {
		t.Fatal("fixture has no result")
	}
	return &env.Chart.Result[0]
}

func TestYDPFieldMapping(t *testing.T) {
	res := parseFixture(t)
	m := &res.Meta

	// price
	if got := m.RegularMarketPrice; got != 195.50 {
		t.Errorf("price = %v, want 195.50", got)
	}
	// change = price - prevClose
	if got := m.RegularMarketPrice - m.ChartPreviousClose; got != 5.50 {
		t.Errorf("change = %v, want 5.50", got)
	}
	// change% = (price-prev)/prev*100
	wantPct := (195.50 - 190.00) / 190.00 * 100
	if got := (m.RegularMarketPrice - m.ChartPreviousClose) / m.ChartPreviousClose * 100; got != wantPct {
		t.Errorf("change%% = %v, want %v", got, wantPct)
	}
	// text fields
	if m.Currency != "USD" {
		t.Errorf("currency = %q, want USD", m.Currency)
	}
	if m.ExchangeName != "NMS" {
		t.Errorf("exchange = %q, want NMS", m.ExchangeName)
	}
	// 52w
	if m.FiftyTwoWeekHigh != 199.62 || m.FiftyTwoWeekLow != 164.08 {
		t.Errorf("52w = %v/%v, want 199.62/164.08", m.FiftyTwoWeekHigh, m.FiftyTwoWeekLow)
	}
	if m.RegularMarketDayHigh != 196.10 || m.RegularMarketDayLow != 193.20 {
		t.Errorf("day high/low = %v/%v", m.RegularMarketDayHigh, m.RegularMarketDayLow)
	}
	if m.RegularMarketVolume != 51234567 {
		t.Errorf("volume = %v", m.RegularMarketVolume)
	}
}

// TestYDHGridBuild verifies YDH's grid construction from the parsed series:
// header row + one row per NON-null bar (the holiday bar is skipped).
func TestYDHGridBuild(t *testing.T) {
	res := parseFixture(t)
	q := res.Indicators.Quote[0]

	grid := [][]any{{"Date", "Open", "High", "Low", "Close", "Volume"}}
	for i, ts := range res.Timestamp {
		if i >= len(q.Open) || q.Open[i] == nil || q.High[i] == nil ||
			q.Low[i] == nil || q.Close[i] == nil || q.Volume[i] == nil {
			continue
		}
		grid = append(grid, []any{
			time.Unix(ts, 0).Format("2006-01-02"), *q.Open[i], *q.High[i], *q.Low[i], *q.Close[i], *q.Volume[i],
		})
	}

	// 1 header + 2 valid bars (3rd is the null holiday).
	if len(grid) != 3 {
		t.Fatalf("grid rows = %d, want 3 (header + 2 bars)", len(grid))
	}
	if len(grid[0]) != 6 {
		t.Fatalf("header cols = %d, want 6", len(grid[0]))
	}
	if grid[0][0] != "Date" || grid[0][5] != "Volume" {
		t.Errorf("header = %v", grid[0])
	}
	// First bar values.
	if grid[1][1] != 100.0 || grid[1][4] != 101.5 || grid[1][5] != 1000000.0 {
		t.Errorf("bar 1 = %v", grid[1])
	}
	// Date is a YYYY-MM-DD string.
	if d, ok := grid[1][0].(string); !ok || len(d) != 10 {
		t.Errorf("date cell = %v (want YYYY-MM-DD string)", grid[1][0])
	}
}

// TestChartErrorEnvelope confirms a Yahoo error envelope is surfaced.
func TestChartErrorEnvelope(t *testing.T) {
	const errJSON = `{"chart":{"result":null,"error":{"code":"Not Found","description":"No data found, symbol may be delisted"}}}`
	var env chartEnvelope
	if err := json.Unmarshal([]byte(errJSON), &env); err != nil {
		t.Fatalf("unmarshal: %v", err)
	}
	if env.Chart.Error == nil {
		t.Fatal("expected non-nil chart.error")
	}
	if env.Chart.Error.Error() == "" {
		t.Error("error string is empty")
	}
}
