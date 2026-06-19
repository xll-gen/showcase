//go:build windows

package main

import (
	"context"
	"fmt"
	"net/url"
	"strconv"
	"strings"
	"time"
)

// ---------------------------------------------------------------------------
// Live Market Data — Yahoo Finance (mode:"rtd-once")
//
// YDP and YDH are written as ORDINARY sync-shaped handlers — take a ctx, fetch,
// return a value — yet they run mode:"rtd-once": the cell shows #GETTING_DATA,
// the handler runs once OFF the calc thread (so Excel never blocks on the HTTP
// round-trip), and the cell settles when the fetch returns. YDH is the headline
// rtd-once-GRID demo: it returns a [][]any that SPILLS an OHLCV table, exercising
// the new grid guest->host + RTD-triggered re-spill machinery end to end.
//
// The Yahoo HTTP client (fetchChart and friends) lives in yahoo.go.
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
