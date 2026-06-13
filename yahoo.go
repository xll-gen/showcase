//go:build windows

package main

import (
	"context"
	"encoding/json"
	"fmt"
	"io"
	"net/http"
	"net/url"
	"time"
)

// yahoo.go is a tiny, stdlib-only client for Yahoo Finance's public
// v8/finance/chart endpoint. It backs the YDP (one quote field) and YDH
// (historical OHLCV table) showcase functions. No new module deps: net/http +
// encoding/json + time only.

const (
	// yahooChartBase is the v8 chart endpoint. {ticker} is appended.
	yahooChartBase = "https://query1.finance.yahoo.com/v8/finance/chart/"
	// yahooUA is a browser-ish User-Agent: Yahoo rejects some default Go UAs
	// (returns 429/empty), so we always set one.
	yahooUA = "Mozilla/5.0 (xll-gen-showcase)"
	// yahooTimeout bounds a single fetch; the handler's ctx may impose less.
	yahooTimeout = 8 * time.Second
)

// chartEnvelope mirrors the relevant slice of the v8/finance/chart JSON:
//
//	{ "chart": { "result": [ { "meta": {...},
//	                            "timestamp": [...],
//	                            "indicators": { "quote": [ {open,high,low,close,volume} ] } } ],
//	             "error": null } }
type chartEnvelope struct {
	Chart struct {
		Result []chartResult `json:"result"`
		Error  *chartError   `json:"error"`
	} `json:"chart"`
}

type chartError struct {
	Code        string `json:"code"`
	Description string `json:"description"`
}

func (e *chartError) Error() string {
	if e == nil {
		return "yahoo: unknown error"
	}
	return fmt.Sprintf("yahoo: %s: %s", e.Code, e.Description)
}

// chartResult is result[0]: meta (quote snapshot) + the historical series.
type chartResult struct {
	Meta struct {
		Currency           string  `json:"currency"`
		Symbol             string  `json:"symbol"`
		ExchangeName       string  `json:"exchangeName"`
		RegularMarketPrice float64 `json:"regularMarketPrice"`
		ChartPreviousClose float64 `json:"chartPreviousClose"`
		PreviousClose      float64 `json:"previousClose"`
		RegularMarketTime  int64   `json:"regularMarketTime"`
		RegularMarketDayHigh   float64 `json:"regularMarketDayHigh"`
		RegularMarketDayLow    float64 `json:"regularMarketDayLow"`
		RegularMarketVolume    float64 `json:"regularMarketVolume"`
		FiftyTwoWeekHigh       float64 `json:"fiftyTwoWeekHigh"`
		FiftyTwoWeekLow        float64 `json:"fiftyTwoWeekLow"`
	} `json:"meta"`
	Timestamp  []int64 `json:"timestamp"`
	Indicators struct {
		Quote []struct {
			Open   []*float64 `json:"open"`
			High   []*float64 `json:"high"`
			Low    []*float64 `json:"low"`
			Close  []*float64 `json:"close"`
			Volume []*float64 `json:"volume"`
		} `json:"quote"`
	} `json:"indicators"`
}

// fetchChart calls v8/finance/chart/{ticker} with the given query params and
// returns the parsed result[0]. It honors ctx (with an ~8s ceiling), sets a
// User-Agent, and turns Yahoo's error envelope / empty result / non-200 into a
// Go error.
func fetchChart(ctx context.Context, ticker string, params url.Values) (*chartResult, error) {
	if ticker == "" {
		return nil, fmt.Errorf("empty ticker")
	}

	// Bound the request to ~8s even if the caller's ctx has no deadline.
	reqCtx, cancel := context.WithTimeout(ctx, yahooTimeout)
	defer cancel()

	u := yahooChartBase + url.PathEscape(ticker)
	if enc := params.Encode(); enc != "" {
		u += "?" + enc
	}

	req, err := http.NewRequestWithContext(reqCtx, http.MethodGet, u, nil)
	if err != nil {
		return nil, fmt.Errorf("build request: %w", err)
	}
	req.Header.Set("User-Agent", yahooUA)
	req.Header.Set("Accept", "application/json")

	resp, err := http.DefaultClient.Do(req)
	if err != nil {
		return nil, fmt.Errorf("fetch %q: %w", ticker, err)
	}
	defer resp.Body.Close()

	body, err := io.ReadAll(io.LimitReader(resp.Body, 8<<20)) // 8 MiB cap
	if err != nil {
		return nil, fmt.Errorf("read body for %q: %w", ticker, err)
	}
	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("yahoo returned HTTP %d for %q", resp.StatusCode, ticker)
	}

	var env chartEnvelope
	if err := json.Unmarshal(body, &env); err != nil {
		return nil, fmt.Errorf("parse JSON for %q: %w", ticker, err)
	}
	if env.Chart.Error != nil {
		return nil, env.Chart.Error
	}
	if len(env.Chart.Result) == 0 {
		return nil, fmt.Errorf("no data for %q", ticker)
	}
	return &env.Chart.Result[0], nil
}
