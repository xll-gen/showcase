//go:build ignore

// Generates the ribbon button icons under icons/ (pure Go stdlib, no deps).
// The PNGs have transparent corners on purpose: they exercise xll-gen's
// alpha-preserving GDI+ decode path (ribbon image files, xll-gen > v0.4.2).
//
// Run from the repo root:  go run tools/gen_icons.go
package main

import (
	"image"
	"image/color"
	"image/png"
	"os"
)

// disk reports whether (x,y) lies inside the circle of radius r at (cx,cy).
func disk(x, y, cx, cy, r float64) bool {
	dx, dy := x-cx, y-cy
	return dx*dx+dy*dy <= r*r
}

// circleIcon draws a filled circle of `bg` on a transparent canvas and lets
// the caller paint a white glyph via mark(x, y) -> bool.
func circleIcon(size int, bg color.NRGBA, mark func(x, y float64) bool) *image.NRGBA {
	img := image.NewNRGBA(image.Rect(0, 0, size, size))
	c := float64(size-1) / 2
	r := float64(size)/2 - 1
	for y := 0; y < size; y++ {
		for x := 0; x < size; x++ {
			fx, fy := float64(x), float64(y)
			if !disk(fx, fy, c, c, r) {
				continue // transparent corner
			}
			px := bg
			if mark(fx, fy) {
				px = color.NRGBA{R: 255, G: 255, B: 255, A: 255}
			}
			img.SetNRGBA(x, y, px)
		}
	}
	return img
}

func save(path string, img image.Image) {
	f, err := os.Create(path)
	if err != nil {
		panic(err)
	}
	defer f.Close()
	if err := png.Encode(f, img); err != nil {
		panic(err)
	}
}

func main() {
	if err := os.MkdirAll("icons", 0o755); err != nil {
		panic(err)
	}

	// build.png (32x32, size: large): green circle + white plus.
	{
		const s = 32
		c := float64(s-1) / 2
		bar := 2.6 // half-thickness of the plus bars
		arm := 9.0 // half-length of the plus bars
		img := circleIcon(s, color.NRGBA{R: 46, G: 139, B: 87, A: 255}, func(x, y float64) bool {
			dx, dy := x-c, y-c
			h := dy >= -bar && dy <= bar && dx >= -arm && dx <= arm
			v := dx >= -bar && dx <= bar && dy >= -arm && dy <= arm
			return h || v
		})
		save("icons/build.png", img)
	}

	// clear.png (16x16, size: normal): red circle + white X.
	{
		const s = 16
		c := float64(s-1) / 2
		img := circleIcon(s, color.NRGBA{R: 200, G: 55, B: 44, A: 255}, func(x, y float64) bool {
			dx, dy := x-c, y-c
			if dx < -4.5 || dx > 4.5 || dy < -4.5 || dy > 4.5 {
				return false
			}
			d1 := dx - dy
			d2 := dx + dy
			return (d1 >= -1.6 && d1 <= 1.6) || (d2 >= -1.6 && d2 <= 1.6)
		})
		save("icons/clear.png", img)
	}
}
