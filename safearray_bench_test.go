//go:build windows

// Benchmarks for the VT_ARRAY|VT_VARIANT SAFEARRAY encode/decode paths that
// back `Range.Value` / `Range.SetValue`. SafeArray* APIs live in oleaut32 and
// need no CoInitialize, so these run without Excel.
//
// Grid sizes bracket the range where the cost becomes visible: 100x100 (10k
// cells) is the smallest size where the per-cell overhead is measurable next
// to the cross-process COM marshal, and 500x500 (250k cells) is a realistic
// "read the whole used range" workload.

package sugar

import (
	"fmt"
	"testing"
)

func benchNumGrid(rows, cols int) [][]interface{} {
	g := make([][]interface{}, rows)
	for r := range g {
		row := make([]interface{}, cols)
		for c := range row {
			row[c] = float64(r*cols + c)
		}
		g[r] = row
	}
	return g
}

func benchStrGrid(rows, cols int) [][]interface{} {
	g := make([][]interface{}, rows)
	for r := range g {
		row := make([]interface{}, cols)
		for c := range row {
			row[c] = fmt.Sprintf("cell-%d-%d", r, c)
		}
		g[r] = row
	}
	return g
}

var benchGrids = []struct {
	name string
	gen  func(int, int) [][]interface{}
}{
	{"num", benchNumGrid},
	{"str", benchStrGrid},
}

var benchSizes = []struct {
	name       string
	rows, cols int
}{
	{"100x100", 100, 100},
	{"500x500", 500, 500},
}

func BenchmarkEncode2D(b *testing.B) {
	for _, g := range benchGrids {
		for _, s := range benchSizes {
			src := g.gen(s.rows, s.cols)
			b.Run(g.name+"/"+s.name, func(b *testing.B) {
				b.ReportAllocs()
				for i := 0; i < b.N; i++ {
					v, err := encodeVariantArray(src)
					if err != nil {
						b.Fatal(err)
					}
					v.Clear()
				}
			})
		}
	}
}

func BenchmarkDecode2D(b *testing.B) {
	for _, g := range benchGrids {
		for _, s := range benchSizes {
			src := g.gen(s.rows, s.cols)
			v, err := encodeVariantArray(src)
			if err != nil {
				b.Fatal(err)
			}
			b.Run(g.name+"/"+s.name, func(b *testing.B) {
				b.ReportAllocs()
				for i := 0; i < b.N; i++ {
					if _, err := decodeVariantArray(v); err != nil {
						b.Fatal(err)
					}
				}
			})
			v.Clear()
		}
	}
}
