package presentation

import (
	"errors"
	"math"
	"sort"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func TestGraphicsGroupMappingIndependentDistribution(t *testing.T) {
	forward, inverse := []float64{}, []float64{}
	count := 0
	for _, rotation := range []int64{0, 5400000, 10800000, 16200000, 1234567, 9999999} {
		for _, flipH := range []bool{false, true} {
			for _, flipV := range []bool{false, true} {
				for _, scale := range []float64{.5, 1, 2} {
					for _, input := range []GroupPoint{{-10, 20}, {200, 700}, {123.25, -456.5}} {
						frame := GroupTransform{PictureTransform: PictureTransform{X: 1000, Y: -2000, Width: int64(6000 * scale), Height: int64(8000 * scale), Rotation: rotation, FlipH: flipH, FlipV: flipV}, ChildX: 10, ChildY: 20, ChildWidth: 600, ChildHeight: 800}
						cx, cy := float64(frame.X)+float64(frame.Width)/2, float64(frame.Y)+float64(frame.Height)/2
						qx, qy := float64(frame.X)+(input.X-float64(frame.ChildX))*(float64(frame.Width)/float64(frame.ChildWidth)), float64(frame.Y)+(input.Y-float64(frame.ChildY))*(float64(frame.Height)/float64(frame.ChildHeight))
						dx, dy := qx-cx, qy-cy
						if flipH {
							dx = -dx
						}
						if flipV {
							dy = -dy
						}
						radius := math.Hypot(dx, dy)
						angle := math.Atan2(dy, dx) + float64(rotation)/60000*math.Pi/180
						reference := GroupPoint{cx + radius*math.Cos(angle), cy + radius*math.Sin(angle)}
						actual, e := MapGroupPoint(frame, input)
						if e != nil {
							t.Fatal(e)
						}
						readback, e := UnmapGroupPoint(frame, actual)
						if e != nil {
							t.Fatal(e)
						}
						forward = append(forward, math.Abs(actual.X-reference.X), math.Abs(actual.Y-reference.Y))
						inverse = append(inverse, math.Abs(readback.X-input.X), math.Abs(readback.Y-input.Y))
						count++
					}
				}
			}
		}
	}
	if count != 216 {
		t.Fatal("matrix count")
	}
	sort.Float64s(forward)
	sort.Float64s(inverse)
	if forward[len(forward)-1] >= 1e-5 || inverse[len(inverse)-1] >= 1e-5 {
		t.Fatalf("mapping exceeds shared tolerance f=%g i=%g", forward[len(forward)-1], inverse[len(inverse)-1])
	}
	t.Logf("mapping error EMUs: count=%d forward median=%g p95=%g max=%g inverse median=%g p95=%g max=%g", count, forward[len(forward)/2], forward[int(float64(len(forward)-1)*.95)], forward[len(forward)-1], inverse[len(inverse)/2], inverse[int(float64(len(inverse)-1)*.95)], inverse[len(inverse)-1])
}
func TestGraphicsGroupMappingRefusalControls(t *testing.T) {
	frame := GroupTransform{PictureTransform: PictureTransform{Width: 3000, Height: 4000}, ChildWidth: 300, ChildHeight: 400}
	for _, action := range []func() (GroupPoint, error){func() (GroupPoint, error) {
		bad := frame
		bad.ChildWidth = 0
		return MapGroupPoint(bad, GroupPoint{1, 1})
	}, func() (GroupPoint, error) {
		bad := frame
		bad.Width = math.MaxInt32
		bad.ChildWidth = 1
		return MapGroupPoint(bad, GroupPoint{math.MaxInt32, 1})
	}, func() (GroupPoint, error) { return UnmapGroupPoint(frame, GroupPoint{math.Inf(1), 0}) }, func() (GroupPoint, error) { return MapGroupPointChain(nil, GroupPoint{}) }, func() (GroupPoint, error) { return UnmapGroupPointChain(make([]GroupTransform, 17), GroupPoint{}) }} {
		_, e := action()
		var r *packaging.Refusal
		if !errors.As(e, &r) || r.Kind != "PPTX_GROUP_UNSUPPORTED" {
			t.Fatalf("mapping refusal %v", e)
		}
	}
}
