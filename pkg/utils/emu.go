package utils

// EMU (English Metric Units) constants used in OOXML.
const (
	EMUsPerInch       = 914400
	EMUsPerPoint      = 12700
	EMUsPerCentimeter = 360000
	EMUsPerPixel      = 9525 // at 96 DPI
	EMUsPerTwip       = 635  // 1 twip = 1/20 of a point
)

// InchesToEMU converts inches to EMUs.
func InchesToEMU(inches float64) int64 {
	return int64(inches * EMUsPerInch)
}

// EMUToInches converts EMUs to inches.
func EMUToInches(emu int64) float64 {
	return float64(emu) / EMUsPerInch
}

// PointsToEMU converts points to EMUs.
func PointsToEMU(points float64) int64 {
	return int64(points * EMUsPerPoint)
}

// EMUToPoints converts EMUs to points.
func EMUToPoints(emu int64) float64 {
	return float64(emu) / EMUsPerPoint
}

// CentimetersToEMU converts centimeters to EMUs.
func CentimetersToEMU(cm float64) int64 {
	return int64(cm * EMUsPerCentimeter)
}

// EMUToCentimeters converts EMUs to centimeters.
func EMUToCentimeters(emu int64) float64 {
	return float64(emu) / EMUsPerCentimeter
}

// PixelsToEMU converts pixels to EMUs (at 96 DPI).
func PixelsToEMU(pixels int) int64 {
	return int64(pixels) * EMUsPerPixel
}

// EMUToPixels converts EMUs to pixels (at 96 DPI).
func EMUToPixels(emu int64) int {
	return int(emu / EMUsPerPixel)
}

// TwipsToEMU converts twips to EMUs.
func TwipsToEMU(twips int64) int64 {
	return twips * EMUsPerTwip
}

// EMUToTwips converts EMUs to twips.
func EMUToTwips(emu int64) int64 {
	return emu / EMUsPerTwip
}

// HalfPointsToPoints converts half-points (used in OOXML) to points.
func HalfPointsToPoints(halfPoints int64) float64 {
	return float64(halfPoints) / 2.0
}

// PointsToHalfPoints converts points to half-points.
func PointsToHalfPoints(points float64) int64 {
	return int64(points * 2.0)
}
