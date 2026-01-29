package utils

import "testing"

func TestEMUConversions(t *testing.T) {
	// Test inch conversions
	if got := InchesToEMU(1.0); got != 914400 {
		t.Errorf("InchesToEMU(1.0) = %d, want 914400", got)
	}
	if got := EMUToInches(914400); got != 1.0 {
		t.Errorf("EMUToInches(914400) = %f, want 1.0", got)
	}

	// Test point conversions
	if got := PointsToEMU(1.0); got != 12700 {
		t.Errorf("PointsToEMU(1.0) = %d, want 12700", got)
	}
	if got := EMUToPoints(12700); got != 1.0 {
		t.Errorf("EMUToPoints(12700) = %f, want 1.0", got)
	}

	// Test centimeter conversions
	if got := CentimetersToEMU(1.0); got != 360000 {
		t.Errorf("CentimetersToEMU(1.0) = %d, want 360000", got)
	}
	if got := EMUToCentimeters(360000); got != 1.0 {
		t.Errorf("EMUToCentimeters(360000) = %f, want 1.0", got)
	}

	// Test pixel conversions (96 DPI)
	if got := PixelsToEMU(1); got != 9525 {
		t.Errorf("PixelsToEMU(1) = %d, want 9525", got)
	}
	if got := EMUToPixels(9525); got != 1 {
		t.Errorf("EMUToPixels(9525) = %d, want 1", got)
	}

	// Test twip conversions
	if got := TwipsToEMU(1); got != 635 {
		t.Errorf("TwipsToEMU(1) = %d, want 635", got)
	}
	if got := EMUToTwips(635); got != 1 {
		t.Errorf("EMUToTwips(635) = %d, want 1", got)
	}

	// Test half-point conversions
	if got := HalfPointsToPoints(24); got != 12.0 {
		t.Errorf("HalfPointsToPoints(24) = %f, want 12.0", got)
	}
	if got := PointsToHalfPoints(12.0); got != 24 {
		t.Errorf("PointsToHalfPoints(12.0) = %d, want 24", got)
	}
}
