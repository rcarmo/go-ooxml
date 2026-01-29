package utils

import "testing"

func TestParseHexColor(t *testing.T) {
	tests := []struct {
		hex     string
		want    Color
		wantErr bool
	}{
		{"FF0000", Color{255, 255, 0, 0}, false},
		{"#FF0000", Color{255, 255, 0, 0}, false},
		{"00FF00", Color{255, 0, 255, 0}, false},
		{"0000FF", Color{255, 0, 0, 255}, false},
		{"FFFFFF", Color{255, 255, 255, 255}, false},
		{"000000", Color{255, 0, 0, 0}, false},
		{"80FF0000", Color{128, 255, 0, 0}, false}, // ARGB with alpha
		{"ff0000", Color{255, 255, 0, 0}, false},   // lowercase
		{"invalid", Color{}, true},
		{"FFF", Color{}, true},     // too short
		{"FFFFFFFFF", Color{}, true}, // too long
	}

	for _, tt := range tests {
		got, err := ParseHexColor(tt.hex)
		if (err != nil) != tt.wantErr {
			t.Errorf("ParseHexColor(%q) error = %v, wantErr %v", tt.hex, err, tt.wantErr)
			continue
		}
		if !tt.wantErr && got != tt.want {
			t.Errorf("ParseHexColor(%q) = %+v, want %+v", tt.hex, got, tt.want)
		}
	}
}

func TestColorToHex(t *testing.T) {
	tests := []struct {
		color  Color
		expect string
	}{
		{Color{255, 255, 0, 0}, "FF0000"},
		{Color{255, 0, 255, 0}, "00FF00"},
		{Color{255, 0, 0, 255}, "0000FF"},
		{Color{128, 255, 0, 0}, "80FF0000"}, // with alpha
	}

	for _, tt := range tests {
		got := tt.color.ToHex()
		if got != tt.expect {
			t.Errorf("%+v.ToHex() = %q, want %q", tt.color, got, tt.expect)
		}
	}
}

func TestColorToARGB(t *testing.T) {
	tests := []struct {
		color  Color
		expect string
	}{
		{Color{255, 255, 0, 0}, "FFFF0000"},
		{Color{128, 0, 255, 0}, "8000FF00"},
	}

	for _, tt := range tests {
		got := tt.color.ToARGB()
		if got != tt.expect {
			t.Errorf("%+v.ToARGB() = %q, want %q", tt.color, got, tt.expect)
		}
	}
}
