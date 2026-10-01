package packaging

import (
	"bytes"
	"hash/crc32"
	"testing"
)

func TestZIPCRC32VectorsAndCustody(t *testing.T) {
	for _, tc := range []struct {
		source []byte
		want   uint32
	}{{[]byte("123456789"), 0xcbf43926}, {[]byte{}, 0}, {[]byte{0, 1, 2, 255}, crc32.ChecksumIEEE([]byte{0, 1, 2, 255})}} {
		original := bytes.Clone(tc.source)
		if got := ZIPCRC32(tc.source); got != tc.want {
			t.Errorf("ZIPCRC32(%q)=%08X want %08X", tc.source, got, tc.want)
		}
		if !bytes.Equal(tc.source, original) {
			t.Fatal("caller payload changed")
		}
	}
}
