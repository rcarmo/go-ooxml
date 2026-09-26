package packaging

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

func TestZIP64IntakeBatch(t *testing.T) {
	for _, deflate := range []bool{false, true} {
		for _, descriptor := range []string{"", "signed", "unsigned"} {
			t.Run(fmt.Sprintf("deflate=%v/descriptor=%s", deflate, descriptor), func(t *testing.T) {
				data := testutil.TinyZIP64(testutil.ZIP64Fixture{Deflate: deflate, Descriptor: descriptor})
				r, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
				if err != nil {
					t.Fatal("fixture cannot open", err)
				}
				for _, f := range r.File {
					reader, err := f.Open()
					if err != nil {
						t.Fatal(err)
					}
					_, err = io.ReadAll(reader)
					_ = reader.Close()
					if err != nil {
						t.Fatal("fixture cannot read", err)
					}
				}
				p, err := OpenPreserved(data, Limits{MaxEntries: 10, MaxPartBytes: 1 << 20})
				if err != nil {
					t.Fatal(err)
				}
				var b bytes.Buffer
				if err = p.WriteTo(&b); err != nil || !bytes.Equal(b.Bytes(), data) {
					t.Fatal("no-op changed", err)
				}
				if _, err = p.SaveAs(filepath.Join(t.TempDir(), "out.zip")); err != nil {
					t.Fatal(err)
				}
			})
		}
	}
	for _, defect := range []string{"gap before ZIP64 end", "gap before locator", "ZIP64 length overflow", "ZIP64 short length", "ZIP64 length crosses locator", "short local extra", "missing local extra", "local size mismatch", "overflow local size", "duplicate local extra", "trailing local extra byte", "short central extra", "duplicate central extra", "trailing central extra byte", "classic count mismatch"} {
		t.Run(defect, func(t *testing.T) {
			data := testutil.TinyZIP64(testutil.ZIP64Fixture{Defect: defect})
			if p, err := OpenPreserved(data, Limits{MaxPartBytes: 1 << 20, MaxTotalBytes: 1 << 20}); err == nil {
				_ = p
				t.Fatal("invalid ZIP64 accepted")
			}
		})
	}
}

// Terminal gaps must fail independently of local ZIP64 support; archive/zip's
// directory reader accepts this regular tiny archive with added ZIP64 endings.
func TestZIP64TerminalGeometryBatch(t *testing.T) {
	for _, defect := range []string{"valid", "gap before ZIP64 end", "gap before locator", "bad record length"} {
		t.Run(defect, func(t *testing.T) {
			p := New()
			_, _ = p.AddPart("data.bin", "application/octet-stream", []byte("x"))
			var b bytes.Buffer
			if err := p.WriteTo(&b); err != nil {
				t.Fatal(err)
			}
			data := b.Bytes()
			eoff := len(data) - 22
			end := bytes.Clone(data[eoff:])
			count := little.Uint16(end[10:])
			size, offset := little.Uint32(end[12:]), little.Uint32(end[16:])
			out := bytes.Clone(data[:eoff])
			if defect == "gap before ZIP64 end" {
				out = append(out, 'x')
			}
			at := len(out)
			z := make([]byte, 56)
			little.PutUint32(z, 0x06064b50)
			little.PutUint64(z[4:], 44)
			if defect == "bad record length" {
				little.PutUint64(z[4:], 45)
			}
			little.PutUint64(z[24:], uint64(count))
			little.PutUint64(z[32:], uint64(count))
			little.PutUint64(z[40:], uint64(size))
			little.PutUint64(z[48:], uint64(offset))
			out = append(out, z...)
			if defect == "gap before locator" {
				out = append(out, 'x')
			}
			loc := make([]byte, 20)
			little.PutUint32(loc, 0x07064b50)
			little.PutUint64(loc[8:], uint64(at))
			little.PutUint32(loc[16:], 1)
			out = append(out, loc...)
			little.PutUint16(end[8:], 0xffff)
			little.PutUint16(end[10:], 0xffff)
			little.PutUint32(end[12:], 0xffffffff)
			little.PutUint32(end[16:], 0xffffffff)
			out = append(out, end...)
			q, err := OpenBytes(out)
			if q != nil {
				_ = q.Close()
			}
			if (err == nil) != (defect == "valid") {
				t.Fatalf("%s: %v", defect, err)
			}
		})
	}
}
