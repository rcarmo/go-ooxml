package packaging

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

func TestZIP64DeliveryBatch(t *testing.T) {
	for _, descriptor := range []string{"", "signed", "unsigned"} {
		t.Run("edited forced ZIP64 "+descriptor, func(t *testing.T) {
			source := testutil.TinyZIP64(testutil.ZIP64Fixture{Descriptor: descriptor, Deflate: true})
			p, err := OpenPreserved(source, Limits{MaxEntries: 10, MaxTotalBytes: 1 << 20})
			if err != nil {
				t.Fatal(err)
			}
			_, hash, _ := p.Part("data.bin")
			if err = p.Replace([]Replacement{{Part: "data.bin", ExpectedSHA256: hash, Data: []byte("updated")}}); err != nil {
				t.Fatal(err)
			}
			dest := filepath.Join(t.TempDir(), "out.zip")
			if _, err = p.SaveAs(dest); err != nil {
				t.Fatal(err)
			}
			out, err := os.ReadFile(dest)
			if err != nil {
				t.Fatal(err)
			}
			q, err := OpenPreserved(out, Limits{})
			if err != nil {
				t.Fatal(err)
			}
			for _, name := range []string{"[Content_Types].xml", "empty.bin"} {
				before, _, _ := p.Part(name)
				after, _, err := q.Part(name)
				if err != nil || !bytes.Equal(before, after) {
					t.Fatal("untouched payload changed", name, err)
				}
			}
			data, _, _ := q.Part("data.bin")
			if string(data) != "updated" {
				t.Fatal("edit not delivered")
			}
			plan, err := q.PlanGraphMutation(GraphMutation{Additions: []PartAddition{{Name: "new.bin", ContentType: "application/octet-stream", Data: []byte("new")}}})
			if err != nil {
				t.Fatal(err)
			}
			if err = q.ApplyGraphPlan(plan); err != nil {
				t.Fatal(err)
			}
			if _, err = q.SaveAs(filepath.Join(t.TempDir(), "added.zip")); err != nil {
				t.Fatal(err)
			}
		})
	}
	t.Run("writer count sentinel at65535 entries", func(t *testing.T) {
		var b bytes.Buffer
		z := zip.NewWriter(&b)
		w, err := z.CreateHeader(&zip.FileHeader{Name: ContentTypesPath, Method: zip.Store})
		if err != nil {
			t.Fatal(err)
		}
		_, _ = w.Write([]byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/octet-stream"/></Types>`))
		for i := 1; i < math.MaxUint16; i++ {
			if _, err = z.CreateHeader(&zip.FileHeader{Name: fmt.Sprintf("p%05d.bin", i), Method: zip.Store}); err != nil {
				t.Fatal(err)
			}
		}
		if err = z.Close(); err != nil {
			t.Fatal(err)
		}
		data := b.Bytes()
		end := len(data) - 22
		if binary.LittleEndian.Uint16(data[end+10:]) != math.MaxUint16 || binary.LittleEndian.Uint32(data[end-20:]) != 0x07064b50 {
			t.Fatal("writer did not emit ZIP64 at sentinel")
		}
		q, err := OpenBytes(data)
		if err != nil {
			t.Fatal(err)
		}
		_ = q.Close()
		_, err = OpenPreserved(data, Limits{MaxEntries: math.MaxUint16 - 1})
		var refusal *Refusal
		if !errors.As(err, &refusal) || refusal.Kind != "resource_limit" {
			t.Fatal("entry budget not enforced", err)
		}
	})
	t.Run("required fields cannot borrow following TLV bytes", func(t *testing.T) {
		extra := make([]byte, 24)
		little.PutUint16(extra, 1)
		little.PutUint16(extra[2:], 8)
		little.PutUint64(extra[4:], 1)
		little.PutUint16(extra[12:], 0xcafe)
		little.PutUint16(extra[14:], 8)
		little.PutUint64(extra[16:], 2)
		if _, err := zip64Values(extra, math.MaxUint32, math.MaxUint32); err == nil {
			t.Fatal("borrowed unrelated extra bytes")
		}
		extra = extra[:12]
		v, err := zip64Values(extra, math.MaxUint32, 7)
		if err != nil || v[0] != 1 || v[1] != 7 {
			t.Fatal(v, err)
		}
	})
}
