package packaging

import (
	"path/filepath"
	"testing"
)

func FuzzPackageAddPartRoundTrip(f *testing.F) {
	f.Add("custom.xml", "application/xml", "<root/>")
	f.Add("data.bin", "application/octet-stream", "payload")

	f.Fuzz(func(t *testing.T, name, contentType, payload string) {
		if len(name) == 0 || len(name) > 64 || len(contentType) == 0 || len(contentType) > 128 {
			return
		}
		if len(payload) > 4096 {
			return
		}

		pkg := New()
		defer pkg.Close()

		if _, err := pkg.AddPart(name, contentType, []byte(payload)); err != nil {
			return
		}

		path := filepath.Join(t.TempDir(), "fuzz.zip")
		if err := pkg.SaveAs(path); err != nil {
			t.Fatalf("SaveAs() error = %v", err)
		}

		reopen, err := Open(path)
		if err != nil {
			t.Fatalf("Open() error = %v", err)
		}
		defer reopen.Close()

		part, err := reopen.GetPart(name)
		if err != nil {
			return
		}
		if _, err := part.Content(); err != nil {
			t.Fatalf("part.Content() error = %v", err)
		}
	})
}
