package acceptance

import (
	"archive/zip"
	"bytes"
	"context"
	"encoding/binary"
	"errors"
	"fmt"
	"os"
	"path/filepath"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

var injected = errors.New("injected stream failure")

type failedWriter struct{}

func (failedWriter) Write([]byte) (int, error) { return 0, injected }

type safetyWorld struct {
	source, original   []byte
	pkg                *packaging.Package
	err                error
	dir, dest, oldPath string
	modified           bool
}

func (s *safetyWorld) archive(kind string) error {
	names := map[string][]string{"duplicate": {"data.xml", "data.xml"}, "collision": {"data.xml", "DATA.xml"}, "traversal": {"../data.xml"}, "absolute": {"/data.xml"}, "backslash": {`a\data.xml`}}[kind]
	if names == nil {
		return fmt.Errorf("unknown kind %s", kind)
	}
	var b bytes.Buffer
	z := zip.NewWriter(&b)
	w, err := z.Create("[Content_Types].xml")
	if err != nil {
		return err
	}
	_, err = w.Write([]byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`))
	if err != nil {
		return err
	}
	for _, name := range names {
		w, err = z.Create(name)
		if err != nil {
			return err
		}
		if _, err = w.Write([]byte(`<data/>`)); err != nil {
			return err
		}
	}
	if err = z.Close(); err != nil {
		return err
	}
	s.source = b.Bytes()
	s.original = bytes.Clone(s.source)
	return nil
}
func (s *safetyWorld) attemptOpen() error { s.pkg, s.err = packaging.OpenBytes(s.source); return nil }
func (s *safetyWorld) refused() error {
	if s.err == nil {
		return fmt.Errorf("unsafe archive accepted")
	}
	if !bytes.Equal(s.source, s.original) {
		return fmt.Errorf("source mutated")
	}
	return nil
}
func (s *safetyWorld) small() error {
	s.pkg = packaging.New()
	_, err := s.pkg.AddPart("data.xml", packaging.ContentTypeXML, []byte(`<data/>`))
	return err
}
func (s *safetyWorld) broken() error {
	if err := s.small(); err != nil {
		return err
	}
	var err error
	s.dir, err = os.MkdirTemp("", "ooxml-acceptance-")
	if err != nil {
		return err
	}
	s.dest = filepath.Join(s.dir, "existing.docx")
	s.original = []byte("previous destination")
	if err = os.WriteFile(s.dest, s.original, 0600); err != nil {
		return err
	}
	if _, err = s.pkg.AddPart(`unsafe\path.xml`, packaging.ContentTypeXML, []byte(`<data/>`)); err != nil {
		return err
	}
	s.oldPath = s.pkg.Path()
	s.modified = s.pkg.IsModified()
	return nil
}
func (s *safetyWorld) savedUnchanged() error {
	if s.err == nil {
		return fmt.Errorf("unsafe member name saved")
	}
	got, err := os.ReadFile(s.dest)
	if err != nil {
		return err
	}
	if !bytes.Equal(got, s.original) {
		return fmt.Errorf("destination truncated or replaced")
	}
	if s.pkg.Path() != s.oldPath || s.pkg.IsModified() != s.modified {
		return fmt.Errorf("package state changed")
	}
	return nil
}
func safetySteps(sc *godog.ScenarioContext) {
	s := &safetyWorld{}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		*s = safetyWorld{}
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, err error) (context.Context, error) {
		if s.pkg != nil {
			_ = s.pkg.Close()
		}
		if s.dir != "" {
			if cleanup := os.RemoveAll(s.dir); err == nil {
				err = cleanup
			}
		}
		return ctx, nil
	})
	sc.Step(`^a synthetic archive with "([^"]+)" member names$`, s.archive)
	sc.Step(`^a synthetic archive with an adjusted prepended prefix$`, func() error {
		if err := s.small(); err != nil {
			return err
		}
		var b bytes.Buffer
		if err := s.pkg.WriteTo(&b); err != nil {
			return err
		}
		prefix := []byte("PREFIX")
		s.source = append(bytes.Clone(prefix), b.Bytes()...)
		for at := len(prefix); ; {
			i := bytes.Index(s.source[at:], []byte{'P', 'K', 1, 2})
			if i < 0 {
				break
			}
			i += at
			old := binary.LittleEndian.Uint32(s.source[i+42 : i+46])
			binary.LittleEndian.PutUint32(s.source[i+42:i+46], old+uint32(len(prefix)))
			at = i + 46
		}
		end := bytes.LastIndex(s.source, []byte{'P', 'K', 5, 6})
		old := binary.LittleEndian.Uint32(s.source[end+16 : end+20])
		binary.LittleEndian.PutUint32(s.source[end+16:end+20], old+uint32(len(prefix)))
		s.original = bytes.Clone(s.source)
		return nil
	})
	sc.Step(`^a synthetic archive with a mismatched local "([^"]+)"$`, func(field string) error {
		if err := s.small(); err != nil {
			return err
		}
		var b bytes.Buffer
		if err := s.pkg.WriteTo(&b); err != nil {
			return err
		}
		s.source = bytes.Clone(b.Bytes())
		switch field {
		case "name":
			s.source[30] ^= 1
		case "method":
			binary.LittleEndian.PutUint16(s.source[8:10], zip.Store)
		case "flags":
			s.source[6] ^= 1
		default:
			return fmt.Errorf("unknown field %s", field)
		}
		s.original = bytes.Clone(s.source)
		return nil
	})
	sc.Step(`^I attempt to open the synthetic archive$`, s.attemptOpen)
	sc.Step(`^archive intake fails without changing the source bytes$`, s.refused)
	sc.Step(`^a small new Office package$`, s.small)
	sc.Step(`^I write it to a stream that fails during finalisation$`, func() error { s.err = s.pkg.WriteTo(failedWriter{}); return nil })
	sc.Step(`^the write reports the injected failure$`, func() error {
		if !errors.Is(s.err, injected) {
			return fmt.Errorf("want injected failure, got %v", s.err)
		}
		return nil
	})
	sc.Step(`^an existing destination and an unserialisable package$`, s.broken)
	sc.Step(`^I attempt to save over the destination$`, func() error { s.err = s.pkg.SaveAs(s.dest); return nil })
	sc.Step(`^saving fails and the destination and package state remain unchanged$`, s.savedUnchanged)
	sc.Step(`^no temporary files remain in the destination directory$`, func() error {
		files, err := os.ReadDir(s.dir)
		if err != nil {
			return err
		}
		if len(files) != 1 || files[0].Name() != filepath.Base(s.dest) {
			return fmt.Errorf("unexpected files: %v", files)
		}
		return nil
	})
}
