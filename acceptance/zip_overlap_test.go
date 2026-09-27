package acceptance

import (
	"archive/zip"
	"bytes"
	"context"
	"encoding/binary"
	"errors"
	"fmt"
	"hash/crc32"
	"io"
	"os"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const overlapFixtureID = "fixture-9286fc07c3f8698f9637cf9b7a0c60461d1f304753ba69b39f655c8c348ea027"

// The standard-library ZIP reader proves readability independently of the
// package admission guard. Local extents are computed from the source bytes.
func overlapSteps(sc *godog.ScenarioContext) {
	var source, original []byte
	var reader *zip.Reader
	var files map[string]*zip.File
	var ranges map[string][2]int64
	var session *packaging.Preserved
	var admissionErr error
	var output []byte
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source, original, output = nil, nil, nil
		reader, files, ranges, session, admissionErr = nil, nil, nil, nil, nil
		return ctx, nil
	})
	sc.Step(`^fixture fixture-9286fc07c3f8698f9637cf9b7a0c60461d1f304753ba69b39f655c8c348ea027 has exactly three distinct STORED members \[Content_Types\]\.xml, outer\.bin and inner\.bin$`, func() error {
		path, err := testutil.LookupFixture(overlapFixtureID)
		if err != nil {
			return err
		}
		source, err = os.ReadFile(path)
		if err != nil {
			return err
		}
		original = bytes.Clone(source)
		if len(source) != 483 {
			return fmt.Errorf("fixture size: %d", len(source))
		}
		reader, err = zip.NewReader(bytes.NewReader(source), int64(len(source)))
		if err != nil {
			return err
		}
		files = map[string]*zip.File{}
		for _, f := range reader.File {
			if f.Method != zip.Store || files[f.Name] != nil {
				return fmt.Errorf("duplicate or non-STORED member %s", f.Name)
			}
			files[f.Name] = f
		}
		if len(files) != 3 || files["[Content_Types].xml"] == nil || files["outer.bin"] == nil || files["inner.bin"] == nil {
			return fmt.Errorf("unexpected member inventory")
		}
		return nil
	})
	sc.Step(`^an independent ZIP reader opens all three members with declared lengths 149, 49 and 10 bytes and matching CRC32 d694f44a, 32c80458 and 4daa6380$`, func() error {
		for _, item := range []struct {
			name string
			size uint64
			crc  uint32
		}{
			{"[Content_Types].xml", 149, 0xd694f44a}, {"outer.bin", 49, 0x32c80458}, {"inner.bin", 10, 0x4daa6380},
		} {
			f := files[item.name]
			if f == nil || f.UncompressedSize64 != item.size || f.CompressedSize64 != item.size || f.CRC32 != item.crc {
				return fmt.Errorf("wrong declared member metadata %s", item.name)
			}
			r, err := f.Open()
			if err != nil {
				return err
			}
			data, readErr := io.ReadAll(r)
			closeErr := r.Close()
			if readErr != nil || closeErr != nil {
				return fmt.Errorf("member %s read: %v / %v", item.name, readErr, closeErr)
			}
			if uint64(len(data)) != item.size || crc32.ChecksumIEEE(data) != item.crc {
				return fmt.Errorf("member %s length/CRC mismatch", item.name)
			}
		}
		return nil
	})
	sc.Step(`^inner\.bin's complete local header and payload lie within outer\.bin's physical payload range in this ZIP32 single-disk archive$`, func() error {
		if len(source) < 22 {
			return fmt.Errorf("short ZIP")
		}
		eocd := source[len(source)-22:]
		le := binary.LittleEndian
		if le.Uint32(eocd) != 0x06054b50 || le.Uint16(eocd[4:]) != 0 || le.Uint16(eocd[6:]) != 0 || le.Uint16(eocd[8:]) != 3 || le.Uint16(eocd[10:]) != 3 || le.Uint16(eocd[20:]) != 0 {
			return fmt.Errorf("not a three-entry ZIP32 single-disk archive")
		}
		directory := int64(le.Uint32(eocd[16:]))
		if directory+int64(le.Uint32(eocd[12:])) != int64(len(source)-22) {
			return fmt.Errorf("directory extent mismatch")
		}
		ranges = map[string][2]int64{}
		payloads := map[string][2]int64{}
		for name, f := range files {
			offset, err := f.DataOffset()
			if err != nil {
				return err
			}
			start := offset - int64(30+len(name)+len(f.Extra))
			if start < 0 || offset > directory || start+30 > int64(len(source)) || le.Uint32(source[start:]) != 0x04034b50 {
				return fmt.Errorf("invalid local header %s", name)
			}
			nameLen := int64(le.Uint16(source[start+26:]))
			extraLen := int64(le.Uint16(source[start+28:]))
			if start+30+nameLen+extraLen != offset || nameLen != int64(len(name)) || !bytes.Equal(source[start+30:start+30+nameLen], []byte(name)) || le.Uint16(source[start+8:]) != zip.Store || le.Uint16(source[start+6:]) != 0 {
				return fmt.Errorf("local metadata mismatch %s", name)
			}
			end := offset + int64(f.CompressedSize64)
			if end > directory {
				return fmt.Errorf("payload exceeds directory %s", name)
			}
			ranges[name], payloads[name] = [2]int64{start, end}, [2]int64{offset, end}
		}
		inner, outer := ranges["inner.bin"], payloads["outer.bin"]
		if !(outer[0] <= inner[0] && inner[0] < inner[1] && inner[1] <= outer[1]) {
			return fmt.Errorf("inner header/payload not nested within outer payload: %v %v", inner, outer)
		}
		return nil
	})
	sc.Step(`^bounded package admission checks the unchanged archive with max entries 4 and max total bytes 4096$`, func() error {
		if !bytes.Equal(source, original) || ranges == nil {
			return fmt.Errorf("unverified or changed source")
		}
		session, admissionErr = packaging.OpenPreserved(source, packaging.Limits{MaxSourceBytes: 1 << 20, MaxEntries: 4, MaxPartBytes: 4096, MaxTotalBytes: 4096})
		return nil
	})
	sc.Step(`^admission refuses overlapping physical member extents as an invalid package, not a name, CRC or resource refusal$`, func() error {
		var refusal *packaging.Refusal
		if !errors.As(admissionErr, &refusal) || refusal.Kind != "invalid_package" || refusal.Operation != "open" || !strings.Contains(refusal.Detail, "overlapping archive entries") {
			return fmt.Errorf("wrong overlap refusal: %v", admissionErr)
		}
		return nil
	})
	sc.Step(`^no package session or output archive is delivered$`, func() error {
		if session != nil || output != nil {
			return fmt.Errorf("package or output delivered")
		}
		return nil
	})
	sc.Step(`^the caller's source buffer remains byte-identical to the sealed fixture$`, func() error {
		if !bytes.Equal(source, original) {
			return fmt.Errorf("caller bytes changed")
		}
		return nil
	})
}
