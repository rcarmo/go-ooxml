package acceptance

import (
	"archive/zip"
	"bytes"
	"compress/bzip2"
	"context"
	"encoding/binary"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"hash/crc32"
	"io"
	"os"
	"strings"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const bzipAdmissionCaseID = "@id-package-admission-unsupported-compression"
const bzipPayloadHex = "425a6839314159265359996746ea00000099000000800520002000219a68334d173c5dc914e14242659d1ba8"
const contentTypesXML = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`

var bzipAdmissionTexts = []string{
	"a ZIP_BZIP2 archive contains a.xml with UTF-8 text <a/>",
	"the package admission guard checks the archive with default limits",
	"package admission is refused",
	"the admission error contains compression",
}

func guardBZIPAdmissionRule(doc *messages.GherkinDocument) error {
	if doc == nil || doc.Feature == nil || doc.Feature.Name != "Bounded ZIP admission rejects unsafe members and unsupported storage" || len(doc.Feature.Tags) != 1 || doc.Feature.Tags[0].Name != "@planned" {
		return fmt.Errorf("BZIP admission feature drift")
	}
	found := 0
	for _, c := range doc.Feature.Children {
		if c.Background != nil || c.Rule != nil {
			return fmt.Errorf("BZIP admission unexpected background or Rule")
		}
		if s := c.Scenario; s != nil {
			for _, tag := range s.Tags {
				if tag.Name == bzipAdmissionCaseID {
					found++
					if len(s.Tags) != 1 || int(tag.Location.Line) != 33 || int(s.Location.Line) != 34 || len(s.Examples) != 0 || len(s.Steps) != 4 {
						return fmt.Errorf("BZIP admission structure drift")
					}
				}
			}
		}
	}
	if found != 1 {
		return fmt.Errorf("BZIP admission scenario count %d", found)
	}
	return nil
}

func guardBZIPAdmissionCase(id string, p *messages.Pickle, line int) error {
	if id != bzipAdmissionCaseID || p == nil || line != 34 || p.Name != "ZIP admission refuses BZIP2 member compression" || len(p.AstNodeIds) != 1 || len(p.Tags) != 2 || p.Tags[0].Name != "@planned" || p.Tags[1].Name != id || len(p.Steps) != 4 {
		return fmt.Errorf("BZIP admission identity drift")
	}
	for i, want := range bzipAdmissionTexts {
		if p.Steps[i].Text != want || p.Steps[i].Argument != nil {
			return fmt.Errorf("BZIP admission step %d drift", i+1)
		}
	}
	return nil
}

func TestBZIPAdmissionGuard(t *testing.T) {
	path := overlapFeaturePath()
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	n := 0
	next := func() string { n++; return fmt.Sprint(n) }
	doc, err := gherkin.ParseGherkinDocument(f, next)
	if err != nil {
		t.Fatal(err)
	}
	if err := guardBZIPAdmissionRule(doc); err != nil {
		t.Fatal(err)
	}
	var picked *messages.Pickle
	for _, p := range gherkin.Pickles(*doc, path, next) {
		for _, tag := range p.Tags {
			if tag.Name == bzipAdmissionCaseID {
				if picked != nil {
					t.Fatal("duplicate BZIP2 pickle")
				}
				picked = p
			}
		}
	}
	if err := guardBZIPAdmissionCase(bzipAdmissionCaseID, picked, 34); err != nil {
		t.Fatal(err)
	}
	for _, tc := range []struct {
		name   string
		change func(*messages.Pickle)
	}{
		{"name", func(p *messages.Pickle) { p.Name += " changed" }},
		{"id", func(p *messages.Pickle) { p.Tags[1].Name = "@id-other" }},
		{"tags", func(p *messages.Pickle) { p.Tags = append(p.Tags, &messages.PickleTag{Name: "@extra"}) }},
		{"examples", func(p *messages.Pickle) { p.AstNodeIds = append(p.AstNodeIds, "row") }},
		{"argument", func(p *messages.Pickle) { p.Steps[1].Argument = &messages.PickleStepArgument{} }},
	} {
		t.Run(tc.name, func(t *testing.T) {
			p := *picked
			p.Tags = append([]*messages.PickleTag(nil), picked.Tags...)
			for i, tag := range p.Tags {
				x := *tag
				p.Tags[i] = &x
			}
			p.Steps = append([]*messages.PickleStep(nil), picked.Steps...)
			for i, step := range p.Steps {
				x := *step
				p.Steps[i] = &x
			}
			p.AstNodeIds = append([]string(nil), picked.AstNodeIds...)
			tc.change(&p)
			if guardBZIPAdmissionCase(bzipAdmissionCaseID, &p, 34) == nil {
				t.Fatal("accepted scenario drift")
			}
		})
	}
	for i := range bzipAdmissionTexts {
		p := *picked
		p.Steps = append([]*messages.PickleStep(nil), picked.Steps...)
		s := *p.Steps[i]
		s.Text += " changed"
		p.Steps[i] = &s
		if guardBZIPAdmissionCase(bzipAdmissionCaseID, &p, 34) == nil {
			t.Fatalf("accepted step %d drift", i+1)
		}
	}
	if guardBZIPAdmissionCase(bzipAdmissionCaseID, picked, 35) == nil {
		t.Fatal("accepted line drift")
	}
}

// Compressed bytes came from `printf '<a/>' | bzip2 -9 -c` (bzip2 1.0.8),
// independently of Go's archive/zip writer and compress/bzip2 reader.
func bzipArchive(method uint16) ([]byte, error) {
	compressed, err := hex.DecodeString(bzipPayloadHex)
	if err != nil {
		return nil, err
	}
	var out bytes.Buffer
	writer := zip.NewWriter(&out)
	ct, err := writer.CreateHeader(&zip.FileHeader{Name: "[Content_Types].xml", Method: zip.Store})
	if err != nil {
		return nil, err
	}
	if _, err := ct.Write([]byte(contentTypesXML)); err != nil {
		return nil, err
	}
	if method == 12 {
		h := &zip.FileHeader{Name: "a.xml", Method: method, Flags: 0x800, CRC32: crc32.ChecksumIEEE([]byte("<a/>")), CompressedSize64: uint64(len(compressed)), UncompressedSize64: 4}
		w, err := writer.CreateRaw(h)
		if err != nil {
			return nil, err
		}
		if _, err := w.Write(compressed); err != nil {
			return nil, err
		}
	} else {
		w, err := writer.CreateHeader(&zip.FileHeader{Name: "a.xml", Method: method})
		if err != nil {
			return nil, err
		}
		if _, err := w.Write([]byte("<a/>")); err != nil {
			return nil, err
		}
	}
	if err := writer.Close(); err != nil {
		return nil, err
	}
	return bytes.Clone(out.Bytes()), nil
}

// Independently inspect the central and local ZIP headers and decode BZIP2 raw
// bytes; a malformed ZIP or a method-12 header with plaintext must not qualify.
func inspectBZIPArchive(archive []byte) error {
	zr, err := zip.NewReader(bytes.NewReader(archive), int64(len(archive)))
	if err != nil {
		return err
	}
	if len(zr.File) != 2 || zr.File[0].Name != "[Content_Types].xml" || zr.File[1].Name != "a.xml" {
		return fmt.Errorf("wrong ZIP member inventory")
	}
	ct, err := zr.File[0].Open()
	if err != nil {
		return err
	}
	ctBytes, readErr := io.ReadAll(ct)
	closeErr := ct.Close()
	if readErr != nil {
		return readErr
	}
	if closeErr != nil {
		return closeErr
	}
	if !bytes.Equal(ctBytes, []byte(contentTypesXML)) {
		return fmt.Errorf("invalid OPC registry")
	}
	f := zr.File[1]
	if f.Method != 12 || f.Flags&0x800 == 0 || f.CompressedSize64 == 0 || f.UncompressedSize64 != 4 || f.CRC32 != crc32.ChecksumIEEE([]byte("<a/>")) {
		return fmt.Errorf("wrong BZIP2 central metadata")
	}
	offset, err := f.DataOffset()
	if err != nil {
		return err
	}
	if offset < 30 || offset+int64(f.CompressedSize64) > int64(len(archive)) {
		return fmt.Errorf("invalid member extent")
	}
	// Walk the EOCD's central records to obtain the local header offset,
	// independently of the ZIP reader's central-directory file inventory.
	if len(archive) < 22 {
		return fmt.Errorf("missing ZIP end record")
	}
	end := len(archive) - 22
	if binary.LittleEndian.Uint32(archive[end:end+4]) != 0x06054b50 || binary.LittleEndian.Uint16(archive[end+10:end+12]) != 2 || binary.LittleEndian.Uint16(archive[end+20:end+22]) != 0 {
		return fmt.Errorf("invalid ZIP end record")
	}
	central := int(binary.LittleEndian.Uint32(archive[end+16 : end+20]))
	centralSize := int(binary.LittleEndian.Uint32(archive[end+12 : end+16]))
	if central < 0 || centralSize < 0 || central+centralSize != end {
		return fmt.Errorf("invalid central directory geometry")
	}
	local := -1
	for i := 0; i < 2; i++ {
		if central+46 > end || binary.LittleEndian.Uint32(archive[central:central+4]) != 0x02014b50 {
			return fmt.Errorf("invalid central member header")
		}
		nameLen := int(binary.LittleEndian.Uint16(archive[central+28 : central+30]))
		extraLen := int(binary.LittleEndian.Uint16(archive[central+30 : central+32]))
		commentLen := int(binary.LittleEndian.Uint16(archive[central+32 : central+34]))
		next := central + 46 + nameLen + extraLen + commentLen
		if next > end {
			return fmt.Errorf("central member exceeds directory")
		}
		if string(archive[central+46:central+46+nameLen]) == f.Name {
			if binary.LittleEndian.Uint16(archive[central+10:central+12]) != 12 || binary.LittleEndian.Uint16(archive[central+8:central+10])&0x800 == 0 || binary.LittleEndian.Uint32(archive[central+16:central+20]) != f.CRC32 || uint64(binary.LittleEndian.Uint32(archive[central+20:central+24])) != f.CompressedSize64 || uint64(binary.LittleEndian.Uint32(archive[central+24:central+28])) != f.UncompressedSize64 {
				return fmt.Errorf("wrong BZIP2 central declaration")
			}
			local = int(binary.LittleEndian.Uint32(archive[central+42 : central+46]))
		}
		central = next
	}
	if central != end || local < 0 || local+30 > len(archive) || binary.LittleEndian.Uint32(archive[local:local+4]) != 0x04034b50 {
		return fmt.Errorf("missing BZIP2 local header")
	}
	localNameLen := int(binary.LittleEndian.Uint16(archive[local+26 : local+28]))
	localExtraLen := int(binary.LittleEndian.Uint16(archive[local+28 : local+30]))
	if local+30+localNameLen+localExtraLen != int(offset) || local+30+localNameLen > len(archive) || string(archive[local+30:local+30+localNameLen]) != f.Name || binary.LittleEndian.Uint16(archive[local+8:local+10]) != 12 || binary.LittleEndian.Uint16(archive[local+6:local+8])&0x800 == 0 {
		return fmt.Errorf("wrong BZIP2 local method, name or data offset")
	}
	rawReader, err := f.OpenRaw()
	if err != nil {
		return err
	}
	raw, err := io.ReadAll(rawReader)
	if err != nil {
		return err
	}
	if len(raw) != int(f.CompressedSize64) || !bytes.Equal(raw, archive[offset:offset+int64(len(raw))]) {
		return fmt.Errorf("raw member extent differs")
	}
	decoded, err := io.ReadAll(io.LimitReader(bzip2.NewReader(bytes.NewReader(raw)), 5))
	if err != nil {
		return err
	}
	if !bytes.Equal(decoded, []byte("<a/>")) || crc32.ChecksumIEEE(decoded) != f.CRC32 || uint64(len(decoded)) != f.UncompressedSize64 {
		return fmt.Errorf("BZIP2 payload, CRC or size differs")
	}
	return nil
}

func checkBZIPRefusal(archive []byte, open func([]byte) (*packaging.Preserved, error)) error {
	original := bytes.Clone(archive)
	p, err := open(archive)
	if p != nil || err == nil || !strings.Contains(strings.ToLower(err.Error()), "compression") || !bytes.Equal(archive, original) {
		return fmt.Errorf("BZIP2 refusal failed: session=%t error=%v sourceChanged=%t", p != nil, err, !bytes.Equal(archive, original))
	}
	var refusal *packaging.Refusal
	if !errors.As(err, &refusal) || refusal.Kind != "invalid_package" || refusal.Operation != "open" || refusal.Part != "a.xml" {
		return fmt.Errorf("unexpected refusal: %v", err)
	}
	return nil
}

func realBZIPOpen(source []byte) (*packaging.Preserved, error) {
	return packaging.OpenPreserved(source, packaging.Limits{})
}

func TestBZIPAdmissionControls(t *testing.T) {
	source, err := bzipArchive(12)
	if err != nil {
		t.Fatal(err)
	}
	if err := inspectBZIPArchive(source); err != nil {
		t.Fatal(err)
	}
	if err := checkBZIPRefusal(source, realBZIPOpen); err != nil {
		t.Fatal(err)
	}
	for _, method := range []uint16{zip.Store, zip.Deflate} {
		sibling, err := bzipArchive(method)
		if err != nil {
			t.Fatal(err)
		}
		p, err := realBZIPOpen(sibling)
		if err != nil || p == nil {
			t.Fatalf("valid sibling method %d: %v", method, err)
		}
		part, _, err := p.Part("a.xml")
		if err != nil || !bytes.Equal(part, []byte("<a/>")) {
			t.Fatalf("sibling method %d payload: %v", method, err)
		}
	}
	// The method guard cannot be replaced by an unconditional rejection, nor by
	// successful intake of method 12 with a fabricated compression error.
	if err := checkBZIPRefusal(source, func(b []byte) (*packaging.Preserved, error) { return &packaging.Preserved{}, nil }); err == nil {
		t.Fatal("accepted method-12 admission")
	}
	broken := bytes.Clone(source)
	broken[0] = 0
	if err := inspectBZIPArchive(broken); err == nil {
		t.Fatal("accepted malformed source as BZIP2 fixture")
	}
}

func bzipAdmissionSteps(sc *godog.ScenarioContext) {
	var source, original []byte
	var refusal error
	var session *packaging.Preserved
	var unsafePairs [][2]string
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source, original = nil, nil
		refusal = nil
		session = nil
		unsafePairs = nil
		return ctx, nil
	})
	sc.Step(`^an ordered ZIP_STORED archive has member pairs encoded as JSON (.+)$`, func(input string) error {
		var pairs [][2]string
		if err := json.Unmarshal([]byte(input), &pairs); err != nil {
			return err
		}
		selected := false
		for _, row := range unsafeRows {
			if samePairs(pairs, row.pairs) {
				selected = true
			}
		}
		if !selected {
			return fmt.Errorf("unknown unsafe-member pairs")
		}
		var err error
		source, err = rawStoredZIP(pairs)
		if err != nil {
			return err
		}
		original = bytes.Clone(source)
		unsafePairs = pairs
		return inspectStoredZIP(source, pairs)
	})
	sc.Step(`^a ZIP_BZIP2 archive contains a\.xml with UTF-8 text <a/>$`, func() error {
		var err error
		source, err = bzipArchive(12)
		if err != nil {
			return err
		}
		original = bytes.Clone(source)
		return inspectBZIPArchive(source)
	})
	sc.Step(`^the package admission guard checks the archive with default limits$`, func() error {
		if len(source) == 0 || !bytes.Equal(source, original) {
			return fmt.Errorf("missing/changed ZIP source")
		}
		session, refusal = realBZIPOpen(source)
		return nil
	})
	sc.Step(`^package admission is refused$`, func() error {
		if session != nil || refusal == nil || !bytes.Equal(source, original) {
			return fmt.Errorf("admission did not refuse unchanged source: %v", refusal)
		}
		if unsafePairs != nil {
			var typed *packaging.Refusal
			op, part, detail := expectedUnsafeRefusal(unsafePairs)
			if !errors.As(refusal, &typed) || typed.Kind != "invalid_package" || typed.Operation != op || typed.Part != part || !strings.Contains(typed.Detail, detail) {
				return fmt.Errorf("wrong unsafe-member refusal: %v", refusal)
			}
		}
		return nil
	})
	sc.Step(`^the admission error contains compression$`, func() error {
		if session != nil || refusal == nil || !strings.Contains(strings.ToLower(refusal.Error()), "compression") || !bytes.Equal(source, original) {
			return fmt.Errorf("missing compression refusal/custody: %v", refusal)
		}
		var typed *packaging.Refusal
		if !errors.As(refusal, &typed) || typed.Kind != "invalid_package" || typed.Operation != "open" || typed.Part != "a.xml" {
			return fmt.Errorf("wrong compression refusal: %v", refusal)
		}
		return nil
	})
}
