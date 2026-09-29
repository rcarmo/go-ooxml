package acceptance

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"encoding/json"
	"errors"
	"fmt"
	"hash/crc32"
	"io"
	"os"
	"strings"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const unsafeMembersCaseID = "@id-package-admission-unsafe-members"

var unsafeRows = []struct {
	label string
	pairs [][2]string
}{
	{"duplicate member names", [][2]string{{"a.xml", "<a/>"}, {"a.xml", "<b/>"}}},
	{"parent traversal name", [][2]string{{"../a.xml", "<a/>"}}},
	{"absolute member name", [][2]string{{"/a.xml", "<a/>"}}},
	{"backslash member name", [][2]string{{"x\\a.xml", "<a/>"}}},
	{"directory entry with bytes", [][2]string{{"a/", "payload"}}},
}

func guardUnsafeMembersFeature(doc *messages.GherkinDocument) error {
	if doc == nil || doc.Feature == nil || doc.Feature.Name != "Bounded ZIP admission rejects unsafe members and unsupported storage" || len(doc.Feature.Tags) != 1 || doc.Feature.Tags[0].Name != "@planned" {
		return fmt.Errorf("unsafe-member feature drift")
	}
	found := 0
	for _, c := range doc.Feature.Children {
		if c.Background != nil || c.Rule != nil {
			return fmt.Errorf("unsafe-member background/Rule drift")
		}
		if s := c.Scenario; s != nil {
			for _, tag := range s.Tags {
				if tag.Name == unsafeMembersCaseID {
					found++
					if len(s.Tags) != 1 || int(tag.Location.Line) != 6 || int(s.Location.Line) != 7 || s.Name != "ZIP admission refuses <variant>" || len(s.Steps) != 3 || len(s.Examples) != 1 || len(s.Examples[0].TableBody) != 5 {
						return fmt.Errorf("unsafe-member scenario structure drift")
					}
					ex := s.Examples[0]
					if len(ex.Tags) != 0 || len(ex.TableHeader.Cells) != 2 || ex.TableHeader.Cells[0].Value != "variant" || ex.TableHeader.Cells[1].Value != "entries_json" {
						return fmt.Errorf("unsafe-member Examples header drift")
					}
					for i, row := range ex.TableBody {
						if int(row.Location.Line) != 14+i || len(row.Cells) != 2 || row.Cells[0].Value != unsafeRows[i].label {
							return fmt.Errorf("unsafe-member example %d drift", i)
						}
						var pairs [][2]string
						if err := json.Unmarshal([]byte(row.Cells[1].Value), &pairs); err != nil || !samePairs(pairs, unsafeRows[i].pairs) {
							return fmt.Errorf("unsafe-member JSON row %d drift: %v", i, err)
						}
					}
				}
			}
		}
	}
	if found != 1 {
		return fmt.Errorf("unsafe-member scenario count %d", found)
	}
	return nil
}

func samePairs(a, b [][2]string) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i] != b[i] {
			return false
		}
	}
	return true
}

func guardUnsafeMembersCase(id string, p *messages.Pickle, line int) error {
	if id != unsafeMembersCaseID || p == nil || len(p.Tags) != 2 || p.Tags[0].Name != "@planned" || p.Tags[1].Name != id || len(p.AstNodeIds) != 2 || len(p.Steps) != 3 {
		return fmt.Errorf("unsafe-member pickle identity drift")
	}
	for i, row := range unsafeRows {
		if line == 14+i {
			if p.Name != "ZIP admission refuses "+row.label || p.Steps[1].Text != "the package admission guard checks the archive with default limits" || p.Steps[2].Text != "package admission is refused" {
				return fmt.Errorf("unsafe-member step/name drift")
			}
			const prefix = "an ordered ZIP_STORED archive has member pairs encoded as JSON "
			if !strings.HasPrefix(p.Steps[0].Text, prefix) {
				return fmt.Errorf("unsafe-member Given drift")
			}
			var pairs [][2]string
			if err := json.Unmarshal([]byte(strings.TrimPrefix(p.Steps[0].Text, prefix)), &pairs); err != nil || !samePairs(pairs, row.pairs) {
				return fmt.Errorf("unsafe-member pickle row %d drift: %v", i, err)
			}
			for _, step := range p.Steps {
				if step.Argument != nil {
					return fmt.Errorf("unsafe-member unexpected step argument")
				}
			}
			return nil
		}
	}
	return fmt.Errorf("unexpected unsafe-member row line %d", line)
}

func TestUnsafeMembersFeatureGuard(t *testing.T) {
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
	if err := guardUnsafeMembersFeature(doc); err != nil {
		t.Fatal(err)
	}
	seen := map[int]bool{}
	for _, p := range gherkin.Pickles(*doc, path, next) {
		for _, tag := range p.Tags {
			if tag.Name == unsafeMembersCaseID {
				line := -1
				for _, c := range doc.Feature.Children {
					if s := c.Scenario; s != nil && s.Name == "ZIP admission refuses <variant>" {
						for _, ex := range s.Examples {
							for _, row := range ex.TableBody {
								if row.Id == p.AstNodeIds[1] {
									line = int(row.Location.Line)
								}
							}
						}
					}
				}
				if err := guardUnsafeMembersCase(unsafeMembersCaseID, p, line); err != nil {
					t.Fatal(err)
				}
				if seen[line] {
					t.Fatalf("duplicate row %d", line)
				}
				seen[line] = true
				copyPickle := *p
				copyPickle.Steps = append([]*messages.PickleStep(nil), p.Steps...)
				changed := *p.Steps[2]
				changed.Text += " changed"
				copyPickle.Steps[2] = &changed
				if guardUnsafeMembersCase(unsafeMembersCaseID, &copyPickle, line) == nil {
					t.Fatalf("accepted altered row %d", line)
				}
			}
		}
	}
	if len(seen) != 5 {
		t.Fatalf("selected %d unsafe rows", len(seen))
	}
}

// Write simple ZIP32 local and central headers directly: archive/zip.Writer
// discards a directory's payload, hiding the fifth contract input.
func rawStoredZIP(pairs [][2]string) ([]byte, error) {
	var body, central bytes.Buffer
	for _, pair := range pairs {
		name, data := []byte(pair[0]), []byte(pair[1])
		if len(name) > 65535 {
			return nil, fmt.Errorf("name too long")
		}
		offset := body.Len()
		crc := crc32.ChecksumIEEE(data)
		local := make([]byte, 30)
		binary.LittleEndian.PutUint32(local, 0x04034b50)
		binary.LittleEndian.PutUint16(local[4:], 20)
		binary.LittleEndian.PutUint16(local[6:], 0x800)
		binary.LittleEndian.PutUint32(local[14:], crc)
		binary.LittleEndian.PutUint32(local[18:], uint32(len(data)))
		binary.LittleEndian.PutUint32(local[22:], uint32(len(data)))
		binary.LittleEndian.PutUint16(local[26:], uint16(len(name)))
		body.Write(local)
		body.Write(name)
		body.Write(data)
		cd := make([]byte, 46)
		binary.LittleEndian.PutUint32(cd, 0x02014b50)
		binary.LittleEndian.PutUint16(cd[4:], 20)
		binary.LittleEndian.PutUint16(cd[6:], 20)
		binary.LittleEndian.PutUint16(cd[8:], 0x800)
		binary.LittleEndian.PutUint32(cd[16:], crc)
		binary.LittleEndian.PutUint32(cd[20:], uint32(len(data)))
		binary.LittleEndian.PutUint32(cd[24:], uint32(len(data)))
		binary.LittleEndian.PutUint16(cd[28:], uint16(len(name)))
		binary.LittleEndian.PutUint32(cd[42:], uint32(offset))
		central.Write(cd)
		central.Write(name)
	}
	end := make([]byte, 22)
	binary.LittleEndian.PutUint32(end, 0x06054b50)
	binary.LittleEndian.PutUint16(end[8:], uint16(len(pairs)))
	binary.LittleEndian.PutUint16(end[10:], uint16(len(pairs)))
	binary.LittleEndian.PutUint32(end[12:], uint32(central.Len()))
	binary.LittleEndian.PutUint32(end[16:], uint32(body.Len()))
	return bytes.Join([][]byte{body.Bytes(), central.Bytes(), end}, nil), nil
}

// Validate the raw input independently of production intake; distinguish an
// unsafe name or directory payload from malformed ZIP geometry or CRC.
func inspectStoredZIP(data []byte, pairs [][2]string) error {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return err
	}
	if len(zr.File) != len(pairs) || len(data) < 22 {
		return fmt.Errorf("member count/end record")
	}
	end := len(data) - 22
	if binary.LittleEndian.Uint32(data[end:]) != 0x06054b50 || int(binary.LittleEndian.Uint16(data[end+10:])) != len(pairs) || binary.LittleEndian.Uint16(data[end+20:]) != 0 {
		return fmt.Errorf("ZIP end drift")
	}
	pos := int(binary.LittleEndian.Uint32(data[end+16:]))
	limit := pos + int(binary.LittleEndian.Uint32(data[end+12:]))
	if limit != end {
		return fmt.Errorf("central extent drift")
	}
	cursor := 0
	for i, pair := range pairs {
		name, payload := []byte(pair[0]), []byte(pair[1])
		f := zr.File[i]
		if pos+46 > limit || binary.LittleEndian.Uint32(data[pos:]) != 0x02014b50 {
			return fmt.Errorf("central header %d", i)
		}
		n := int(binary.LittleEndian.Uint16(data[pos+28:]))
		x := int(binary.LittleEndian.Uint16(data[pos+30:]))
		c := int(binary.LittleEndian.Uint16(data[pos+32:]))
		next := pos + 46 + n + x + c
		if next > limit || n != len(name) || !bytes.Equal(data[pos+46:pos+46+n], name) || int(binary.LittleEndian.Uint32(data[pos+42:])) != cursor {
			return fmt.Errorf("central name/offset %d", i)
		}
		crc := crc32.ChecksumIEEE(payload)
		if binary.LittleEndian.Uint16(data[pos+10:]) != 0 || binary.LittleEndian.Uint16(data[pos+8:]) != 0x800 || binary.LittleEndian.Uint32(data[pos+16:]) != crc || int(binary.LittleEndian.Uint32(data[pos+20:])) != len(payload) || int(binary.LittleEndian.Uint32(data[pos+24:])) != len(payload) || f.Name != pair[0] || f.Method != zip.Store || f.CRC32 != crc || f.UncompressedSize64 != uint64(len(payload)) {
			return fmt.Errorf("central method/CRC/size %d", i)
		}
		if cursor+30 > len(data) || binary.LittleEndian.Uint32(data[cursor:]) != 0x04034b50 {
			return fmt.Errorf("local header %d", i)
		}
		ln := int(binary.LittleEndian.Uint16(data[cursor+26:]))
		lx := int(binary.LittleEndian.Uint16(data[cursor+28:]))
		payloadStart := cursor + 30 + ln + lx
		payloadEnd := payloadStart + len(payload)
		if ln != len(name) || lx != 0 || payloadEnd > int(binary.LittleEndian.Uint32(data[end+16:])) || !bytes.Equal(data[cursor+30:cursor+30+ln], name) || binary.LittleEndian.Uint16(data[cursor+8:]) != 0 || binary.LittleEndian.Uint16(data[cursor+6:]) != 0x800 || binary.LittleEndian.Uint32(data[cursor+14:]) != crc || int(binary.LittleEndian.Uint32(data[cursor+18:])) != len(payload) || int(binary.LittleEndian.Uint32(data[cursor+22:])) != len(payload) || !bytes.Equal(data[payloadStart:payloadEnd], payload) {
			return fmt.Errorf("local method/CRC/physical payload %d", i)
		}
		offset, err := f.DataOffset()
		if err != nil || offset != int64(payloadStart) {
			return fmt.Errorf("ZIP reader offset %d: %v", i, err)
		}
		if !strings.HasSuffix(pair[0], "/") {
			reader, err := f.Open()
			if err != nil {
				return err
			}
			decoded, readErr := io.ReadAll(reader)
			closeErr := reader.Close()
			if readErr != nil {
				return readErr
			}
			if closeErr != nil {
				return closeErr
			}
			if !bytes.Equal(decoded, payload) {
				return fmt.Errorf("independent ZIP payload %d", i)
			}
		}
		cursor = payloadEnd
		pos = next
	}
	if pos != limit || cursor != int(binary.LittleEndian.Uint32(data[end+16:])) {
		return fmt.Errorf("ZIP member geometry drift")
	}
	return nil
}

func expectedUnsafeRefusal(pairs [][2]string) (operation, part, detail string) {
	switch {
	case samePairs(pairs, unsafeRows[0].pairs):
		return "open", "a.xml", "duplicate"
	case samePairs(pairs, unsafeRows[4].pairs):
		return "open", "a/", "directory entry has payload bytes"
	default:
		return "validate", pairs[0][0], "unsafe or non-canonical"
	}
}

func checkUnsafeAdmission(data []byte, pairs [][2]string) error {
	original := bytes.Clone(data)
	p, err := packaging.OpenPreserved(data, packaging.Limits{})
	if p != nil || err == nil || !bytes.Equal(data, original) {
		return fmt.Errorf("unsafe admission: session=%t error=%v sourceChanged=%t", p != nil, err, !bytes.Equal(data, original))
	}
	var refusal *packaging.Refusal
	op, part, detail := expectedUnsafeRefusal(pairs)
	if !errors.As(err, &refusal) || refusal.Kind != "invalid_package" || refusal.Operation != op || refusal.Part != part || !strings.Contains(refusal.Detail, detail) {
		return fmt.Errorf("wrong unsafe-member refusal: %v", err)
	}
	return nil
}

func TestUnsafeMembersControls(t *testing.T) {
	for _, row := range unsafeRows {
		data, err := rawStoredZIP(row.pairs)
		if err != nil {
			t.Fatal(err)
		}
		if err := inspectStoredZIP(data, row.pairs); err != nil {
			t.Fatalf("%s: %v", row.label, err)
		}
		if err := checkUnsafeAdmission(data, row.pairs); err != nil {
			t.Fatalf("%s: %v", row.label, err)
		}
	}
	validPairs := [][2]string{{"[Content_Types].xml", contentTypesXML}, {"a.xml", "<a/>"}, {"empty/", ""}}
	valid, err := rawStoredZIP(validPairs)
	if err != nil {
		t.Fatal(err)
	}
	if err := inspectStoredZIP(valid, validPairs); err != nil {
		t.Fatal(err)
	}
	original := bytes.Clone(valid)
	pkg, err := packaging.OpenPreserved(valid, packaging.Limits{})
	if err != nil || pkg == nil {
		t.Fatalf("valid STORED OPC: %v", err)
	}
	part, _, err := pkg.Part("a.xml")
	if err != nil || !bytes.Equal(part, []byte("<a/>")) || !bytes.Equal(valid, original) {
		t.Fatalf("positive OPC payload/custody: %v", err)
	}
	corrupt := bytes.Clone(valid)
	corrupt[0] = 0
	if err := inspectStoredZIP(corrupt, validPairs); err == nil {
		t.Fatal("malformed geometry passed independent guard")
	}
}
