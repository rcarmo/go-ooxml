package acceptance

import (
	"archive/zip"
	"bytes"
	"compress/flate"
	"context"
	"encoding/binary"
	"encoding/json"
	"errors"
	"fmt"
	"hash/crc32"
	"os"
	"path/filepath"
	"reflect"
	"unicode/utf16"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

type batch2State struct {
	archive, initial, output      []byte
	entries, read                 []packaging.ZIP32Entry
	edited, baseline, reopened    *packaging.Preserved
	envelope                      *packaging.Envelope
	originalParts, savedParts     map[string][]byte
	failure                       error
	result                        any
	crc                           uint32
	left, right                   []byte
	same                          bool
	diff                          packaging.PackageDiff
	format, target, value, reason string
	officeIDs                     []string
	corpusGo, corpusPy            []corpusEntry
	goCount, pyCount              int
	checks                        map[string]string
	temp                          string
}

// Build the selected ZIP32 reader geometry independently of production APIs:
// local and central CRC/size fields are populated and bit 3 is clear.
func (s *batch2State) build(entries []packaging.ZIP32Entry, methods []uint16, comment string) error {
	var body, central bytes.Buffer
	put16 := func(b *bytes.Buffer, n uint16) { _ = binary.Write(b, binary.LittleEndian, n) }
	put32 := func(b *bytes.Buffer, n uint32) { _ = binary.Write(b, binary.LittleEndian, n) }
	for i, e := range entries {
		method := uint16(zip.Store)
		if methods != nil {
			method = methods[i]
		}
		if method != zip.Store && method != zip.Deflate {
			return fmt.Errorf("ZIP fixture method %d", method)
		}
		packed := bytes.Clone(e.Data)
		if method == zip.Deflate {
			var b bytes.Buffer
			w, err := flate.NewWriter(&b, flate.DefaultCompression)
			if err != nil {
				return err
			}
			if _, err = w.Write(e.Data); err != nil {
				return err
			}
			if err = w.Close(); err != nil {
				return err
			}
			packed = b.Bytes()
		}
		if len(e.Name) > 65535 || uint64(len(packed)) >= 1<<32 || uint64(len(e.Data)) >= 1<<32 || uint64(body.Len()) >= 1<<32 {
			return fmt.Errorf("ZIP fixture exceeds ZIP32")
		}
		crc := crc32.ChecksumIEEE(e.Data)
		offset := body.Len()
		flags := uint16(0x800)
		put32(&body, 0x04034b50)
		put16(&body, 20)
		put16(&body, flags)
		put16(&body, method)
		put16(&body, 0)
		put16(&body, 0)
		put32(&body, crc)
		put32(&body, uint32(len(packed)))
		put32(&body, uint32(len(e.Data)))
		put16(&body, uint16(len(e.Name)))
		put16(&body, 0)
		body.WriteString(e.Name)
		body.Write(packed)
		put32(&central, 0x02014b50)
		put16(&central, 20)
		put16(&central, 20)
		put16(&central, flags)
		put16(&central, method)
		put16(&central, 0)
		put16(&central, 0)
		put32(&central, crc)
		put32(&central, uint32(len(packed)))
		put32(&central, uint32(len(e.Data)))
		put16(&central, uint16(len(e.Name)))
		for j := 0; j < 4; j++ {
			put16(&central, 0)
		}
		put32(&central, 0)
		put32(&central, uint32(offset))
		central.WriteString(e.Name)
	}
	if len(entries) >= 65535 || len(comment) > 65535 || uint64(body.Len()+central.Len()+22+len(comment)) >= 1<<32 {
		return fmt.Errorf("ZIP fixture end exceeds ZIP32")
	}
	archive := append([]byte{}, body.Bytes()...)
	centralAt := len(archive)
	archive = append(archive, central.Bytes()...)
	var end bytes.Buffer
	put32(&end, 0x06054b50)
	put16(&end, 0)
	put16(&end, 0)
	put16(&end, uint16(len(entries)))
	put16(&end, uint16(len(entries)))
	put32(&end, uint32(central.Len()))
	put32(&end, uint32(centralAt))
	put16(&end, uint16(len(comment)))
	end.WriteString(comment)
	archive = append(archive, end.Bytes()...)
	s.archive = archive
	s.initial = bytes.Clone(archive)
	s.entries = entries
	return nil
}
func (s *batch2State) openGraphFixture(id string, add bool) error {
	path, err := testutil.LookupFixture(id)
	if err != nil {
		return err
	}
	s.archive, err = os.ReadFile(path)
	if err != nil {
		return err
	}
	s.initial = bytes.Clone(s.archive)
	s.originalParts, err = opcZipMembers(s.archive)
	if err != nil {
		return err
	}
	s.edited, err = packaging.OpenPreserved(s.archive, packaging.Limits{})
	if err != nil {
		return err
	}
	if _, err = s.edited.Graph(); err != nil {
		return err
	}
	if add {
		if err = s.addGraphPart(); err != nil {
			return err
		}
		var b bytes.Buffer
		if err = s.edited.WriteTo(&b); err != nil {
			return err
		}
		s.output = b.Bytes()
		s.savedParts, err = opcZipMembers(s.output)
		if err != nil {
			return err
		}
	}
	return nil
}
func (s *batch2State) addGraphPart() error {
	change := packaging.GraphMutation{Additions: []packaging.PartAddition{{Name: "custom/data.bin", ContentType: "application/octet-stream", Data: []byte{7, 8, 9}}}, Relationships: []packaging.RelationshipAddition{{Source: "", ID: "rIdData", Type: "urn:test/data", TargetPart: "custom/data.bin"}}}
	plan, err := s.edited.PlanGraphMutation(change)
	if err != nil {
		return err
	}
	return s.edited.ApplyGraphPlan(plan)
}
func (s *batch2State) saveGraph() error {
	path := filepath.Join(s.temp, "result.docx")
	if _, err := s.edited.SaveAs(path); err != nil {
		return err
	}
	var err error
	s.output, err = os.ReadFile(path)
	if err != nil {
		return err
	}
	s.savedParts, err = opcZipMembers(s.output)
	if err != nil {
		return err
	}
	s.reopened, err = packaging.OpenPreserved(s.output, packaging.Limits{})
	if err != nil {
		return err
	}
	_, err = s.reopened.Graph()
	return err
}
func (s *batch2State) fixedEnvelope() error {
	var err error
	types := `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="` + packaging.NSContentTypes + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>`
	rels := `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="` + packaging.NSRelationships + `"><Relationship Id="rId1" Type="` + packaging.RelTypeOfficeDocument + `" Target="doc/main.xml"/></Relationships>`
	entries := []packaging.ZIP32Entry{{Name: packaging.ContentTypesPath, Data: []byte(types)}, {Name: packaging.PackageRelsPath, Data: []byte(rels)}, {Name: "doc/main.xml", Data: []byte(`<?xml version="1.0" encoding="UTF-8"?><main xmlns="urn:acceptance"><value>Original</value></main>`)}, {Name: "custom/data.bin", Data: []byte{0, 255, 1, 254, 2, 253}}}
	if err = s.build(entries, nil, ""); err != nil {
		return err
	}
	s.originalParts, err = opcZipMembers(s.archive)
	if err != nil {
		return err
	}
	s.envelope, err = packaging.OpenEnvelope(s.archive)
	return err
}
func batch2Reason(err error, want string) error {
	var p *packaging.ProfileError
	if !errors.As(err, &p) || p.Reason != want {
		return fmt.Errorf("reason=%v want %s", err, want)
	}
	return nil
}
func batch2Refusal(err error, want string) error {
	var p *packaging.Refusal
	if !errors.As(err, &p) || p.Kind != want {
		return fmt.Errorf("refusal=%v want %s", err, want)
	}
	return nil
}
func batch2Entries(got, want []packaging.ZIP32Entry) error {
	if len(got) != len(want) {
		return fmt.Errorf("entry count %d want %d", len(got), len(want))
	}
	for i := range got {
		if got[i].Name != want[i].Name || !bytes.Equal(got[i].Data, want[i].Data) {
			return fmt.Errorf("entry %d mismatch", i)
		}
	}
	return nil
}
func batch2SameParts(before, after map[string][]byte, exclude ...string) error {
	skip := map[string]bool{}
	for _, key := range exclude {
		skip[key] = true
	}
	for key, value := range before {
		if skip[key] {
			continue
		}
		if !bytes.Equal(value, after[key]) {
			return fmt.Errorf("member %s changed/missing", key)
		}
	}
	return nil
}
func batch2Steps(sc *godog.ScenarioContext) {
	s := &batch2State{}
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		*s = batch2State{checks: map[string]string{}}
		var err error
		s.temp, err = os.MkdirTemp("", "go-batch2-*")
		return ctx, err
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, err error) (context.Context, error) {
		if s.temp != "" {
			_ = os.RemoveAll(s.temp)
		}
		return ctx, nil
	})
	batch2MoreSteps(sc, s)
	batch2CustodySteps(sc, s)
	batch2OfficeSteps(sc, s)

	sc.Step(`^checksum input is the UTF-8 string "123456789"$`, func() error { s.entries = []packaging.ZIP32Entry{{Data: []byte("123456789")}}; return nil })
	sc.Step(`^its ZIP CRC32 is calculated$`, func() error { s.crc = packaging.ZIPCRC32(s.entries[0].Data); return nil })
	sc.Step(`^the unsigned checksum equals hexadecimal CBF43926$`, func() error {
		if s.crc != 0xcbf43926 || s.crc != crc32.ChecksumIEEE([]byte("123456789")) {
			return fmt.Errorf("CRC32 %08x", s.crc)
		}
		return nil
	})

	sc.Step(`^a single-disk UTF-8 ZIP32 archive has these central-directory ordered members and no data descriptors$`, func(table *godog.Table) error {
		if len(table.Rows) != 5 {
			return fmt.Errorf("ZIP member table drift")
		}
		entries := []packaging.ZIP32Entry{}
		methods := []uint16{}
		for _, row := range table.Rows[1:] {
			if len(row.Cells) != 3 {
				return fmt.Errorf("ZIP row width")
			}
			var value string
			if err := json.Unmarshal([]byte(row.Cells[2].Value), &value); err != nil {
				return err
			}
			entries = append(entries, packaging.ZIP32Entry{Name: row.Cells[0].Value, Data: []byte(value)})
			switch row.Cells[1].Value {
			case "STORED":
				methods = append(methods, zip.Store)
			case "DEFLATED":
				methods = append(methods, zip.Deflate)
			default:
				return fmt.Errorf("ZIP method")
			}
		}
		if !reflect.DeepEqual([]string{entries[0].Name, entries[1].Name, entries[2].Name, entries[3].Name}, []string{"z.bin", "dir/", "a.xml", "m.bin"}) {
			return fmt.Errorf("ZIP member order drift")
		}
		return s.build(entries, methods, "kept as declared ZIP comment")
	})
	sc.Step(`^its declared archive comment is JSON "kept as declared ZIP comment" and all CRC32 and sizes agree with the exact UTF-8 payloads$`, func() error {
		zr, err := zip.NewReader(bytes.NewReader(s.archive), int64(len(s.archive)))
		if err != nil {
			return err
		}
		if zr.Comment != "kept as declared ZIP comment" || len(zr.File) != 4 {
			return fmt.Errorf("ZIP comment/entry count")
		}
		for i, f := range zr.File {
			if f.CRC32 != crc32.ChecksumIEEE(s.entries[i].Data) || int(f.UncompressedSize64) != len(s.entries[i].Data) {
				return fmt.Errorf("ZIP CRC/size %d", i)
			}
		}
		return nil
	})
	sc.Step(`^the production ZIP reader reads the archive with default limits$`, func() { s.read, s.failure = packaging.ReadZIP32(s.archive, packaging.ZIP32Limits{}) })
	sc.Step(`^returned file names are exactly \["z.bin","a.xml","m.bin"\] in central-directory order and dir/ is absent$`, func() error {
		if s.failure != nil || len(s.read) != 3 {
			return fmt.Errorf("ZIP read: %v", s.failure)
		}
		for i, name := range []string{"z.bin", "a.xml", "m.bin"} {
			if s.read[i].Name != name {
				return fmt.Errorf("ZIP order %d", i)
			}
		}
		return nil
	})
	sc.Step(`^the three returned payloads equal their declared strings with lengths 6, 4 and 6 and matching independent CRC32 values$`, func() error {
		want := []string{"z-last", "<a/>", "middle"}
		for i, e := range s.read {
			if string(e.Data) != want[i] || len(e.Data) != []int{6, 4, 6}[i] || packaging.ZIPCRC32(e.Data) != crc32.ChecksumIEEE([]byte(want[i])) {
				return fmt.Errorf("ZIP payload %d", i)
			}
		}
		return nil
	})
	sc.Step(`^the caller's source archive bytes remain unchanged$`, func() error {
		if !bytes.Equal(s.archive, s.initial) {
			return fmt.Errorf("ZIP caller changed")
		}
		return nil
	})

	sc.Step(`^the eleven strict ZIP32 unsafe-structure recipes duplicate, case-collision, traversal, encryption, method, multi-disk, missing-ZIP64, local-name, CRC, stored-size and trailing-byte$`, func() { s.checks = map[string]string{} })
	sc.Step(`^the production ZIP reader checks every recipe with default limits$`, func() error {
		for _, row := range zipReaderRows {
			if row.variant == "deflate size overrun" {
				continue
			}
			entries, err := decodeReasonEntries(row.input)
			if err != nil {
				return err
			}
			archive, err := rawZIPReasonSample(entries, row.mutation)
			if err != nil {
				return err
			}
			original := bytes.Clone(archive)
			result, err := packaging.ReadZIP32(archive, packaging.ZIP32Limits{})
			if result != nil || !bytes.Equal(archive, original) {
				return fmt.Errorf("partial ZIP result/caller mutation: %s", row.variant)
			}
			if err = batch2Reason(err, row.reason); err != nil {
				return fmt.Errorf("%s: %w", row.variant, err)
			}
			s.checks[row.variant] = row.reason
		}
		return nil
	})
	sc.Step(`^each recipe refuses with its exact documented zip reason and no member result$`, func() error {
		if len(s.checks) != 11 {
			return fmt.Errorf("ZIP unsafe recipe count %d", len(s.checks))
		}
		return nil
	})
	sc.Step(`^duplicate and case-collision refuse as zip-duplicate-entry and zip-case-collision$`, func() error {
		return batch2Check(s.checks, map[string]string{"duplicate names": "zip-duplicate-entry", "ASCII case collision": "zip-case-collision"})
	})
	sc.Step(`^traversal refuses as zip-name-invalid$`, func() error { return batch2Check(s.checks, map[string]string{"parent traversal": "zip-name-invalid"}) })
	sc.Step(`^encryption and method refuse as zip-encryption-unsupported and zip-method-unsupported$`, func() error {
		return batch2Check(s.checks, map[string]string{"encryption flag": "zip-encryption-unsupported", "unsupported method": "zip-method-unsupported"})
	})
	sc.Step(`^multi-disk and missing-ZIP64 refuse as zip-multi-disk-unsupported and zip-structure-invalid$`, func() error {
		return batch2Check(s.checks, map[string]string{"multiple disks": "zip-multi-disk-unsupported", "missing ZIP64 records": "zip-structure-invalid"})
	})
	sc.Step(`^local-name refuses as zip-local-metadata-mismatch$`, func() error {
		return batch2Check(s.checks, map[string]string{"local name mismatch": "zip-local-metadata-mismatch"})
	})
	sc.Step(`^CRC, stored-size and trailing-byte refuse as zip-crc-mismatch, zip-size-mismatch and zip-end-record-missing with every caller archive unchanged$`, func() error {
		return batch2Check(s.checks, map[string]string{"CRC mismatch": "zip-crc-mismatch", "stored size mismatch": "zip-size-mismatch", "undeclared trailing byte": "zip-end-record-missing"})
	})

	sc.Step(`^the left XML is (.*)$`, func(value string) { s.left = []byte(value) })
	sc.Step(`^the right XML is (.*)$`, func(value string) { s.right = []byte(value) })
	sc.Step(`^the conservative XML comparator compares their UTF-8 bytes$`, func() { s.same = losslessxml.Equivalent(s.left, s.right) })
	sc.Step(`^the comparison result is true$`, func() error {
		if !s.same {
			return fmt.Errorf("OPC comparison different")
		}
		return nil
	})

	sc.Step(`^graph fixture (fixture-[a-f0-9]+) is opened through the production package editor$`, func(id string) error { return s.openGraphFixture(id, false) })
	sc.Step(`^graph fixture (fixture-[a-f0-9]+) has custom/data.bin payload 070809 with type application/octet-stream and internal root edge rIdData of type urn:test/data$`, func(id string) error { return s.openGraphFixture(id, true) })
	sc.Step(`^graph fixture (fixture-[a-f0-9]+) is opened twice as independent package snapshots$`, func(id string) error {
		if err := s.openGraphFixture(id, false); err != nil {
			return err
		}
		var err error
		s.baseline, err = packaging.OpenPreserved(s.archive, packaging.Limits{})
		return err
	})
	sc.Step(`^custom/data.bin with hexadecimal payload 070809 and type application/octet-stream is added with internal root relationship rIdData of type urn:test/data$`, func() error { return s.addGraphPart() })
	sc.Step(`^removal of custom/data.bin is attempted without detaching rIdData$`, func() error {
		_, hash, err := s.edited.Part("custom/data.bin")
		if err != nil {
			return err
		}
		plan, failure := s.edited.PlanGraphMutation(packaging.GraphMutation{Deletions: []packaging.PartDeletion{{Name: "custom/data.bin", ExpectedSHA256: hash}}})
		s.failure = failure
		if plan != nil {
			s.result = plan
		}
		return nil
	})
	sc.Step(`^graph editing refuses with code "opc-part-referenced" and no successful edit result$`, func() error {
		if s.result != nil {
			return fmt.Errorf("graph returned partial plan")
		}
		return batch2Refusal(s.failure, "opc-part-referenced")
	})
	sc.Step(`^the whole current archive and every current part remain exactly unchanged after the refusal$`, func() error {
		var b bytes.Buffer
		if err := s.edited.WriteTo(&b); err != nil {
			return err
		}
		if !bytes.Equal(b.Bytes(), s.output) {
			return fmt.Errorf("graph refusal changed current archive")
		}
		current, err := opcZipMembers(b.Bytes())
		if err != nil {
			return err
		}
		return batch2SameParts(s.savedParts, current)
	})
	sc.Step(`^root relationship rIdData is explicitly detached and custom/data.bin is removed$`, func() error {
		_, hash, err := s.edited.Part("custom/data.bin")
		if err != nil {
			return err
		}
		plan, err := s.edited.PlanGraphMutation(packaging.GraphMutation{Removals: []packaging.RelationshipRemoval{{Source: "", ID: "rIdData"}}, Deletions: []packaging.PartDeletion{{Name: "custom/data.bin", ExpectedSHA256: hash}}})
		if err != nil {
			return err
		}
		return s.edited.ApplyGraphPlan(plan)
	})
	sc.Step(`^only the effective content type of word/document.xml in the second snapshot is set to application/vnd.test.document\+xml$`, func() error {
		plan, err := s.edited.PlanGraphMutation(packaging.GraphMutation{ContentTypes: []packaging.ContentTypeChange{{Part: "word/document.xml", ContentType: "application/vnd.test.document+xml"}}})
		if err != nil {
			return err
		}
		return s.edited.ApplyGraphPlan(plan)
	})
	sc.Step(`^saving and reopening returns custom/data.bin with exact payload 070809, effective type application/octet-stream and root edge rIdData resolving to custom/data.bin$`, func() error {
		if err := s.saveGraph(); err != nil {
			return err
		}
		part, _, err := s.reopened.Part("custom/data.bin")
		if err != nil || !bytes.Equal(part, []byte{7, 8, 9}) {
			return fmt.Errorf("new payload: %v", err)
		}
		graph, err := s.reopened.Graph()
		if err != nil {
			return err
		}
		for _, p := range graph.Parts {
			if p.Name == "custom/data.bin" && p.ContentType == "application/octet-stream" {
				for _, e := range graph.Edges {
					if e.Source == "" && e.ID == "rIdData" && e.Type == "urn:test/data" && e.ResolvedPart == p.Name {
						return nil
					}
				}
			}
		}
		return fmt.Errorf("missing effective MIME or root edge")
	})
	sc.Step(`^every original member except \[Content_Types\]\.xml and _rels/\.rels retains its exact payload and no other member is added or removed$`, func() error {
		if len(s.savedParts) != len(s.originalParts)+1 {
			return fmt.Errorf("unexpected added members")
		}
		return batch2SameParts(s.originalParts, s.savedParts, packaging.ContentTypesPath, packaging.PackageRelsPath)
	})
	sc.Step(`^saving and reopening has no custom/data.bin, no /custom/data.bin content-type override and no rIdData root edge$`, func() error {
		if err := s.saveGraph(); err != nil {
			return err
		}
		if _, ok := s.savedParts["custom/data.bin"]; ok {
			return fmt.Errorf("part retained")
		}
		if bytes.Contains(s.savedParts[packaging.ContentTypesPath], []byte("/custom/data.bin")) {
			return fmt.Errorf("override retained")
		}
		graph, err := s.reopened.Graph()
		if err != nil {
			return err
		}
		for _, e := range graph.Edges {
			if e.ID == "rIdData" && e.Source == "" {
				return fmt.Errorf("edge retained")
			}
		}
		return nil
	})
	sc.Step(`^every retained internal relationship target resolves and every original non-registry member retains its exact payload$`, func() error {
		g, err := s.reopened.Graph()
		if err != nil {
			return err
		}
		for _, e := range g.Edges {
			if !e.External && e.ResolvedPart == "" {
				return fmt.Errorf("unresolved edge %s", e.ID)
			}
		}
		return batch2SameParts(s.originalParts, s.savedParts, packaging.ContentTypesPath, packaging.PackageRelsPath)
	})
	sc.Step(`^the production payload-and-content-type diff reports changed members \["\[Content_Types\]\.xml","word/document\.xml"\] and no added or removed members$`, func() error {
		var err error
		s.diff, err = packaging.DiffPreserved(s.baseline, s.edited)
		if err != nil {
			return err
		}
		if !reflect.DeepEqual(s.diff.Changed, []string{packaging.ContentTypesPath, "word/document.xml"}) || len(s.diff.Added) != 0 || len(s.diff.Removed) != 0 {
			return fmt.Errorf("content-type diff %+v", s.diff)
		}
		return nil
	})
	sc.Step(`^word/document\.xml retains identical payload bytes and its original MIME changes to application/vnd\.test\.document\+xml$`, func() error {
		a, _, err := s.baseline.Part("word/document.xml")
		if err != nil {
			return err
		}
		b, _, err := s.edited.Part("word/document.xml")
		if err != nil {
			return err
		}
		if !bytes.Equal(a, b) {
			return fmt.Errorf("type-only changed bytes")
		}
		g, err := s.edited.Graph()
		if err != nil {
			return err
		}
		for _, p := range g.Parts {
			if p.Name == "word/document.xml" && p.ContentType == "application/vnd.test.document+xml" {
				return nil
			}
		}
		return fmt.Errorf("effective MIME unchanged")
	})
	sc.Step(`^all unrelated payloads, the first snapshot and caller input bytes remain unchanged$`, func() error {
		if !bytes.Equal(s.archive, s.initial) {
			return fmt.Errorf("caller changed")
		}
		var b bytes.Buffer
		if err := s.baseline.WriteTo(&b); err != nil {
			return err
		}
		if !bytes.Equal(b.Bytes(), s.initial) {
			return fmt.Errorf("first snapshot changed")
		}
		for name, raw := range s.originalParts {
			if name == packaging.ContentTypesPath {
				continue
			}
			now, _, err := s.edited.Part(name)
			if err != nil || !bytes.Equal(raw, now) {
				return fmt.Errorf("unrelated %s: %v", name, err)
			}
		}
		return nil
	})
}

func batch2Check(got, want map[string]string) error {
	for key, reason := range want {
		if got[key] != reason {
			return fmt.Errorf("%s: got %s want %s", key, got[key], reason)
		}
	}
	return nil
}
func batch2UTF16(text string) []byte {
	out := []byte{0xff, 0xfe}
	for _, u := range utf16.Encode([]rune(text)) {
		out = binary.LittleEndian.AppendUint16(out, u)
	}
	return out
}
