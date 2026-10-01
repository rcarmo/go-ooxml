package acceptance

import (
	"archive/zip"
	"bytes"
	"encoding/binary"
	"errors"
	"fmt"
	"reflect"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func batch2MoreSteps(sc *godog.ScenarioContext, s *batch2State) {
	sc.Step(`^a ZIP_DEFLATED archive contains a\.xml with UTF-8 XML enclosing exactly 10000 spaces between <a> and </a>$`, func() error {
		return s.build([]packaging.ZIP32Entry{{Name: "a.xml", Data: []byte("<a>" + strings.Repeat(" ", 10000) + "</a>")}}, []uint16{zip.Deflate}, "")
	})
	sc.Step(`^the package admission guard checks the archive with only (max_members|max_member_bytes|max_total_bytes|max_ratio) set to (\d+)$`, func(name string, value int) error {
		limits := packaging.ArchiveAdmissionLimits{}
		s.reason = map[string]string{"max_members": "zip-too-many-entries", "max_member_bytes": "zip-entry-too-large", "max_total_bytes": "zip-total-too-large", "max_ratio": "zip-compression-ratio-exceeded"}[name]
		switch name {
		case "max_members":
			limits.MaxMembers = &value
		case "max_member_bytes":
			n := uint64(value)
			limits.MaxMemberBytes = &n
		case "max_total_bytes":
			n := uint64(value)
			limits.MaxTotalBytes = &n
		case "max_ratio":
			n := float64(value)
			limits.MaxRatio = &n
		}
		s.read, s.failure = packaging.AdmitArchive(s.archive, limits)
		return nil
	})
	sc.Step(`^a ZIP_STORED archive contains a\.xml with (UTF-8|UTF-16 with BOM) text (.*)$`, func(encoding, text string) error {
		data := []byte(text)
		if encoding == "UTF-16 with BOM" {
			data = batch2UTF16(text)
		}
		return s.build([]packaging.ZIP32Entry{{Name: "a.xml", Data: data}}, []uint16{zip.Store}, "")
	})
	sc.Step(`^the package admission guard checks the archive with default limits$`, func() { s.read, s.failure = packaging.AdmitArchive(s.archive, packaging.ArchiveAdmissionLimits{}) })
	sc.Step(`^package admission is refused$`, func() error {
		if s.read != nil || !bytes.Equal(s.archive, s.initial) {
			return fmt.Errorf("admission returned parts or changed caller")
		}
		if s.reason != "" {
			return batch2Reason(s.failure, s.reason)
		}
		return batch2Reason(s.failure, "opc-xml-member-invalid")
	})

	sc.Step(`^a raw-DEFLATE ZIP32 budget archive contains a\.bin with 4096 A bytes followed by b\.bin with two b bytes$`, func() error {
		return s.build([]packaging.ZIP32Entry{{Name: "a.bin", Data: bytes.Repeat([]byte{'A'}, 4096)}, {Name: "b.bin", Data: []byte("bb")}}, []uint16{zip.Deflate, zip.Deflate}, "")
	})
	sc.Step(`^a sibling archive retains that declared geometry but replaces the compressed a\.bin body with eight FF bytes$`, func() error {
		sibling := bytes.Clone(s.archive)
		first := bytes.Index(sibling, []byte{'P', 'K', 3, 4})
		if first < 0 {
			return fmt.Errorf("missing local header")
		}
		local := first + 30 + int(binary.LittleEndian.Uint16(sibling[first+26:first+28])) + int(binary.LittleEndian.Uint16(sibling[first+28:first+30]))
		central := bytes.Index(sibling, []byte{'P', 'K', 1, 2})
		if central < 0 {
			return fmt.Errorf("missing central directory")
		}
		size := int(binary.LittleEndian.Uint32(sibling[central+20 : central+24]))
		if size < 8 {
			return fmt.Errorf("unexpected compressed geometry")
		}
		copy(sibling[local:local+8], bytes.Repeat([]byte{255}, 8))
		sibling = append(sibling[:local+8], sibling[local+size:]...)
		delta := size - 8
		central -= delta
		binary.LittleEndian.PutUint32(sibling[central+20:central+24], 8)
		binary.LittleEndian.PutUint32(sibling[first+18:first+22], 8)
		// Update all later local offsets and the EOCD central start/length.
		at := central
		for i := 0; i < 2; i++ {
			if binary.LittleEndian.Uint32(sibling[at:at+4]) != 0x02014b50 {
				return fmt.Errorf("central record")
			}
			offset := binary.LittleEndian.Uint32(sibling[at+42 : at+46])
			if i > 0 {
				binary.LittleEndian.PutUint32(sibling[at+42:at+46], offset-uint32(delta))
			}
			at += 46 + int(binary.LittleEndian.Uint16(sibling[at+28:at+30])) + int(binary.LittleEndian.Uint16(sibling[at+30:at+32])) + int(binary.LittleEndian.Uint16(sibling[at+32:at+34]))
		}
		end := len(sibling) - 22
		binary.LittleEndian.PutUint32(sibling[end+16:end+20], uint32(central))
		s.output = sibling
		return nil
	})
	sc.Step(`^each archive is read separately with one limit maxArchiveBytes length-minus-one, maxEntries 1, maxEntryBytes 8, maxTotalBytes 8 or maxCompressionRatio 2$`, func() error {
		valid, err := packaging.ReadZIP32(s.archive, packaging.ZIP32Limits{})
		if err != nil || len(valid) != 2 || string(valid[0].Data) != strings.Repeat("A", 4096) || string(valid[1].Data) != "bb" {
			return fmt.Errorf("positive ZIP budget vector %v", err)
		}
		for _, tc := range []struct {
			name, want string
			limit      packaging.ZIP32Limits
		}{{"archive", "zip-archive-too-large", packaging.ZIP32Limits{}}, {"entries", "zip-too-many-entries", packaging.ZIP32Limits{MaxEntries: 1}}, {"entry", "zip-entry-too-large", packaging.ZIP32Limits{MaxEntryBytes: 8}}, {"total", "zip-total-too-large", packaging.ZIP32Limits{MaxTotalBytes: 8}}, {"ratio", "zip-compression-ratio-exceeded", packaging.ZIP32Limits{MaxCompressionRatio: 2}}} {
			for _, input := range [][]byte{s.archive, s.output} {
				limits := tc.limit
				if tc.name == "archive" {
					limits.MaxArchiveBytes = int64(len(input) - 1)
				}
				source := bytes.Clone(input)
				result, err := packaging.ReadZIP32(input, limits)
				if result != nil || !bytes.Equal(source, input) {
					return fmt.Errorf("budget %s delivered/changed", tc.name)
				}
				if err = batch2Reason(err, tc.want); err != nil {
					return fmt.Errorf("budget %s: %w", tc.name, err)
				}
			}
			s.checks[tc.name] = tc.want
		}
		return nil
	})
	sc.Step(`^both archives refuse the same respective reasons zip-archive-too-large, zip-too-many-entries, zip-entry-too-large, zip-total-too-large and zip-compression-ratio-exceeded before member output$`, func() error {
		if len(s.checks) != 5 {
			return fmt.Errorf("budget checks %v", s.checks)
		}
		return nil
	})
	sc.Step(`^default reading of the valid archive returns the two exact payloads and default reading of the invalid-DEFLATE sibling refuses a payload error$`, func() error {
		valid, err := packaging.ReadZIP32(s.archive, packaging.ZIP32Limits{})
		if err != nil || len(valid) != 2 {
			return fmt.Errorf("valid budget read: %v", err)
		}
		bad, err := packaging.ReadZIP32(s.output, packaging.ZIP32Limits{})
		if bad != nil || err == nil {
			return fmt.Errorf("invalid compressed body accepted")
		}
		var refusal *packaging.ProfileError
		if !errors.As(err, &refusal) || refusal.Reason != "zip-size-mismatch" {
			return fmt.Errorf("invalid payload reason %v", err)
		}
		return nil
	})
	sc.Step(`^every caller archive remains unchanged$`, func() error {
		if !bytes.Equal(s.archive, s.initial) {
			return fmt.Errorf("valid ZIP changed")
		}
		return nil
	})

	sc.Step(`^ordered ZIP writer members are \[Content_Types\]\.xml with UTF-8 JSON "<Types/>", custom/data\.bin with bytes 00 through FF and word/document\.xml with UTF-8 <w:document> followed by 2048 A characters and </w:document>$`, func() {
		data := make([]byte, 256)
		for i := range data {
			data[i] = byte(i)
		}
		s.entries = []packaging.ZIP32Entry{{Name: packaging.ContentTypesPath, Data: []byte("<Types/>")}, {Name: "custom/data.bin", Data: data}, {Name: "word/document.xml", Data: []byte("<w:document>" + strings.Repeat("A", 2048) + "</w:document>")}}
	})
	sc.Step(`^the production ZIP writer writes the same members twice using its deterministic ZIP32 profile$`, func() error {
		var err error
		s.archive, err = packaging.WriteZIP32Deterministic(s.entries)
		if err != nil {
			return err
		}
		s.output, err = packaging.WriteZIP32Deterministic(s.entries)
		return err
	})
	sc.Step(`^both outputs are byte-identical ZIP32 archives and reopen with the three exact member names and payloads$`, func() error {
		if !bytes.Equal(s.archive, s.output) {
			return fmt.Errorf("writer nondeterministic")
		}
		var err error
		s.read, err = packaging.ReadZIP32(s.archive, packaging.ZIP32Limits{})
		if err != nil {
			return err
		}
		return batch2Entries(s.read, s.entries)
	})
	sc.Step(`^every file uses STORED or DEFLATED encoding and this input emits at least one of each method$`, func() error {
		zr, err := zip.NewReader(bytes.NewReader(s.archive), int64(len(s.archive)))
		if err != nil {
			return err
		}
		methods := map[uint16]bool{}
		for _, file := range zr.File {
			if file.Method != zip.Store && file.Method != zip.Deflate {
				return fmt.Errorf("unsupported method")
			}
			methods[file.Method] = true
		}
		if !methods[zip.Store] || !methods[zip.Deflate] {
			return fmt.Errorf("writer missing method")
		}
		return nil
	})
	sc.Step(`^every local and central member name uses the UTF-8 ZIP flag$`, func() error {
		zr, err := zip.NewReader(bytes.NewReader(s.archive), int64(len(s.archive)))
		if err != nil {
			return err
		}
		central := bytes.Index(s.archive, []byte{'P', 'K', 1, 2})
		if central < 0 {
			return fmt.Errorf("central directory absent")
		}
		at := central
		for i, f := range zr.File {
			if f.Flags&0x800 == 0 || at+46 > len(s.archive) || binary.LittleEndian.Uint32(s.archive[at:at+4]) != 0x02014b50 || binary.LittleEndian.Uint16(s.archive[at+8:at+10])&0x800 == 0 {
				return fmt.Errorf("central UTF-8 flag/geometry %d", i)
			}
			local := int(binary.LittleEndian.Uint32(s.archive[at+42 : at+46]))
			if local+30 > central || binary.LittleEndian.Uint32(s.archive[local:local+4]) != 0x04034b50 || binary.LittleEndian.Uint16(s.archive[local+6:local+8])&0x800 == 0 {
				return fmt.Errorf("local UTF-8 flag/geometry %d", i)
			}
			at += 46 + int(binary.LittleEndian.Uint16(s.archive[at+28:at+30])) + int(binary.LittleEndian.Uint16(s.archive[at+30:at+32])) + int(binary.LittleEndian.Uint16(s.archive[at+32:at+34]))
		}
		return nil
	})
	sc.Step(`^an independent writer input word/document\.xml=one and WORD/document\.xml=two refuses as zip-case-collision with no archive and unchanged caller members$`, func() error {
		entries := []packaging.ZIP32Entry{{Name: "word/document.xml", Data: []byte("one")}, {Name: "WORD/document.xml", Data: []byte("two")}}
		before := []packaging.ZIP32Entry{{Name: entries[0].Name, Data: bytes.Clone(entries[0].Data)}, {Name: entries[1].Name, Data: bytes.Clone(entries[1].Data)}}
		output, err := packaging.WriteZIP32Deterministic(entries)
		if output != nil || !reflect.DeepEqual(entries, before) {
			return fmt.Errorf("writer returned archive/mutated ordered caller members")
		}
		return batch2Reason(err, "zip-case-collision")
	})

	sc.Step(`^the original ZIP_STORED package has these ordered UTF-8 members$`, func(table *godog.Table) error {
		entries, err := batch2TableEntries(table)
		if err != nil {
			return err
		}
		return s.build(entries, nil, "")
	})
	sc.Step(`^the modified ZIP_STORED package has these ordered UTF-8 members$`, func(table *godog.Table) error {
		entries, err := batch2TableEntries(table)
		if err != nil {
			return err
		}
		var b batch2State
		if err = b.build(entries, nil, ""); err != nil {
			return err
		}
		s.output = b.archive
		return nil
	})
	sc.Step(`^the semantic package diff compares original and modified packages$`, func() error {
		left, right := bytes.Clone(s.archive), bytes.Clone(s.output)
		var err error
		s.diff, err = packaging.DiffZIP32(s.archive, s.output, packaging.ZIP32Limits{})
		if err != nil {
			return err
		}
		if !bytes.Equal(s.archive, left) || !bytes.Equal(s.output, right) {
			return fmt.Errorf("semantic diff mutated a caller archive")
		}
		return nil
	})
	for _, tc := range []struct {
		pattern string
		get     func() []string
	}{{`^the equivalent_xml member list is \["a\.xml"\]$`, func() []string { return s.diff.EquivalentXML }}, {`^the changed member list is \["b\.bin"\]$`, func() []string { return s.diff.Changed }}, {`^the added member list is \["c\.bin"\]$`, func() []string { return s.diff.Added }}, {`^the removed member list is \[\]$`, func() []string { return s.diff.Removed }}} {
		tc := tc
		sc.Step(tc.pattern, func() error {
			want := []string{}
			switch {
			case strings.Contains(tc.pattern, "a\\.xml"):
				want = []string{"a.xml"}
			case strings.Contains(tc.pattern, "b\\.bin"):
				want = []string{"b.bin"}
			case strings.Contains(tc.pattern, "c\\.bin"):
				want = []string{"c.bin"}
			}
			if !reflect.DeepEqual(tc.get(), want) {
				return fmt.Errorf("diff got %v want %v", tc.get(), want)
			}
			return nil
		})
	}
}

func batch2TableEntries(table *godog.Table) ([]packaging.ZIP32Entry, error) {
	entries := []packaging.ZIP32Entry{}
	for i, row := range table.Rows {
		if i == 0 {
			continue
		}
		if len(row.Cells) != 2 {
			return nil, fmt.Errorf("ZIP table width")
		}
		entries = append(entries, packaging.ZIP32Entry{Name: row.Cells[0].Value, Data: []byte(row.Cells[1].Value)})
	}
	return entries, nil
}
