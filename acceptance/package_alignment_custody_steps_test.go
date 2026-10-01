package acceptance

import (
	"bytes"
	"encoding/binary"
	"encoding/xml"
	"errors"
	"fmt"
	"os"
	"strings"
	"unicode/utf16"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/losslessxml"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func batch2CustodySteps(sc *godog.ScenarioContext, s *batch2State) {
	sc.Step(`^the custody envelope contains main XML part doc/main\.xml with decoded text Original and opaque custom/data\.bin payload hexadecimal 00FF01FE02FD$`, func() error { return s.fixedEnvelope() })
	sc.Step(`^an immediate production transaction sets main XML text to rolled back, sets custom/data\.bin to hexadecimal 09080706 and then raises callback-failure after both writes$`, func() error {
		marker := errors.New("callback-failure")
		reached := 0
		s.result, s.failure = s.envelope.Transaction(packaging.TransactionImmediate, func(work *packaging.Envelope) (any, error) {
			if err := work.ReplaceXMLText("doc/main.xml", "Original", "rolled back"); err != nil {
				return nil, err
			}
			reached++
			if err := work.SetPart("custom/data.bin", []byte{9, 8, 7, 6}); err != nil {
				return nil, err
			}
			reached++
			a, err := work.Part("doc/main.xml")
			if err != nil {
				return nil, err
			}
			b, err := work.Part("custom/data.bin")
			if err != nil {
				return nil, err
			}
			if !bytes.Contains(a, []byte("rolled back")) || !bytes.Equal(b, []byte{9, 8, 7, 6}) {
				return nil, fmt.Errorf("callback writes not reached")
			}
			return nil, marker
		})
		if reached != 2 || s.failure != marker {
			return fmt.Errorf("callback did not fail after both writes: reached=%d err=%v", reached, s.failure)
		}
		return nil
	})
	sc.Step(`^callback-failure propagates with no transaction result and the whole current archive equals its original bytes$`, func() error {
		if s.result != nil || s.failure == nil || s.failure.Error() != "callback-failure" {
			return fmt.Errorf("callback error/result %v %v", s.result, s.failure)
		}
		saved, err := s.envelope.Bytes()
		if err != nil {
			return err
		}
		if !bytes.Equal(saved, s.initial) || !bytes.Equal(s.archive, s.initial) {
			return fmt.Errorf("transaction changed original archive")
		}
		return nil
	})
	sc.Step(`^doc/main\.xml still decodes to Original, custom/data\.bin is exactly 00FF01FE02FD and every original member remains present$`, func() error {
		main, err := s.envelope.Part("doc/main.xml")
		if err != nil {
			return err
		}
		value, err := batch2Value(main)
		if err != nil || value != "Original" {
			return fmt.Errorf("original XML value %q %v", value, err)
		}
		binaryPart, err := s.envelope.Part("custom/data.bin")
		if err != nil || !bytes.Equal(binaryPart, []byte{0, 255, 1, 254, 2, 253}) {
			return fmt.Errorf("opaque rollback: %v", err)
		}
		now, err := s.envelope.Bytes()
		if err != nil {
			return err
		}
		parts, err := opcZipMembers(now)
		if err != nil {
			return err
		}
		return batch2SameParts(s.originalParts, parts)
	})
	sc.Step(`^a production text edit sets doc/main\.xml value text to JSON "Updated <value>" and saves then reopens the package$`, func() error {
		if err := s.envelope.ReplaceXMLText("doc/main.xml", "Original", "Updated <value>"); err != nil {
			return err
		}
		if err := s.envelope.SaveAs(s.temp + "/output.docx"); err != nil {
			return err
		}
		output, err := os.ReadFile(s.temp + "/output.docx")
		if err != nil {
			return err
		}
		s.output = output
		s.savedParts, err = opcZipMembers(output)
		if err != nil {
			return err
		}
		s.envelope, err = packaging.OpenEnvelope(output)
		return err
	})
	sc.Step(`^doc/main\.xml decodes to JSON "Updated <value>" and custom/data\.bin is exactly 00FF01FE02FD$`, func() error {
		main, err := s.envelope.Part("doc/main.xml")
		if err != nil {
			return err
		}
		value, err := batch2Value(main)
		if err != nil || value != "Updated <value>" {
			return fmt.Errorf("reopened value %q: %v", value, err)
		}
		opaque, err := s.envelope.Part("custom/data.bin")
		if err != nil || !bytes.Equal(opaque, []byte{0, 255, 1, 254, 2, 253}) {
			return fmt.Errorf("reopened opaque %v", err)
		}
		return nil
	})
	sc.Step(`^the member-name set is unchanged, every other member payload is unchanged and the caller's original archive bytes remain unchanged$`, func() error {
		if len(s.originalParts) != len(s.savedParts) || !bytes.Equal(s.archive, s.initial) {
			return fmt.Errorf("member/caller custody")
		}
		return batch2SameParts(s.originalParts, s.savedParts, "doc/main.xml")
	})

	sc.Step(`^a ZIP contains word/document\.xml with UTF-8 XML text (.*)$`, func(text string) error {
		if text != opcDetachedMain {
			return fmt.Errorf("base XML drift")
		}
		var err error
		s.archive, err = opcDetachedArchive()
		if err != nil {
			return err
		}
		s.initial = bytes.Clone(s.archive)
		s.originalParts, err = opcDetachedMembers(s.archive)
		return err
	})
	sc.Step(`^its content types use namespace (.*)$`, func(uri string) error {
		if uri != opcDetachedTypesNS {
			return fmt.Errorf("content-type namespace")
		}
		return opcDetachedAssertTypes(s.originalParts[packaging.ContentTypesPath])
	})
	sc.Step(`^its content types have these defaults and overrides$`, func(table *godog.Table) error {
		if len(table.Rows) != len(opcDetachedTypeRows) {
			return fmt.Errorf("type table size")
		}
		for i, row := range table.Rows {
			for j, cell := range row.Cells {
				if cell.Value != opcDetachedTypeRows[i][j] {
					return fmt.Errorf("type table drift")
				}
			}
		}
		return nil
	})
	sc.Step(`^_rels/\.rels uses namespace (.*)$`, func(uri string) error {
		if uri != packaging.NSRelationships {
			return fmt.Errorf("relationship namespace")
		}
		return nil
	})
	sc.Step(`^its root relationship is rId1 of type (.*) targeting word/document\.xml$`, func(typ string) error {
		if typ != packaging.RelTypeOfficeDocument {
			return fmt.Errorf("root relation type")
		}
		return nil
	})
	sc.Step(`^the package editor opens the base archive bytes$`, func() error {
		var err error
		s.edited, err = packaging.OpenPreserved(s.archive, packaging.Limits{})
		return err
	})
	sc.Step(`^every byte in the caller's original archive array is overwritten with zero$`, func() error {
		if s.edited == nil {
			return fmt.Errorf("not opened")
		}
		clear(s.archive)
		return nil
	})
	sc.Step(`^every byte in the array returned by get for word/document\.xml is overwritten with zero$`, func() error {
		b, _, err := s.edited.Part("word/document.xml")
		if err != nil {
			return err
		}
		clear(b)
		return nil
	})
	sc.Step(`^a fresh get of word/document\.xml contains the UTF-8 text Alpha$`, func() error {
		b, _, err := s.edited.Part("word/document.xml")
		if err != nil || !bytes.Equal(b, s.originalParts["word/document.xml"]) {
			return fmt.Errorf("detached member %v", err)
		}
		return nil
	})
	sc.Step(`^serializing the package returns the exact original archive bytes$`, func() error {
		var b bytes.Buffer
		if err := s.edited.WriteTo(&b); err != nil {
			return err
		}
		if !bytes.Equal(b.Bytes(), s.initial) {
			return fmt.Errorf("detached no-op differs")
		}
		return nil
	})
	sc.Step(`^word/document\.xml instead has a little-endian BOM and UTF-16LE text (.*)$`, func(text string) error {
		if text != `<?xml version="1.0" encoding="UTF-16"?><document>Alpha</document>` {
			return fmt.Errorf("UTF-16 declaration drift")
		}
		wide := batch2UTF16(text)
		entries := []packaging.ZIP32Entry{{Name: packaging.ContentTypesPath, Data: s.originalParts[packaging.ContentTypesPath]}, {Name: packaging.PackageRelsPath, Data: s.originalParts[packaging.PackageRelsPath]}, {Name: "word/document.xml", Data: wide}}
		if err := s.build(entries, nil, ""); err != nil {
			return err
		}
		s.originalParts["word/document.xml"] = bytes.Clone(wide)
		return nil
	})
	sc.Step(`^the package editor opens the package and sets word/document\.xml to its decoded text with Alpha replaced by Beta$`, func() error {
		var err error
		s.envelope, err = packaging.OpenEnvelope(s.archive)
		if err != nil {
			return err
		}
		return s.envelope.ReplaceXMLText("word/document.xml", "Alpha", "Beta")
	})
	sc.Step(`^its serialized archive is read through the ZIP layer$`, func() error {
		var err error
		s.output, err = s.envelope.Bytes()
		if err != nil {
			return err
		}
		s.read, err = packaging.ReadZIP32(s.output, packaging.ZIP32Limits{})
		if err != nil {
			return err
		}
		s.savedParts, err = opcZipMembers(s.output)
		if err != nil {
			return err
		}
		s.envelope, err = packaging.OpenEnvelope(s.output)
		if err != nil {
			return err
		}
		if !bytes.Equal(s.archive, s.initial) || len(s.originalParts) != len(s.savedParts) {
			return fmt.Errorf("UTF-16 saved member inventory/caller custody")
		}
		return batch2SameParts(s.originalParts, s.savedParts, "word/document.xml")
	})
	sc.Step(`^the saved document member starts with hexadecimal bytes FF FE$`, func() error {
		for _, part := range s.read {
			if part.Name == "word/document.xml" {
				if bytes.HasPrefix(part.Data, []byte{0xff, 0xfe}) {
					return nil
				}
				return fmt.Errorf("UTF-16 BOM lost")
			}
		}
		return fmt.Errorf("saved document missing")
	})
	sc.Step(`^decoding the member as UTF-16LE contains Beta and encoding="UTF-16"$`, func() error {
		part, err := s.envelope.Part("word/document.xml")
		if err != nil || !bytes.Equal(part, s.savedParts["word/document.xml"]) {
			return fmt.Errorf("UTF-16 reopened member mismatch: %v", err)
		}
		for _, part := range s.read {
			if part.Name != "word/document.xml" {
				continue
			}
			if len(part.Data)%2 != 0 {
				return fmt.Errorf("odd UTF-16")
			}
			units := make([]uint16, (len(part.Data)-2)/2)
			for i := range units {
				units[i] = binary.LittleEndian.Uint16(part.Data[2+2*i : 4+2*i])
			}
			text := string(utf16.Decode(units))
			if !strings.Contains(text, "Beta") || !strings.Contains(text, `encoding="UTF-16"`) {
				return fmt.Errorf("UTF-16 XML %q", text)
			}
			return nil
		}
		return fmt.Errorf("document absent")
	})

	sc.Step(`^the go-ooxml and python-office-mcp-server fixture corpora are enumerated$`, func() error {
		var err error
		s.checks = map[string]string{}
		s.corpusGo, s.corpusPy, err = readCorpusEntries()
		return err
	})
	sc.Step(`^each OOXML fixture package is opened and serialized without edits through the OPC layer$`, func() error {
		if len(s.corpusGo) != 36 || len(s.corpusPy) != 35 {
			return fmt.Errorf("corpus inventory mismatch")
		}
		var err error
		s.goCount, err = visitCorpus(s.corpusGo, loadCorpusFixture, func(b []byte) error { return corpusRoundTrip(b, realCorpusOpen, realCorpusEmit) })
		if err != nil {
			return err
		}
		s.pyCount, err = visitCorpus(s.corpusPy, loadCorpusFixture, func(b []byte) error { return corpusRoundTrip(b, realCorpusOpen, realCorpusEmit) })
		return err
	})
	sc.Step(`^every reopened package matches its original whole-archive bytes$`, func() error {
		if s.goCount != len(s.corpusGo) || s.pyCount != len(s.corpusPy) {
			return fmt.Errorf("corpus reopen mismatch")
		}
		return nil
	})
	sc.Step(`^both fixture corpora contribute their exact known nonzero fixture counts$`, func() error { return requireCorpusCounts(s.goCount, s.pyCount) })
}

func batch2Value(data []byte) (string, error) {
	doc, err := losslessxml.Parse(data)
	if err != nil {
		return "", err
	}
	var value string
	count := 0
	for _, e := range doc.Elements() {
		if e.Name() == (xml.Name{Space: "urn:acceptance", Local: "value"}) {
			text, ok := e.Text()
			if !ok {
				return "", fmt.Errorf("mixed value")
			}
			value = text
			count++
		}
	}
	if count != 1 {
		return "", fmt.Errorf("value count %d", count)
	}
	return value, nil
}
