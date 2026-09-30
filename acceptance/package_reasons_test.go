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
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"

	gherkin "github.com/cucumber/gherkin/go/v26"
	"github.com/cucumber/godog"
	messages "github.com/cucumber/messages/go/v21"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

const (
	opcReasonOpenID          = "@id-bun-opc-open-refusal"
	opcReasonSaveID          = "@id-bun-opc-save-invalid-target-custody"
	opcReasonSymlinkID       = "@id-bun-opc-symlink-destination-refusal"
	opcDeferredTransactionID = "@id-bun-opc-async-transaction-refusal"
	opcOpaqueTransactionID   = "@id-bun-opc-thenable-transaction-result"
	zipReasonReaderID        = "@id-bun-zip32-reader-refusal"
	zipReasonWriterID        = "@id-bun-zip32-writer-refusal"
	zipReasonBoundsID        = "@id-bun-zip32-configured-bounds"
)

type reasonRow struct{ variant, input, mutation, reason string }

var opcReasonRows = []reasonRow{
	{"escaped part name", "", "rename the document, content-type override and relationship target to word/%66oo.xml", "opc-part-name-invalid"},
	{"escaped target", "", "change only the root relationship target to word/%66oo.xml", "opc-target-invalid"},
	{"duplicate default", "", "replace the xml default with a second rels default of type application/xml", "opc-content-types-invalid"},
	{"missing target", "", "change only the root relationship target to word/missing.xml", "opc-relationship-target-missing"},
	{"duplicate relation ID", "", "duplicate the complete rId1 root relationship", "opc-relationship-duplicate"},
}
var zipReaderRows = []reasonRow{
	{"duplicate names", `[["word/document.xml","one"],["word/document.xml","two"]]`, "none", "zip-duplicate-entry"},
	{"ASCII case collision", `[["word/document.xml","one"],["WORD/document.xml","two"]]`, "none", "zip-case-collision"},
	{"parent traversal", `[["../word/document.xml","bad"]]`, "none", "zip-name-invalid"},
	{"encryption flag", `[["word/document.xml","secret"]]`, "general-purpose bit 0 is set in both headers", "zip-encryption-unsupported"},
	{"unsupported method", `[["word/document.xml","x"]]`, "both methods are 12 and payload bytes are stored uncompressed", "zip-method-unsupported"},
	{"multiple disks", `[["word/document.xml","x"]]`, "end record disk number is 1", "zip-multi-disk-unsupported"},
	{"missing ZIP64 records", `[["word/document.xml","x"]]`, "both end-record counts are 65535 without ZIP64 records", "zip-structure-invalid"},
	{"local name mismatch", `[["word/document.xml","x"]]`, "local name is word/other.xml", "zip-local-metadata-mismatch"},
	{"CRC mismatch", `[["word/document.xml","payload"]]`, "both CRC fields are hexadecimal DEADBEEF", "zip-crc-mismatch"},
	{"stored size mismatch", `[["word/document.xml","payload"]]`, "method is STORED and all size fields are 99", "zip-size-mismatch"},
	{"deflate size overrun", `[["word/document.xml","A"]]`, "payload repeats A 4096 times but both expanded sizes are 32", "zip-size-mismatch"},
	{"undeclared trailing byte", `[["word/document.xml","x"]]`, "a newline byte follows the complete uncommented archive", "zip-end-record-missing"},
}
var zipWriterRows = []reasonRow{
	{"ASCII case collision", `[["word/document.xml","one"],["WORD/document.xml","two"]]`, "", "zip-case-collision"},
	{"non-empty directory entry", `[["word/","not empty"]]`, "", "zip-directory-entry-invalid"},
}
var zipBudgetRows = []reasonRow{
	{"maxArchiveBytes", "archive byte length minus 1", "", "zip-archive-too-large"},
	{"maxEntries", "1", "", "zip-too-many-entries"},
	{"maxEntryBytes", "8", "", "zip-entry-too-large"},
	{"maxTotalBytes", "8", "", "zip-total-too-large"},
	{"maxCompressionRatio", "2", "", "zip-compression-ratio-exceeded"},
}

// This profile is opt-in only; the complete checkout is separately checked by
// loadReferencePin, including HEAD, clean tree, tracked bytes and manifest seal.
func packageSelectedReasonCases() int {
	if portableTransactionCandidate() {
		return 12
	} // three prior + seven reasons + two transactions
	if packageReasonsCandidate() {
		return 10
	} // three prior + five open + two save
	return 3
}

func portableTransactionCandidate() bool {
	if !packageReasonsCandidate() {
		return false
	}
	b, err := os.ReadFile(packagePreservationFeaturePath())
	return err == nil && bytes.Contains(b, []byte("@profile-portable-transactions @id-bun-opc-async-transaction-refusal")) && bytes.Contains(b, []byte("@profile-portable-transactions @id-bun-opc-thenable-transaction-result"))
}
func zipSelectedReasonCases() int {
	if packageReasonsCandidate() {
		return 19
	} // 12 reader + 2 writer + 5 budgets
	return 0
}

func packageReasonsCandidate() bool {
	for _, feature := range []struct{ path, marker string }{
		{packagePreservationFeaturePath(), "@profile-package-refusal-reasons @id-bun-opc-open-refusal"},
		{zip32FeaturePath(), "@profile-zip32-refusal-reasons @id-bun-zip32-reader-refusal"},
	} {
		b, err := os.ReadFile(feature.path)
		if err != nil || !bytes.Contains(b, []byte(feature.marker)) {
			return false
		}
	}
	return true
}

func expectedProfileReason(err error, reason string) error {
	var refusal *packaging.ProfileError
	if !errors.As(err, &refusal) || refusal.Reason != reason {
		return fmt.Errorf("structured reason = %v, want %s", err, reason)
	}
	// A native unrelated failure or another structured reason is not this row.
	if errors.As(errors.New("unrelated I/O"), &refusal) || reason == "" {
		return fmt.Errorf("unrelated failure classified")
	}
	return nil
}

func guardReasonCase(p *messages.Pickle, id string, line int, profile, name string, steps []string) error {
	if p == nil || p.Name != name || len(p.AstNodeIds) == 0 || len(p.Steps) != len(steps) || len(p.Tags) != 3 || p.Tags[0].Name != "@planned" || p.Tags[1].Name != profile || p.Tags[2].Name != id || line <= 0 {
		return fmt.Errorf("reason case %s identity drift at %d: name=%q ast=%d steps=%d want=%d", id, line, p.Name, len(p.AstNodeIds), len(p.Steps), len(steps))
	}
	for i, want := range steps {
		if p.Steps[i].Text != want {
			return fmt.Errorf("reason case %s step %d drift", id, i+1)
		}
	}
	return nil
}

func guardReasonFeature(doc *messages.GherkinDocument, path string) error {
	if doc == nil || doc.Feature == nil {
		return fmt.Errorf("missing reason feature")
	}
	if path == packagePreservationFeaturePath() {
		if doc.Feature.Name != "OPC package custody, transactions and save destinations" {
			return fmt.Errorf("OPC reason feature drift")
		}
	} else if path == zip32FeaturePath() {
		if doc.Feature.Name != "ZIP32 reading, writing and bounded admission" {
			return fmt.Errorf("ZIP32 reason feature drift")
		}
	} else {
		return fmt.Errorf("unexpected reason feature")
	}
	return nil
}

func guardOPCReasonPickle(id string, p *messages.Pickle, line int) error {
	profile := "@profile-package-refusal-reasons"
	if id == opcDeferredTransactionID || id == opcOpaqueTransactionID {
		profile = "@profile-portable-transactions"
	}
	switch id {
	case opcDeferredTransactionID:
		return guardReasonCase(p, id, line, profile, "Deferred transactions refuse before invoking the edit callback", append(append([]string{}, opcDetachedBackground...),
			"the package editor has opened the base archive bytes",
			"a deferred transaction is requested with a callback that would replace Alpha with Beta and set a ran flag",
			"the transaction refuses with reason opc-deferred-transaction and no result",
			"the ran flag is false and the document text still contains Alpha",
			"the current package archive and caller source bytes remain unchanged"))
	case opcOpaqueTransactionID:
		return guardReasonCase(p, id, line, profile, "An immediate transaction returns its opaque token without evaluating it", append(append([]string{}, opcDetachedBackground...),
			"the package editor has opened the base archive bytes",
			"an opaque token has an evaluation hook that throws if invoked",
			"an immediate transaction replaces Alpha with Beta and returns that token",
			"the returned token has the original identity and its evaluation count is zero",
			"saving and reopening reads Beta with every unrelated member payload unchanged"))
	}

	switch id {
	case opcReasonOpenID:
		for _, row := range opcReasonRows {
			if p.Name == "Reject "+row.variant+" with its structured validation reason" {
				return guardReasonCase(p, id, line, profile, p.Name, append(append([]string{}, opcDetachedBackground...),
					"the base package has the mutation "+row.mutation, "the package editor opens its archive bytes",
					"opening refuses with reason "+row.reason+" and no package result", "the caller's original archive bytes remain unchanged"))
			}
		}
	case opcReasonSaveID:
		return guardReasonCase(p, id, line, profile, "A failed validation leaves an existing output file untouched", append(append([]string{}, opcDetachedBackground...),
			"an existing destination file contains the base package archive bytes", "the package editor has opened those bytes and deleted word/document.xml",
			"the package is saved to the existing destination", "saving refuses with reason opc-relationship-target-missing before destination replacement",
			"no successful save receipt is returned", "the destination file bytes equal the original archive bytes"))
	case opcReasonSymlinkID:
		return guardReasonCase(p, id, line, profile, "A symlink destination is refused without changing its target", append(append([]string{}, opcDetachedBackground...),
			"a regular destination file contains the base package archive bytes", "a symlink points to that file", "the package editor has opened the base archive bytes",
			"the unchanged package is saved through the symlink path", "saving refuses with reason opc-symlink-destination before destination replacement",
			"the symlink still points to the original regular destination without a successful save receipt", "the regular destination file bytes equal the original archive bytes"))
	}
	return fmt.Errorf("unrecognized OPC reason case %s/%q", id, p.Name)
}

func guardZIPReasonPickle(id string, p *messages.Pickle, line int) error {
	profile := "@profile-zip32-refusal-reasons"
	switch id {
	case zipReasonReaderID:
		for _, row := range zipReaderRows {
			if p.Name == "Refuse "+row.variant+" with its structured ZIP32 reason" {
				return guardReasonCase(p, id, line, profile, p.Name, []string{"a ZIP32 reader sample with these ordered member and payload pairs encoded as JSON " + row.input, "the sample has the mutation " + row.mutation, "the ZIP reader reads the sample with default limits", "reading refuses with reason " + row.reason + " and no member payload result", "the caller's original archive bytes remain unchanged"})
			}
		}
	case zipReasonWriterID:
		for _, row := range zipWriterRows {
			if p.Name == "Refuse "+row.variant+" before returning an archive" {
				return guardReasonCase(p, id, line, profile, p.Name, []string{"ordered writer entries are encoded as JSON " + row.input, "the ZIP writer writes the entries with default options", "writing refuses with reason " + row.reason + " and no archive result", "the ordered caller entry names and payload bytes remain unchanged"})
			}
		}
	case zipReasonBoundsID:
		for _, row := range zipBudgetRows {
			if p.Name == "Refuse the configured "+row.variant+" threshold" {
				return guardReasonCase(p, id, line, profile, p.Name, []string{"a raw-DEFLATE ZIP32 archive contains a.bin with 4096 A bytes followed by b.bin with two b bytes", "the ZIP reader reads the archive with only " + row.variant + " set to " + row.input, "reading refuses with reason " + row.reason + " and no member payload result", "the caller's original archive bytes remain unchanged"})
			}
		}
	}
	return fmt.Errorf("unrecognized ZIP32 reason case %s/%q", id, p.Name)
}

func TestPackageReasonGuardsAndNegativeControls(t *testing.T) {
	loadReferencePin(t)
	if !packageReasonsCandidate() {
		t.Skip("published reference retains earlier error API profiles")
	}
	for _, path := range []string{packagePreservationFeaturePath(), zip32FeaturePath()} {
		f, err := os.Open(path)
		if err != nil {
			t.Fatal(err)
		}
		n := 0
		next := func() string { n++; return fmt.Sprint(n) }
		doc, err := gherkin.ParseGherkinDocument(f, next)
		_ = f.Close()
		if err != nil {
			t.Fatal(err)
		}
		if err := guardReasonFeature(doc, path); err != nil {
			t.Fatal(err)
		}
		seen := map[string]int{}
		for _, p := range gherkin.Pickles(*doc, path, next) {
			for _, tag := range p.Tags {
				var guard func(string, *messages.Pickle, int) error
				if path == zip32FeaturePath() && (tag.Name == zipReasonReaderID || tag.Name == zipReasonWriterID || tag.Name == zipReasonBoundsID) {
					guard = guardZIPReasonPickle
				}
				if path == packagePreservationFeaturePath() && (tag.Name == opcReasonOpenID || tag.Name == opcReasonSaveID || tag.Name == opcReasonSymlinkID || tag.Name == opcDeferredTransactionID || tag.Name == opcOpaqueTransactionID) {
					guard = guardOPCReasonPickle
				}
				if guard == nil {
					continue
				}
				seen[tag.Name]++
				if err := guard(tag.Name, p, 1); err != nil {
					t.Fatal(err)
				}
				copyP := *p
				copyP.Steps = append([]*messages.PickleStep(nil), p.Steps...)
				copyStep := *copyP.Steps[len(copyP.Steps)-2]
				copyStep.Text += " wrong"
				copyP.Steps[len(copyP.Steps)-2] = &copyStep
				if guard(tag.Name, &copyP, 1) == nil {
					t.Fatal("reason step drift accepted", tag.Name)
				}
			}
		}
		want := map[string]int{opcReasonOpenID: 5, opcReasonSaveID: 1, opcReasonSymlinkID: 1, zipReasonReaderID: 12, zipReasonWriterID: 2, zipReasonBoundsID: 5}
		if portableTransactionCandidate() {
			want[opcDeferredTransactionID] = 1
			want[opcOpaqueTransactionID] = 1
		}
		for id, n := range want {
			if path == zip32FeaturePath() != strings.HasPrefix(id, "@id-bun-zip32-") {
				continue
			}
			if seen[id] != n {
				t.Fatalf("%s cases %d want %d", id, seen[id], n)
			}
		}
	}
	for _, tc := range []struct{ reason, wrong string }{{"zip-crc-mismatch", "zip-size-mismatch"}, {"opc-target-invalid", "opc-relationship-target-missing"}} {
		if expectedProfileReason(&packaging.ProfileError{Reason: tc.wrong}, tc.reason) == nil {
			t.Fatal("wrong structured reason accepted")
		}
	}
}

type zipProfileFixture struct {
	name, localName          string
	payload                  []byte
	method, flags            uint16
	crc                      uint32
	fileSize, compressedSize int
}

func buildReasonZIP(specs []zipProfileFixture) []byte {
	var body, central bytes.Buffer
	put16 := func(b *bytes.Buffer, n uint16) { _ = binary.Write(b, binary.LittleEndian, n) }
	put32 := func(b *bytes.Buffer, n uint32) { _ = binary.Write(b, binary.LittleEndian, n) }
	for _, s := range specs {
		localName := s.localName
		if localName == "" {
			localName = s.name
		}
		method := s.method
		if method == 0 {
			method = 8
		}
		var packed bytes.Buffer
		if method == 8 {
			w, _ := flate.NewWriter(&packed, flate.DefaultCompression)
			_, _ = w.Write(s.payload)
			_ = w.Close()
		} else {
			packed.Write(s.payload)
		}
		csize := packed.Len()
		if s.compressedSize > 0 {
			csize = s.compressedSize
		}
		size := len(s.payload)
		if s.fileSize > 0 {
			size = s.fileSize
		}
		crc := crc32.ChecksumIEEE(s.payload)
		if s.crc != 0 {
			crc = s.crc
		}
		offset := body.Len()
		put32(&body, 0x04034b50)
		put16(&body, 20)
		put16(&body, s.flags|0x800)
		put16(&body, method)
		put16(&body, 0)
		put16(&body, 0)
		put32(&body, crc)
		put32(&body, uint32(csize))
		put32(&body, uint32(size))
		put16(&body, uint16(len(localName)))
		put16(&body, 0)
		body.WriteString(localName)
		body.Write(packed.Bytes())
		put32(&central, 0x02014b50)
		put16(&central, 20)
		put16(&central, 20)
		put16(&central, s.flags|0x800)
		put16(&central, method)
		put16(&central, 0)
		put16(&central, 0)
		put32(&central, crc)
		put32(&central, uint32(csize))
		put32(&central, uint32(size))
		put16(&central, uint16(len(s.name)))
		put16(&central, 0)
		put16(&central, 0)
		put16(&central, 0)
		put16(&central, 0)
		put32(&central, 0)
		put32(&central, uint32(offset))
		central.WriteString(s.name)
	}
	out := append([]byte{}, body.Bytes()...)
	centralAt := len(out)
	out = append(out, central.Bytes()...)
	var end bytes.Buffer
	put32(&end, 0x06054b50)
	put16(&end, 0)
	put16(&end, 0)
	put16(&end, uint16(len(specs)))
	put16(&end, uint16(len(specs)))
	put32(&end, uint32(central.Len()))
	put32(&end, uint32(centralAt))
	put16(&end, 0)
	return append(out, end.Bytes()...)
}

func decodeReasonEntries(raw string) ([]packaging.ZIP32Entry, error) {
	var pairs [][]string
	if err := json.Unmarshal([]byte(raw), &pairs); err != nil {
		return nil, err
	}
	if len(pairs) == 0 {
		return nil, fmt.Errorf("empty ZIP entries")
	}
	entries := make([]packaging.ZIP32Entry, len(pairs))
	for i, pair := range pairs {
		if len(pair) != 2 {
			return nil, fmt.Errorf("invalid ZIP entry %d", i)
		}
		entries[i] = packaging.ZIP32Entry{Name: pair[0], Data: []byte(pair[1])}
	}
	return entries, nil
}
func rawZIPReasonSample(entries []packaging.ZIP32Entry, mutation string) ([]byte, error) {
	specs := make([]zipProfileFixture, len(entries))
	for i, e := range entries {
		specs[i] = zipProfileFixture{name: e.Name, payload: e.Data}
	}
	var apply func([]byte)
	switch mutation {
	case "none":
	case "general-purpose bit 0 is set in both headers":
		specs[0].flags = 1
	case "both methods are 12 and payload bytes are stored uncompressed":
		apply = func(d []byte) {
			binary.LittleEndian.PutUint16(d[8:10], 12)
			at := bytes.Index(d, []byte{'P', 'K', 1, 2})
			binary.LittleEndian.PutUint16(d[at+10:at+12], 12)
		}
	case "end record disk number is 1":
		apply = func(d []byte) { binary.LittleEndian.PutUint16(d[len(d)-18:], 1) }
	case "both end-record counts are 65535 without ZIP64 records":
		apply = func(d []byte) {
			binary.LittleEndian.PutUint16(d[len(d)-14:], 65535)
			binary.LittleEndian.PutUint16(d[len(d)-12:], 65535)
		}
	case "local name is word/other.xml":
		specs[0].localName = "word/other.xml"
	case "both CRC fields are hexadecimal DEADBEEF":
		specs[0].crc = 0xdeadbeef
	case "method is STORED and all size fields are 99":
		apply = func(d []byte) {
			at := bytes.Index(d, []byte{'P', 'K', 1, 2})
			binary.LittleEndian.PutUint16(d[8:10], 0)
			binary.LittleEndian.PutUint16(d[at+10:at+12], 0)
			for _, off := range []int{18, 22, at + 20, at + 24} {
				binary.LittleEndian.PutUint32(d[off:off+4], 99)
			}
		}
	case "payload repeats A 4096 times but both expanded sizes are 32":
		specs[0].payload = []byte(strings.Repeat("A", 4096))
		specs[0].fileSize = 32
	case "a newline byte follows the complete uncommented archive":
		apply = func(d []byte) {}
	default:
		return nil, fmt.Errorf("unknown ZIP sample mutation %q", mutation)
	}
	data := buildReasonZIP(specs)
	if mutation == "a newline byte follows the complete uncommented archive" {
		data = append(data, '\n')
	} else if apply != nil {
		apply(data)
	}
	return data, nil
}

func equalReasonEntries(a, b []packaging.ZIP32Entry) bool {
	if len(a) != len(b) {
		return false
	}
	for i := range a {
		if a[i].Name != b[i].Name || !bytes.Equal(a[i].Data, b[i].Data) {
			return false
		}
	}
	return true
}

// The three-member positive control is independently generated with archive/zip;
// no profile writer or refusal label is used as its acceptance oracle.
func reasonOPCArchive(types, rels, part string) ([]byte, error) {
	var b bytes.Buffer
	z := zip.NewWriter(&b)
	for _, entry := range []struct{ name, text string }{{"[Content_Types].xml", types}, {"_rels/.rels", rels}, {part, opcDetachedMain}} {
		w, err := z.Create(entry.name)
		if err != nil {
			return nil, err
		}
		if _, err = io.WriteString(w, entry.text); err != nil {
			return nil, err
		}
	}
	if err := z.Close(); err != nil {
		return nil, err
	}
	return bytes.Clone(b.Bytes()), nil
}
func reasonOPCContents(mutation string) (string, string, string, error) {
	ct := `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="` + opcDetachedTypesNS + `"><Default Extension="rels" ContentType="` + opcDetachedRelsType + `"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="` + opcDetachedMainType + `"/></Types>`
	rel := `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="` + opcDetachedRelsNS + `"><Relationship Id="rId1" Type="` + opcDetachedOfficeType + `" Target="word/document.xml"/></Relationships>`
	part := "word/document.xml"
	switch mutation {
	case "":
	case opcReasonRows[0].mutation:
		part = "word/%66oo.xml"
		ct = strings.Replace(ct, "/word/document.xml", "/word/%66oo.xml", 1)
		rel = strings.Replace(rel, "word/document.xml", "word/%66oo.xml", 1)
	case opcReasonRows[1].mutation:
		rel = strings.Replace(rel, "word/document.xml", "word/%66oo.xml", 1)
	case opcReasonRows[2].mutation:
		ct = strings.Replace(ct, `Default Extension="xml"`, `Default Extension="rels"`, 1)
	case opcReasonRows[3].mutation:
		rel = strings.Replace(rel, "word/document.xml", "word/missing.xml", 1)
	case opcReasonRows[4].mutation:
		at := strings.Index(rel, "<Relationship ")
		end := at + strings.Index(rel[at:], "/>") + 2
		rel = rel[:end] + rel[at:end] + rel[end:]
	default:
		return "", "", "", fmt.Errorf("unknown OPC mutation %q", mutation)
	}
	return ct, rel, part, nil
}

type packageReasonState struct {
	archive, original                    []byte
	entries, writerBefore, outputEntries []packaging.ZIP32Entry
	outputBytes                          []byte
	packageResult                        *packaging.Envelope
	err                                  error
	mutation                             string
	limit                                packaging.ZIP32Limits
	selectedReason                       string
	destination, link                    string
	tempDir                              string
	saveSucceeded                        bool
	callbackRan                          bool
	transactionResult                    any
	token                                *opaquePackageToken
}

type opaquePackageToken struct{ evaluations int }

func (t *opaquePackageToken) Evaluate() any { t.evaluations++; panic("opaque token was evaluated") }

func packageReasonSteps(sc *godog.ScenarioContext, s *packageReasonState) {
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		*s = packageReasonState{}
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, err error) (context.Context, error) {
		if s.tempDir != "" {
			_ = os.RemoveAll(s.tempDir)
		}
		return ctx, nil
	})
	// Existing OPC Background is validated by opcDetachedByteSteps; these
	// scenario-specific steps create their own independently generated archive.
	sc.Step(`^the base package has the mutation (.+)$`, func(mutation string) error {
		ct, rels, part, err := reasonOPCContents(mutation)
		if err != nil {
			return err
		}
		s.mutation = mutation
		s.archive, err = reasonOPCArchive(ct, rels, part)
		if err != nil {
			return err
		}
		s.original = bytes.Clone(s.archive)
		control, err := reasonOPCArchiveControl()
		if err != nil {
			return err
		}
		_, err = packaging.OpenEnvelope(control)
		return err
	})
	sc.Step(`^the package editor opens its archive bytes$`, func() error {
		if s.original == nil {
			return fmt.Errorf("missing OPC source")
		}
		s.packageResult, s.err = packaging.OpenEnvelope(s.archive)
		return nil
	})
	sc.Step(`^opening refuses with reason (opc-[a-z-]+) and no package result$`, func(reason string) error {
		if s.packageResult != nil {
			return fmt.Errorf("OPC refusal returned partial editor")
		}
		if len(s.archive) == 0 || !bytes.Equal(s.archive, s.original) {
			return fmt.Errorf("OPC caller archive changed")
		}
		s.selectedReason = reason
		return expectedProfileReason(s.err, reason)
	})
	sc.Step(`^(?:an existing|a regular) destination file contains the base package archive bytes$`, func() error {
		control, err := reasonOPCArchiveControl()
		if err != nil {
			return err
		}
		s.archive, s.original = control, bytes.Clone(control)
		var dirErr error
		s.tempDir, dirErr = os.MkdirTemp("", "go-profile-*")
		if dirErr != nil {
			return dirErr
		}
		s.destination = filepath.Join(s.tempDir, "target.docx")
		return os.WriteFile(s.destination, s.archive, 0600)
	})
	sc.Step(`^the package editor has opened those bytes and deleted word/document\.xml$`, func() error {
		var err error
		s.packageResult, err = packaging.OpenEnvelope(s.archive)
		if err != nil {
			return err
		}
		return s.packageResult.DeletePart("word/document.xml")
	})
	sc.Step(`^a symlink points to that file$`, func() error { s.link = s.destination + ".link"; return os.Symlink(s.destination, s.link) })
	sc.Step(`^the package editor has opened the base archive bytes$`, func() error {
		if s.archive == nil {
			var err error
			s.archive, err = reasonOPCArchiveControl()
			if err != nil {
				return err
			}
			s.original = bytes.Clone(s.archive)
		}
		var err error
		s.packageResult, err = packaging.OpenEnvelope(s.archive)
		return err
	})
	sc.Step(`^a deferred transaction is requested with a callback that would replace Alpha with Beta and set a ran flag$`, func() error {
		if s.packageResult == nil {
			return fmt.Errorf("missing OPC envelope")
		}
		s.callbackRan = false
		s.transactionResult, s.err = s.packageResult.Transaction(packaging.TransactionDeferred, func(work *packaging.Envelope) (any, error) {
			s.callbackRan = true
			part, err := work.Part("word/document.xml")
			if err != nil {
				return nil, err
			}
			return nil, work.SetPart("word/document.xml", []byte(strings.Replace(string(part), "Alpha", "Beta", 1)))
		})
		return nil
	})
	sc.Step(`^the transaction refuses with reason opc-deferred-transaction and no result$`, func() error {
		if s.transactionResult != nil {
			return fmt.Errorf("deferred callback returned a result")
		}
		return expectedProfileReason(s.err, "opc-deferred-transaction")
	})
	sc.Step(`^the ran flag is false and the document text still contains Alpha$`, func() error {
		if s.callbackRan {
			return fmt.Errorf("deferred callback ran")
		}
		part, err := s.packageResult.Part("word/document.xml")
		if err != nil || !bytes.Equal(part, []byte(opcDetachedMain)) {
			return fmt.Errorf("deferred changed document: %v", err)
		}
		return nil
	})
	sc.Step(`^the current package archive and caller source bytes remain unchanged$`, func() error {
		if !bytes.Equal(s.archive, s.original) {
			return fmt.Errorf("caller archive changed")
		}
		current, err := s.packageResult.Bytes()
		if err != nil || !bytes.Equal(current, s.original) {
			return fmt.Errorf("current package archive changed: %v", err)
		}
		return nil
	})
	sc.Step(`^an opaque token has an evaluation hook that throws if invoked$`, func() error { s.token = &opaquePackageToken{}; return nil })
	sc.Step(`^an immediate transaction replaces Alpha with Beta and returns that token$`, func() error {
		if s.packageResult == nil || s.token == nil {
			return fmt.Errorf("missing OPC editor/token")
		}
		s.transactionResult, s.err = s.packageResult.Transaction(packaging.TransactionImmediate, func(work *packaging.Envelope) (any, error) {
			part, err := work.Part("word/document.xml")
			if err != nil {
				return nil, err
			}
			if !bytes.Equal(part, []byte(opcDetachedMain)) {
				return nil, fmt.Errorf("original document text drift")
			}
			changed := []byte(strings.Replace(string(part), "Alpha", "Beta", 1))
			if err = work.SetPart("word/document.xml", changed); err != nil {
				return nil, err
			}
			return s.token, nil
		})
		return s.err
	})
	sc.Step(`^the returned token has the original identity and its evaluation count is zero$`, func() error {
		if s.err != nil || s.token == nil || s.transactionResult != s.token || s.token.evaluations != 0 {
			return fmt.Errorf("opaque token identity/evaluation drift: %v", s.err)
		}
		return nil
	})
	sc.Step(`^saving and reopening reads Beta with every unrelated member payload unchanged$`, func() error {
		if s.packageResult == nil {
			return fmt.Errorf("missing OPC editor")
		}
		dir, err := os.MkdirTemp("", "go-profile-transaction-*")
		if err != nil {
			return err
		}
		defer os.RemoveAll(dir)
		dest := filepath.Join(dir, "edited.docx")
		if err = s.packageResult.SaveAs(dest); err != nil {
			return err
		}
		archive, err := os.ReadFile(dest)
		if err != nil {
			return err
		}
		reopened, err := packaging.OpenEnvelope(archive)
		if err != nil {
			return err
		}
		before, err := packaging.OpenEnvelope(s.original)
		if err != nil {
			return err
		}
		originalEntries, err := packaging.ReadZIP32(s.original, packaging.ZIP32Limits{})
		if err != nil {
			return err
		}
		newEntries, err := packaging.ReadZIP32(archive, packaging.ZIP32Limits{})
		if err != nil || len(newEntries) != len(originalEntries) {
			return fmt.Errorf("package inventory changed: %v", err)
		}
		for i, entry := range originalEntries {
			if newEntries[i].Name != entry.Name {
				return fmt.Errorf("member order/name changed")
			}
			if entry.Name == "word/document.xml" {
				if !bytes.Equal(newEntries[i].Data, []byte(strings.Replace(opcDetachedMain, "Alpha", "Beta", 1))) {
					return fmt.Errorf("saved Beta missing")
				}
				continue
			}
			b, e := before.Part(entry.Name)
			if e != nil || !bytes.Equal(b, newEntries[i].Data) {
				return fmt.Errorf("unrelated member changed: %s %v", entry.Name, e)
			}
		}
		part, err := reopened.Part("word/document.xml")
		if err != nil || !bytes.Equal(part, []byte(strings.Replace(opcDetachedMain, "Alpha", "Beta", 1))) || !bytes.Equal(s.archive, s.original) || s.token.evaluations != 0 {
			return fmt.Errorf("reopened text/source/token custody: %v", err)
		}
		return nil
	})
	sc.Step(`^(?:the package is saved to the existing destination|the unchanged package is saved through the symlink path)$`, func() error {
		if s.packageResult == nil {
			return fmt.Errorf("missing OPC editor")
		}
		target := s.destination
		if s.link != "" {
			target = s.link
		}
		s.err = s.packageResult.SaveAs(target)
		s.saveSucceeded = s.err == nil
		return nil
	})
	sc.Step(`^saving refuses with reason (opc-[a-z-]+) before destination replacement$`, func(reason string) error {
		if s.saveSucceeded {
			return fmt.Errorf("save incorrectly succeeded")
		}
		s.selectedReason = reason
		return expectedProfileReason(s.err, reason)
	})
	sc.Step(`^no successful save receipt is returned$`, func() error {
		if s.saveSucceeded || s.err == nil {
			return fmt.Errorf("unexpected successful save receipt")
		}
		return nil
	})
	sc.Step(`^the symlink still points to the original regular destination without a successful save receipt$`, func() error {
		if s.saveSucceeded || s.err == nil || s.link == "" {
			return fmt.Errorf("symlink save succeeded")
		}
		info, err := os.Lstat(s.link)
		if err != nil || info.Mode()&os.ModeSymlink == 0 {
			return fmt.Errorf("symlink replaced: %v", err)
		}
		to, err := os.Readlink(s.link)
		if err != nil || to != s.destination {
			return fmt.Errorf("symlink target changed: %v", err)
		}
		return nil
	})
	sc.Step(`^the (?:regular )?destination file bytes equal the original archive bytes$`, func() error {
		b, err := os.ReadFile(s.destination)
		if err != nil || !bytes.Equal(b, s.original) {
			return fmt.Errorf("OPC destination changed: %v", err)
		}
		control, err := packaging.OpenEnvelope(b)
		if err != nil || control == nil {
			return fmt.Errorf("OPC destination no longer opens: %v", err)
		}
		return nil
	})
	sc.Step(`^a ZIP32 reader sample with these ordered member and payload pairs encoded as JSON (.+)$`, func(raw string) error {
		var err error
		s.entries, err = decodeReasonEntries(raw)
		return err
	})
	sc.Step(`^the sample has the mutation (.+)$`, func(mutation string) error {
		if len(s.entries) == 0 {
			return fmt.Errorf("missing ordered source entries")
		}
		var err error
		s.archive, err = rawZIPReasonSample(s.entries, mutation)
		if err != nil {
			return err
		}
		s.original = bytes.Clone(s.archive)
		good := buildReasonZIP([]zipProfileFixture{{name: "word/document.xml", payload: []byte("valid")}})
		entries, err := packaging.ReadZIP32(good, packaging.ZIP32Limits{})
		if err != nil || len(entries) != 1 || entries[0].Name != "word/document.xml" || string(entries[0].Data) != "valid" {
			return fmt.Errorf("raw ZIP positive control: %v", err)
		}
		return nil
	})
	sc.Step(`^the ZIP reader reads the sample with default limits$`, func() error {
		s.outputEntries, s.err = packaging.ReadZIP32(s.archive, packaging.ZIP32Limits{})
		return nil
	})
	sc.Step(`^reading refuses with reason (zip-[a-z-]+) and no member payload result$`, func(reason string) error {
		if s.outputEntries != nil {
			return fmt.Errorf("ZIP refusal returned member payloads")
		}
		if len(s.archive) == 0 || !bytes.Equal(s.archive, s.original) {
			return fmt.Errorf("ZIP caller archive changed")
		}
		s.selectedReason = reason
		return expectedProfileReason(s.err, reason)
	})
	sc.Step(`^ordered writer entries are encoded as JSON (.+)$`, func(raw string) error {
		var err error
		s.entries, err = decodeReasonEntries(raw)
		if err != nil {
			return err
		}
		s.writerBefore = make([]packaging.ZIP32Entry, len(s.entries))
		for i, e := range s.entries {
			s.writerBefore[i] = packaging.ZIP32Entry{Name: e.Name, Data: bytes.Clone(e.Data)}
		}
		good, err := packaging.WriteZIP32([]packaging.ZIP32Entry{{Name: "a.bin", Data: []byte("a")}, {Name: "b.bin", Data: []byte("bb")}})
		if err != nil {
			return err
		}
		check, err := packaging.ReadZIP32(good, packaging.ZIP32Limits{})
		if err != nil || len(check) != 2 || check[0].Name != "a.bin" || string(check[0].Data) != "a" || check[1].Name != "b.bin" || string(check[1].Data) != "bb" {
			return fmt.Errorf("ZIP writer positive control: %v", err)
		}
		return nil
	})
	sc.Step(`^the ZIP writer writes the entries with default options$`, func() error { s.outputBytes, s.err = packaging.WriteZIP32(s.entries); return nil })
	sc.Step(`^writing refuses with reason (zip-[a-z-]+) and no archive result$`, func(reason string) error {
		if s.outputBytes != nil {
			return fmt.Errorf("ZIP refusal returned partial archive")
		}
		s.selectedReason = reason
		return expectedProfileReason(s.err, reason)
	})
	sc.Step(`^the ordered caller entry names and payload bytes remain unchanged$`, func() error {
		if !equalReasonEntries(s.entries, s.writerBefore) {
			return fmt.Errorf("ZIP writer modified caller entries")
		}
		return nil
	})
	sc.Step(`^a raw-DEFLATE ZIP32 archive contains a\.bin with 4096 A bytes followed by b\.bin with two b bytes$`, func() error {
		s.archive = buildReasonZIP([]zipProfileFixture{{name: "a.bin", payload: []byte(strings.Repeat("A", 4096))}, {name: "b.bin", payload: []byte("bb")}})
		s.original = bytes.Clone(s.archive)
		entries, err := packaging.ReadZIP32(s.archive, packaging.ZIP32Limits{})
		if err != nil || len(entries) != 2 || entries[0].Name != "a.bin" || string(entries[0].Data) != strings.Repeat("A", 4096) || entries[1].Name != "b.bin" || string(entries[1].Data) != "bb" {
			return fmt.Errorf("ZIP budget positive control: %v", err)
		}
		return nil
	})
	sc.Step(`^the ZIP reader reads the archive with only (maxArchiveBytes|maxEntries|maxEntryBytes|maxTotalBytes|maxCompressionRatio) set to (.+)$`, func(key, value string) error {
		if s.archive == nil {
			return fmt.Errorf("missing ZIP budget archive")
		}
		n := 0
		if value == "archive byte length minus 1" {
			n = len(s.archive) - 1
		} else {
			var err error
			n, err = strconv.Atoi(value)
			if err != nil {
				return err
			}
		}
		if n <= 0 {
			return fmt.Errorf("invalid positive ZIP budget")
		}
		var limits packaging.ZIP32Limits
		switch key {
		case "maxArchiveBytes":
			limits.MaxArchiveBytes = int64(n)
		case "maxEntries":
			limits.MaxEntries = n
		case "maxEntryBytes":
			limits.MaxEntryBytes = uint64(n)
		case "maxTotalBytes":
			limits.MaxTotalBytes = uint64(n)
		case "maxCompressionRatio":
			limits.MaxCompressionRatio = float64(n)
		}
		s.outputEntries, s.err = packaging.ReadZIP32(s.archive, limits)
		return nil
	})
}

func reasonOPCArchiveControl() ([]byte, error) {
	ct, rel, part, err := reasonOPCContents("")
	if err != nil {
		return nil, err
	}
	return reasonOPCArchive(ct, rel, part)
}
