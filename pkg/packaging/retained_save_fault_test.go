package packaging

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	"github.com/rcarmo/go-ooxml/internal/testutil"
)

// Faults are injected into the actual Preserved.SaveAs serialization and
// verification inputs after a valid retained property edit is staged. No
// delivery mock or parallel global hook is involved.
func TestRetainedStyleSaveAsFaults(t *testing.T) {
	cases := []struct{ name, path, part, old, new string }{
		{"pptx", "fixtures/pptx/shape-style/shape-style-b4e7fd03880f.pptx", "ppt/slides/slide2.xml", `<a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:ln>`, `<a:solidFill><a:srgbClr val="A1B2C3"/></a:solidFill><a:ln>`},
		{"docx", "fixtures/docx/direct-formatting/direct-formatting-2d32cedb722e.docx", "word/document.xml", `w:val="112233"`, `w:val="A1B2C3"`},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			source, e := os.ReadFile(testutil.ReferencePath(filepath.FromSlash(c.path)))
			// The default pinned distribution predates this sealed candidate.
			// An explicit root or pin is a candidate run: never silently skip
			// a missing control fixture in that mode.
			if os.IsNotExist(e) && os.Getenv("OOXML_FIXTURES_ROOT") == "" && os.Getenv("OOXML_REFERENCE_PIN") == "" {
				t.Skip("retained-style candidate fixture absent from default pin")
			}
			if e != nil {
				t.Fatal(e)
			}
			callerCopy := bytes.Clone(source)
			pkg, e := OpenPreserved(source, Limits{})
			if e != nil {
				t.Fatal(e)
			}
			body, hash, e := pkg.Part(c.part)
			if e != nil {
				t.Fatal(e)
			}
			if bytes.Count(body, []byte(c.old)) != 1 {
				t.Fatal("fixture edit not unique")
			}
			replacement := bytes.Replace(body, []byte(c.old), []byte(c.new), 1)
			if e = pkg.Replace([]Replacement{{Part: c.part, ExpectedSHA256: hash, Data: replacement}}); e != nil {
				t.Fatal(e)
			}
			good := bytes.Clone(pkg.changed[c.part])
			sourceCopy := bytes.Clone(pkg.source)
			beforePart, heldHash, e := pkg.Part(c.part)
			if e != nil {
				t.Fatal(e)
			}
			for _, phase := range []string{"serialization", "verification"} {
				for _, existing := range []bool{true, false} {
					name := phase + "-absent"
					if existing {
						name = phase + "-existing"
					}
					t.Run(name, func(t *testing.T) {
						dir := t.TempDir()
						dest := filepath.Join(dir, "output."+c.name)
						oldDest := []byte("original destination")
						if existing {
							if e = os.WriteFile(dest, oldDest, 0640); e != nil {
								t.Fatal(e)
							}
						}
						switch phase {
						case "serialization":
							pkg.source = []byte("corrupt ZIP source")
						case "verification":
							// Swap only the serialized ZIP source with another valid
							// admitted archive. The verified SaveAs must detect its
							// mismatched inventory/payload before rename.
							other := "fixtures/docx/direct-formatting/direct-formatting-2d32cedb722e.docx"
							if c.name == "docx" {
								other = "fixtures/pptx/shape-style/shape-style-b4e7fd03880f.pptx"
							}
							pkg.source, e = os.ReadFile(testutil.ReferencePath(filepath.FromSlash(other)))
							if e != nil {
								t.Fatal(e)
							}
						}
						_, e := pkg.SaveAs(dest)
						pkg.source = bytes.Clone(sourceCopy)
						pkg.changed[c.part] = bytes.Clone(good)
						if !bytes.Equal(source, callerCopy) {
							t.Fatal("actual caller input changed")
						}
						if e == nil {
							t.Fatal("fault did not reject actual SaveAs")
						}
						saved, readErr := os.ReadFile(dest)
						if existing {
							if readErr != nil || !bytes.Equal(saved, oldDest) {
								t.Fatalf("existing destination changed: %v", readErr)
							}
						} else if !os.IsNotExist(readErr) {
							t.Fatalf("absent destination created: %v", readErr)
						}
						entries, er := os.ReadDir(dir)
						if er != nil {
							t.Fatal(er)
						}
						want := 0
						if existing {
							want = 1
						}
						if len(entries) != want {
							t.Fatalf("staged files remain: %v", entries)
						}
						held, afterHash, er := pkg.Part(c.part)
						if er != nil || afterHash != heldHash || !bytes.Equal(held, beforePart) {
							t.Fatal("fault consumed staged edit or held fingerprint")
						}
						goodDest := filepath.Join(t.TempDir(), "recovered."+c.name)
						if _, er = pkg.SaveAs(goodDest); er != nil {
							t.Fatalf("session unusable after fault: %v", er)
						}
						recovered, er := os.ReadFile(goodDest)
						if er != nil {
							t.Fatal(er)
						}
						opened, er := OpenPreserved(recovered, Limits{})
						if er != nil {
							t.Fatal(er)
						}
						got, _, er := opened.Part(c.part)
						if er != nil || !bytes.Equal(got, beforePart) {
							t.Fatalf("recovered edit lost: %v", er)
						}
					})
				}
			}
		})
	}
}
