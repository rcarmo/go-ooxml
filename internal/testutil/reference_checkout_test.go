package testutil

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"testing"
)

// Root-module batches also enforce reference custody; acceptance is a separate
// module and cannot be relied on by historical version/main test commands.
func TestPinnedReferenceCheckout(t *testing.T) {
	path := os.Getenv("OOXML_REFERENCE_PIN")
	if path == "" {
		_, file, _, ok := runtime.Caller(0)
		if !ok {
			t.Fatal("cannot locate native pin")
		}
		path = filepath.Join(filepath.Dir(file), "..", "..", "spec", "reference-distribution.json")
	}
	b, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	var pin ReferenceIdentity
	if err = json.Unmarshal(b, &pin); err != nil {
		t.Fatal(err)
	}
	candidate := os.Getenv("OOXML_REFERENCE_PIN") != "" && pin.TagObject == "" && strings.HasPrefix(pin.Tag, "candidate-")
	if candidate && os.Getenv("OOXML_FIXTURES_ROOT") == "" {
		t.Fatal("candidate pin requires explicit root")
	}
	if err = VerifyReferenceCheckout(ReferenceRoot(), pin, candidate); err != nil {
		t.Fatal(err)
	}
}

func checkoutGit(t *testing.T, root string, args ...string) string {
	t.Helper()
	base := []string{"-C", root, "-c", "user.name=Rui Carmo", "-c", "user.email=rcarmo@users.noreply.github.com", "-c", "commit.gpgsign=false", "-c", "tag.gpgsign=false"}
	cmd := exec.Command("git", append(base, args...)...)
	b, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("git %v: %v %s", args, err, b)
	}
	return strings.TrimSpace(string(b))
}
func privateReference(t *testing.T) (string, ReferenceIdentity) {
	t.Helper()
	root := t.TempDir()
	checkoutGit(t, root, "init", "--quiet")
	for _, path := range []string{"manifest.json", "shared/v2/pack/pack-manifest.json", "facts/constants.json", "ledgers/workflows.json"} {
		full := filepath.Join(root, path)
		if err := os.MkdirAll(filepath.Dir(full), 0755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(full, []byte("{}\n"), 0644); err != nil {
			t.Fatal(err)
		}
	}
	checkoutGit(t, root, "add", ".")
	checkoutGit(t, root, "commit", "--quiet", "-m", "Native private reference fixture")
	checkoutGit(t, root, "tag", "-a", "v1.0.0", "-m", "Native reference release")
	sum := sha256.Sum256([]byte("{}\n"))
	seal := hex.EncodeToString(sum[:])
	return root, ReferenceIdentity{Commit: checkoutGit(t, root, "rev-parse", "HEAD"), Tag: "v1.0.0", TagObject: checkoutGit(t, root, "rev-parse", "refs/tags/v1.0.0"), Manifest: seal, Pack: seal}
}

func TestReferenceCheckoutIntegrityBatch(t *testing.T) {
	write := func(t *testing.T, root, path string) {
		t.Helper()
		if err := os.WriteFile(filepath.Join(root, path), []byte("changed\n"), 0644); err != nil {
			t.Fatal(err)
		}
	}
	for _, tc := range []struct {
		name              string
		change            func(*testing.T, string, *ReferenceIdentity)
		candidate, refuse bool
	}{
		{"clean annotated release", nil, false, false},
		{"dirty constants outside asset manifest", func(t *testing.T, r string, p *ReferenceIdentity) { write(t, r, "facts/constants.json") }, false, true},
		{"staged workflow outside asset manifest", func(t *testing.T, r string, p *ReferenceIdentity) {
			write(t, r, "ledgers/workflows.json")
			checkoutGit(t, r, "add", "ledgers/workflows.json")
		}, false, true},
		{"deleted tracked workflow", func(t *testing.T, r string, p *ReferenceIdentity) {
			if err := os.Remove(filepath.Join(r, "ledgers/workflows.json")); err != nil {
				t.Fatal(err)
			}
		}, false, true},
		{"untracked fact", func(t *testing.T, r string, p *ReferenceIdentity) { write(t, r, "facts/extra.json") }, false, true},
		{"changed manifest seal", func(t *testing.T, r string, p *ReferenceIdentity) { write(t, r, "manifest.json") }, false, true},
		{"changed pack seal", func(t *testing.T, r string, p *ReferenceIdentity) { write(t, r, "shared/v2/pack/pack-manifest.json") }, false, true},
		{"wrong pinned commit", func(t *testing.T, r string, p *ReferenceIdentity) { p.Commit = strings.Repeat("a", 40) }, false, true},
		{"clean different checkout", func(t *testing.T, r string, p *ReferenceIdentity) {
			write(t, r, "facts/constants.json")
			checkoutGit(t, r, "add", ".")
			checkoutGit(t, r, "commit", "--quiet", "-m", "Different snapshot")
		}, false, true},
		{"lightweight release tag", func(t *testing.T, r string, p *ReferenceIdentity) {
			checkoutGit(t, r, "tag", "-d", p.Tag)
			checkoutGit(t, r, "tag", p.Tag)
			p.TagObject = p.Commit
		}, false, true},
		{"wrong annotated tag object", func(t *testing.T, r string, p *ReferenceIdentity) { p.TagObject = p.Commit }, false, true},
		{"annotated tag points elsewhere", func(t *testing.T, r string, p *ReferenceIdentity) {
			old := p.Commit
			checkoutGit(t, r, "commit", "--allow-empty", "--quiet", "-m", "Other commit")
			checkoutGit(t, r, "tag", "-f", "-a", p.Tag, "-m", "Other release")
			p.TagObject = checkoutGit(t, r, "rev-parse", "refs/tags/"+p.Tag)
			checkoutGit(t, r, "checkout", "--quiet", "--detach", old)
		}, false, true},
		{"missing release annotation", func(t *testing.T, r string, p *ReferenceIdentity) { p.Tag = ""; p.TagObject = "" }, false, true},
		{"invalid commit syntax", func(t *testing.T, r string, p *ReferenceIdentity) { p.Commit = "HEAD" }, false, true},
		{"clean candidate", func(t *testing.T, r string, p *ReferenceIdentity) { p.Tag = "candidate-v0.2.0"; p.TagObject = "" }, true, false},
		{"dirty candidate", func(t *testing.T, r string, p *ReferenceIdentity) {
			p.Tag = "candidate-v0.2.0"
			p.TagObject = ""
			write(t, r, "facts/constants.json")
		}, true, true},
		{"assume unchanged cannot hide edit", func(t *testing.T, r string, p *ReferenceIdentity) {
			checkoutGit(t, r, "update-index", "--assume-unchanged", "facts/constants.json")
			write(t, r, "facts/constants.json")
		}, false, true},
		{"missing git metadata", func(t *testing.T, r string, p *ReferenceIdentity) {
			if err := os.Rename(filepath.Join(r, ".git"), filepath.Join(t.TempDir(), "saved-git")); err != nil {
				t.Fatal(err)
			}
		}, false, true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			root, p := privateReference(t)
			if tc.change != nil {
				tc.change(t, root, &p)
			}
			before := checkoutSnapshot(t, root)
			err := VerifyReferenceCheckout(root, p, tc.candidate)
			if (err != nil) != tc.refuse {
				t.Fatalf("refuse=%v, got %v", tc.refuse, err)
			}
			if after := checkoutSnapshot(t, root); after != before {
				t.Fatal("verification changed checkout bytes")
			}
		})
	}
}

func checkoutSnapshot(t *testing.T, root string) string {
	t.Helper()
	h := sha256.New()
	err := filepath.WalkDir(root, func(path string, d os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if d.IsDir() {
			return nil
		}
		rel, err := filepath.Rel(root, path)
		if err != nil {
			return err
		}
		b, err := os.ReadFile(path)
		if err != nil {
			return err
		}
		h.Write([]byte(rel))
		h.Write(b)
		return nil
	})
	if err != nil {
		t.Fatal(err)
	}
	return hex.EncodeToString(h.Sum(nil))
}
