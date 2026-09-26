package testutil

import (
	"bytes"
	"context"
	"crypto/sha1" // Git blob identity, not a cryptographic trust decision.
	"crypto/sha256"
	"encoding/hex"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"time"
)

// ReferenceIdentity pins the shared distribution. Candidate checks are separate
// from release validation and never grant annotated-release status.
type ReferenceIdentity struct {
	Schema    int    `json:"schema"`
	Commit    string `json:"commit"`
	Tag       string `json:"tag"`
	TagObject string `json:"tag_object"`
	Manifest  string `json:"manifest_sha256"`
	Pack      string `json:"shared_pack_sha256"`
}

// VerifyReferenceCheckout checks the fixture distribution before native suites
// use it. It never changes the checkout, tags, index, or Git configuration.
func VerifyReferenceCheckout(root string, pin ReferenceIdentity, candidate bool) error {
	if !referenceHex(pin.Commit, 40) || !referenceHex(pin.Manifest, 64) || (pin.Schema == 1 && !referenceHex(pin.Pack, 64)) || (pin.Schema == 2 && pin.Pack != "") || (pin.Schema != 1 && pin.Schema != 2) {
		return fmt.Errorf("invalid reference commit or seals")
	}
	if candidate {
		if pin.TagObject != "" || !strings.HasPrefix(pin.Tag, "candidate-") {
			return fmt.Errorf("candidate requires an explicitly untagged candidate pin")
		}
	} else if pin.Tag == "" || strings.HasPrefix(pin.Tag, "candidate-") || !referenceHex(pin.TagObject, 40) {
		return fmt.Errorf("release requires pinned annotated tag")
	}
	root, err := filepath.EvalSymlinks(root)
	if err != nil {
		return err
	}
	root, err = filepath.Abs(root)
	if err != nil {
		return err
	}
	if _, err = os.Lstat(filepath.Join(root, ".git")); err != nil {
		return fmt.Errorf("reference checkout lacks its own git metadata: %w", err)
	}
	top, err := referenceGit(root, "rev-parse", "--show-toplevel")
	if err != nil {
		return err
	}
	topPath, err := filepath.EvalSymlinks(strings.TrimSpace(string(top)))
	if err != nil || topPath != root {
		return fmt.Errorf("reference root is not repository root")
	}
	head, err := referenceGit(root, "rev-parse", "--verify", "HEAD^{commit}")
	if err != nil {
		return err
	}
	if strings.TrimSpace(string(head)) != pin.Commit {
		return fmt.Errorf("reference HEAD differs from pinned commit")
	}
	if !candidate {
		ref := "refs/tags/" + pin.Tag
		if _, err = referenceGit(root, "check-ref-format", ref); err != nil {
			return err
		}
		object, err := referenceGit(root, "rev-parse", "--verify", ref)
		if err != nil {
			return err
		}
		if strings.TrimSpace(string(object)) != pin.TagObject {
			return fmt.Errorf("reference tag object differs")
		}
		kind, err := referenceGit(root, "cat-file", "-t", pin.TagObject)
		if err != nil {
			return err
		}
		if strings.TrimSpace(string(kind)) != "tag" {
			return fmt.Errorf("reference tag is not annotated")
		}
		commit, err := referenceGit(root, "rev-parse", "--verify", ref+"^{commit}")
		if err != nil {
			return err
		}
		if strings.TrimSpace(string(commit)) != pin.Commit {
			return fmt.Errorf("reference tag does not peel to pinned commit")
		}
	}
	if err = referenceClean(root); err != nil {
		return err
	}
	if err = referenceTrackedBytes(root); err != nil {
		return err
	}
	seals := []struct{ path, hash string }{{"manifest.json", pin.Manifest}}
	if pin.Schema == 1 {
		seals = append(seals, struct{ path, hash string }{"shared/v2/pack/pack-manifest.json", pin.Pack})
	}
	for _, f := range seals {
		b, err := os.ReadFile(filepath.Join(root, filepath.FromSlash(f.path)))
		if err != nil {
			return err
		}
		h := sha256.Sum256(b)
		if hex.EncodeToString(h[:]) != f.hash {
			return fmt.Errorf("reference seal mismatch: %s", f.path)
		}
	}
	return referenceClean(root)
}

func referenceHex(s string, n int) bool {
	if len(s) != n || strings.ToLower(s) != s {
		return false
	}
	_, err := hex.DecodeString(s)
	return err == nil
}

func referenceGit(root string, args ...string) ([]byte, error) {
	ctx, cancel := context.WithTimeout(context.Background(), 15*time.Second)
	defer cancel()
	cmd := exec.CommandContext(ctx, "git", append([]string{"-C", root, "-c", "core.fsmonitor=false"}, args...)...)
	for _, e := range os.Environ() {
		if !strings.HasPrefix(e, "GIT_") {
			cmd.Env = append(cmd.Env, e)
		}
	}
	cmd.Env = append(cmd.Env, "GIT_OPTIONAL_LOCKS=0", "GIT_TERMINAL_PROMPT=0")
	b, err := cmd.CombinedOutput()
	if err != nil {
		return nil, fmt.Errorf("reference git %v: %w: %s", args, err, b)
	}
	return b, nil
}
func referenceClean(root string) error {
	b, err := referenceGit(root, "status", "--porcelain=v1", "--untracked-files=all", "--ignore-submodules=none")
	if err != nil {
		return err
	}
	if len(bytes.TrimSpace(b)) != 0 {
		return fmt.Errorf("dirty reference checkout: %s", b)
	}
	return nil
}

// Compare every tracked regular file with the pinned Git tree, independent of
// index assume-unchanged/skip-worktree flags, stat caches, or clean filters.
func referenceTrackedBytes(root string) error {
	b, err := referenceGit(root, "ls-tree", "-r", "-z", "--full-tree", "HEAD")
	if err != nil {
		return err
	}
	for _, row := range bytes.Split(b, []byte{0}) {
		if len(row) == 0 {
			continue
		}
		header, name, ok := bytes.Cut(row, []byte{'\t'})
		fields := strings.Fields(string(header))
		path := string(name)
		if !ok || len(fields) != 3 || fields[1] != "blob" || (fields[0] != "100644" && fields[0] != "100755") || !filepath.IsLocal(path) || strings.Contains(path, "\\") {
			return fmt.Errorf("unsupported reference tree entry %q", row)
		}
		full := filepath.Join(root, filepath.FromSlash(path))
		resolved, err := filepath.EvalSymlinks(full)
		if err != nil {
			return err
		}
		if resolved != full {
			return fmt.Errorf("symlink in reference tree: %s", path)
		}
		info, err := os.Lstat(full)
		if err != nil {
			return err
		}
		if !info.Mode().IsRegular() {
			return fmt.Errorf("nonregular tracked reference: %s", path)
		}
		if runtime.GOOS != "windows" && (info.Mode().Perm()&0111 != 0) != (fields[0] == "100755") {
			return fmt.Errorf("reference executable mode differs: %s", path)
		}
		f, err := os.Open(full)
		if err != nil {
			return err
		}
		h := sha1.New()
		_, _ = fmt.Fprintf(h, "blob %d%c", info.Size(), 0)
		_, err = io.Copy(h, f)
		closeErr := f.Close()
		if err != nil {
			return err
		}
		if closeErr != nil {
			return closeErr
		}
		if hex.EncodeToString(h.Sum(nil)) != fields[2] {
			return fmt.Errorf("tracked reference bytes differ: %s", path)
		}
	}
	return nil
}
