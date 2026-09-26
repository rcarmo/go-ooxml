package packaging

import (
	"archive/zip"
	"fmt"
	"io"
	"math"
)

// Limits supplies opt-in archive eligibility budgets. Zero disables a budget.
// Source bytes and entry count are checked before payload inflation. Applications
// processing untrusted input should choose finite limits for every field.
type Limits struct {
	MaxSourceBytes int64
	MaxEntries     int
	MaxPartBytes   uint64
	MaxTotalBytes  uint64
}

func limitRefusal(part, budget string) error {
	return &Refusal{Kind: "resource_limit", Operation: "open", Part: part, Detail: budget + " budget exceeded"}
}

func validateBudgets(files []*zip.File, limits Limits) error {
	if limits.MaxEntries > 0 && len(files) > limits.MaxEntries {
		return limitRefusal("", "entry count")
	}
	var total uint64
	for _, f := range files {
		if f.FileInfo().IsDir() {
			continue
		}
		n := f.UncompressedSize64
		if n >= math.MaxInt64 || n > uint64(^uint(0)>>1) {
			return limitRefusal(f.Name, "addressable payload")
		}
		if limits.MaxPartBytes > 0 && n > limits.MaxPartBytes {
			return limitRefusal(f.Name, "part bytes")
		}
		if math.MaxUint64-total < n {
			return limitRefusal(f.Name, "total bytes overflow")
		}
		total += n
		if limits.MaxTotalBytes > 0 && total > limits.MaxTotalBytes {
			return limitRefusal(f.Name, "total bytes")
		}
	}
	return nil
}

// OpenReaderWithLimits opens a package under caller-specified resource budgets.
// It does not read external relationships or execute macros.
func OpenReaderWithLimits(r io.ReaderAt, size int64, limits Limits) (*Package, error) {
	if size < 0 || limits.MaxSourceBytes < 0 || limits.MaxEntries < 0 {
		return nil, fmt.Errorf("negative archive size or resource budget")
	}
	if limits.MaxSourceBytes > 0 && size > limits.MaxSourceBytes {
		return nil, limitRefusal("", "source bytes")
	}
	return openReader(r, size, limits)
}
