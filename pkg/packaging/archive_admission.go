package packaging

import (
	"bytes"
	"fmt"
)

// ArchiveAdmissionLimits is a bounded raw-ZIP/XML member profile distinct from
// legacy Limits and ZIP32Limits. Omitted fields use finite ZIP32 defaults; an
// explicitly set zero member count admits no entries. This does not require an
// OPC content-type registry and never relaxes the existing Envelope profile.
type ArchiveAdmissionLimits struct {
	MaxMembers      *int
	MaxMemberBytes  *uint64
	MaxTotalBytes   *uint64
	MaxRatio        *float64
	MaxArchiveBytes *int64
}

// AdmitArchive checks ZIP32 geometry and every XML/.rels member before returning
// detached entries. Zero is meaningful for an explicitly supplied budget.
func AdmitArchive(source []byte, limits ArchiveAdmissionLimits) ([]ZIP32Entry, error) {
	opts := ZIP32Limits{}
	if limits.MaxMembers != nil {
		if *limits.MaxMembers < 0 {
			return nil, fmt.Errorf("negative member budget")
		}
		if *limits.MaxMembers == 0 {
			return nil, zipReason("zip-too-many-entries", "")
		}
		opts.MaxEntries = *limits.MaxMembers
	}
	if limits.MaxMemberBytes != nil {
		if *limits.MaxMemberBytes == 0 {
			return nil, zipReason("zip-entry-too-large", "")
		}
		opts.MaxEntryBytes = *limits.MaxMemberBytes
	}
	if limits.MaxTotalBytes != nil {
		if *limits.MaxTotalBytes == 0 {
			return nil, zipReason("zip-total-too-large", "")
		}
		opts.MaxTotalBytes = *limits.MaxTotalBytes
	}
	if limits.MaxRatio != nil {
		if *limits.MaxRatio <= 0 {
			return nil, fmt.Errorf("invalid compression-ratio budget")
		}
		opts.MaxCompressionRatio = *limits.MaxRatio
	}
	if limits.MaxArchiveBytes != nil {
		if *limits.MaxArchiveBytes < 0 {
			return nil, fmt.Errorf("negative archive budget")
		}
		if *limits.MaxArchiveBytes == 0 {
			return nil, zipReason("zip-archive-too-large", "")
		}
		opts.MaxArchiveBytes = *limits.MaxArchiveBytes
	}
	return ReadXMLMembers(bytes.Clone(source), opts)
}
