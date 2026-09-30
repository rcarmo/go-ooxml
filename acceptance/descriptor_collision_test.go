package acceptance

import (
	"archive/zip"
	"bytes"
	"compress/flate"
	"context"
	"encoding/binary"
	"errors"
	"fmt"
	"hash/crc32"
	"io"
	"strings"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/internal/testutil"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func descriptorCollisionSteps(sc *godog.ScenarioContext, reasons *packageReasonState) {
	const signature = uint32(0x08074b50)
	var source, snapshot []byte
	var file *zip.File
	var descriptorAt, centralAt int
	var admissionErr error
	var pkg *packaging.Package
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		source, snapshot, file = nil, nil, nil
		descriptorAt, centralAt = 0, 0
		admissionErr, pkg = nil, nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
		if pkg != nil {
			_ = pkg.Close()
		}
		return ctx, nil
	})
	sc.Step(`^a single-disk ZIP32 archive has one DEFLATED data\.bin member with declared size 7 and stored payload "payload"$`, func() error {
		var err error
		source, err = testutil.UnsignedDescriptorCollisionZIP()
		if err != nil {
			return err
		}
		snapshot = bytes.Clone(source)
		zr, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
		if err != nil || len(zr.File) != 1 {
			return fmt.Errorf("ZIP inventory: %v", err)
		}
		file = zr.File[0]
		if file.Name != "data.bin" || file.Method != zip.Deflate || file.Flags&8 == 0 || file.UncompressedSize64 != 7 {
			return fmt.Errorf("wrong member metadata")
		}
		// Raw decompression is independent of archive/zip's declared CRC check.
		// archive/zip's file.Open is intentionally not used for this check.
		data, err := independentDeflatePayload(source, file)
		if err != nil || string(data) != "payload" {
			return fmt.Errorf("wrong DEFLATED bytes: %v", err)
		}
		return nil
	})
	sc.Step(`^its unsigned twelve-byte data descriptor and central directory both declare CRC32 08074B50, equal to the optional descriptor signature value$`, func() error {
		le := binary.LittleEndian
		if len(source) < 22 || le.Uint32(source[len(source)-22:]) != 0x06054b50 {
			return fmt.Errorf("missing ZIP32 EOCD")
		}
		end := source[len(source)-22:]
		if le.Uint16(end[4:]) != 0 || le.Uint16(end[6:]) != 0 || le.Uint16(end[8:]) != 1 || le.Uint16(end[10:]) != 1 || le.Uint16(end[20:]) != 0 {
			return fmt.Errorf("not single-disk one-member ZIP32")
		}
		centralAt = int(le.Uint32(end[16:]))
		if centralAt < 0 || centralAt > len(source)-22-46 || centralAt+int(le.Uint32(end[12:])) != len(source)-22 || le.Uint32(source[centralAt:]) != 0x02014b50 {
			return fmt.Errorf("central directory extent mismatch")
		}
		offset, err := file.DataOffset()
		if err != nil {
			return err
		}
		descriptorAt = int(offset + int64(file.CompressedSize64))
		if descriptorAt < 0 || descriptorAt > centralAt-12 || descriptorAt+12 != centralAt || le.Uint32(source[descriptorAt:]) != signature || le.Uint32(source[centralAt+16:]) != signature || file.CRC32 != signature {
			return fmt.Errorf("not a 12-byte unsigned CRC/signature collision")
		}
		if le.Uint32(source[descriptorAt+4:]) != uint32(file.CompressedSize64) || le.Uint32(source[descriptorAt+8:]) != uint32(file.UncompressedSize64) || le.Uint32(source[centralAt+20:]) != uint32(file.CompressedSize64) || le.Uint32(source[centralAt+24:]) != 7 {
			return fmt.Errorf("descriptor/directory sizes disagree")
		}
		return nil
	})
	sc.Step(`^independent CRC32 of the decompressed payload differs from 08074B50$`, func() error {
		data, err := independentDeflatePayload(source, file)
		if err != nil || string(data) != "payload" || crc32.ChecksumIEEE(data) == signature {
			return fmt.Errorf("payload CRC not independently wrong: %v", err)
		}
		return nil
	})
	sc.Step(`^ZIP admission validates the descriptor shape and then opens the archive with default bounds$`, func() error {
		if !bytes.Equal(source, snapshot) || file == nil || descriptorAt+12 != centralAt {
			return fmt.Errorf("unverified/changed fixture")
		}
		pkg, admissionErr = packaging.OpenReaderWithLimits(bytes.NewReader(source), int64(len(source)), packaging.Limits{})
		return nil
	})
	sc.Step(`^the unsigned descriptor is recognised as twelve bytes without borrowing a four-byte signature or central-directory bytes$`, func() error {
		var refusal *packaging.Refusal
		if descriptorAt+12 != centralAt || binary.LittleEndian.Uint32(source[centralAt:]) != 0x02014b50 || !errors.As(admissionErr, &refusal) || refusal.Kind != "invalid_package" || strings.Contains(strings.ToLower(refusal.Detail), "descriptor") {
			return fmt.Errorf("geometry rejected or borrowed bytes: %v", admissionErr)
		}
		return nil
	})
	sc.Step(`^complete admission refuses the corrupt payload as a CRC or invalid-package failure, not as an ambiguous descriptor-shape failure$`, func() error {
		var refusal *packaging.Refusal
		if !errors.As(admissionErr, &refusal) || refusal.Kind != "invalid_package" || refusal.Operation != "open" || !strings.Contains(strings.ToLower(refusal.Detail), "checksum") || strings.Contains(strings.ToLower(refusal.Detail), "descriptor") {
			return fmt.Errorf("wrong payload CRC refusal: %v", admissionErr)
		}
		return nil
	})
	sc.Step(`^no package or member payloads are delivered$`, func() error {
		if pkg != nil {
			return fmt.Errorf("partial package or member payload delivered")
		}
		return nil
	})
	sc.Step(`^the caller's original archive bytes remain unchanged$`, func() error {
		if reasons.archive != nil {
			if len(reasons.archive) == 0 || !bytes.Equal(reasons.archive, reasons.original) {
				return fmt.Errorf("profile caller bytes changed")
			}
			return nil
		}
		if !bytes.Equal(source, snapshot) {
			return fmt.Errorf("caller bytes changed")
		}
		return nil
	})
}

func independentDeflatePayload(source []byte, file *zip.File) ([]byte, error) {
	offset, err := file.DataOffset()
	if err != nil {
		return nil, err
	}
	if offset < 0 || offset > int64(len(source)) || file.CompressedSize64 > uint64(int64(len(source))-offset) {
		return nil, fmt.Errorf("compressed range outside archive")
	}
	r := flate.NewReader(bytes.NewReader(source[offset : offset+int64(file.CompressedSize64)]))
	defer r.Close()
	return io.ReadAll(r)
}
