package packaging

import "hash/crc32"

// ZIPCRC32 returns the unsigned IEEE CRC-32 used by ZIP for the exact payload
// bytes supplied by the caller. It does not modify or retain the input.
func ZIPCRC32(payload []byte) uint32 { return crc32.ChecksumIEEE(payload) }
