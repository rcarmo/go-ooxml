package testutil

// MutateBytes returns a mutated copy of data using simple strategies for fuzzing.
func MutateBytes(data []byte, offset uint16, xor byte) []byte {
	if len(data) == 0 {
		return data
	}
	mutated := make([]byte, len(data))
	copy(mutated, data)

	index := int(offset) % len(mutated)
	mutated[index] ^= xor

	if len(mutated) > 1 && xor&0x1 == 0x1 {
		swap := (index + 1) % len(mutated)
		mutated[index], mutated[swap] = mutated[swap], mutated[index]
	}
	if len(mutated) > 2 && xor&0x2 == 0x2 {
		mutated[(index+2)%len(mutated)] = 0
	}

	return mutated
}
