package acceptance

import (
	"encoding/json"
	"os"
)

const uniformAPI18Commit = "0531afe2f0bddb50879cf7f11d5f7035e3fe374f"
const uniformAPI18Manifest = "4d2af615e5dab3b98f83dac94a3d35d21f9190ceb49cd7956b2f82e8a7f370b5"

func uniformAPI18Candidate() bool {
	if batch2RootAllowed(os.Getenv("OOXML_FIXTURES_ROOT")) != nil || os.Getenv("OOXML_REFERENCE_PIN") == "" {
		return false
	}
	raw, err := os.ReadFile(os.Getenv("OOXML_REFERENCE_PIN"))
	if err != nil {
		return false
	}
	var p referencePin
	if json.Unmarshal(raw, &p) != nil {
		return false
	}
	return p.Schema == 2 && p.Commit == uniformAPI18Commit && p.Manifest == uniformAPI18Manifest && p.Tag == "candidate-uniform-api18" && p.TagObject == "" && p.Assets == 369
}
