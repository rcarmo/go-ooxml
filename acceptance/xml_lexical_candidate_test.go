package acceptance

import (
	"encoding/json"
	"os"
)

// Only the exact clean sealed shared candidate enables the new 20-ID lane.
// This never changes the published default pin or the earlier candidate lanes.
const xmlLexicalCommit = "6b601a4252a90759c2e73c31cc9d8d10e6bca0bd"
const xmlLexicalManifest = "6504c9e64974644c85ff098a8f540c084bbe7ec9a756b44c8de594d897a24072"

func xmlLexicalCandidate() bool {
	path := os.Getenv("OOXML_REFERENCE_PIN")
	if path == "" || os.Getenv("OOXML_FIXTURES_ROOT") == "" {
		return false
	}
	b, err := os.ReadFile(path)
	if err != nil {
		return false
	}
	var pin referencePin
	if json.Unmarshal(b, &pin) != nil {
		return false
	}
	// Marker alone is not authority. loadReferencePin verifies exact HEAD,
	// manifest, tracked bytes and clean root before any selected case executes.
	return pin.Schema == 2 && pin.Commit == xmlLexicalCommit && pin.Manifest == xmlLexicalManifest && pin.TagObject == "" && pin.Tag == "candidate-xml-lexical-alignment"
}
