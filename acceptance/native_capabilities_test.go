package acceptance

import (
	"encoding/json"
	"fmt"
	"os"
	"testing"
)

type contractLink struct {
	ID          string   `json:"id"`
	Families    []string `json:"families"`
	GoAPI       string   `json:"go_api"`
	ScenarioIDs []string `json:"scenario_ids"`
	State       string   `json:"state"`
	Limits      string   `json:"limits"`
}

func validateLinks(links []contractLink, ids map[string]bool) error {
	seen := map[string]bool{}
	for _, link := range links {
		key := link.ID
		if key == "" {
			return fmt.Errorf("missing capability ID %s", key)
		}
		if seen[key] {
			return fmt.Errorf("duplicate capability mapping %s", key)
		}
		seen[key] = true
		if link.State != "partial" || link.Limits == "" || link.GoAPI == "" || len(link.Families) == 0 || len(link.ScenarioIDs) == 0 {
			return fmt.Errorf("unqualified link %s", key)
		}
		for _, id := range link.ScenarioIDs {
			if !ids[id] {
				return fmt.Errorf("unimplemented scenario %s", id)
			}
		}
	}
	return nil
}
func TestNativeCapabilityLinks(t *testing.T) {
	raw, err := os.ReadFile("../spec/native-capabilities.json")
	if err != nil {
		t.Fatal(err)
	}
	var doc struct {
		Schema int            `json:"schema"`
		Links  []contractLink `json:"links"`
	}
	if err = json.Unmarshal(raw, &doc); err != nil {
		t.Fatal(err)
	}
	if doc.Schema != 1 || len(doc.Links) == 0 {
		t.Fatal("empty or invalid mapping")
	}
	implemented, _, err := inventoryCases()
	if err != nil {
		t.Fatal(err)
	}
	ids := map[string]bool{}
	for c := range implemented {
		ids[c.ID] = true
	}
	if err = validateLinks(doc.Links, ids); err != nil {
		t.Fatal(err)
	}
	for _, change := range []func(*contractLink){func(l *contractLink) { l.ID = "" }, func(l *contractLink) { l.ScenarioIDs = []string{"@MISSING-001"} }, func(l *contractLink) { l.State = "complete" }, func(l *contractLink) { l.Limits = "" }} {
		bad := doc.Links[0]
		change(&bad)
		if err := validateLinks([]contractLink{bad}, ids); err == nil {
			t.Fatal("bad link accepted")
		}
	}
	if err := validateLinks([]contractLink{doc.Links[0], doc.Links[0]}, ids); err == nil {
		t.Fatal("duplicate accepted")
	}
}
