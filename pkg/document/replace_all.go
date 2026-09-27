package document

import (
	"errors"
	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

// MatchRefusal records a safe refusal for an independently inspected match.
type MatchRefusal struct {
	Text   string `json:"text"`
	Kind   string `json:"kind"`
	Detail string `json:"detail"`
}
type ReplaceAllResult struct {
	Schema   int            `json:"schema"`
	Matched  int            `json:"matched"`
	Changed  int            `json:"changed"`
	Skipped  int            `json:"skipped"`
	Refusals []MatchRefusal `json:"refusals"`
}

func (s *EditSession) ReplaceAll(needle, replacement string, normalized bool) (ReplaceAllResult, error) {
	result := ReplaceAllResult{Schema: 1, Refusals: []MatchRefusal{}}
	matches, err := s.Search(needle, SearchOptions{Normalized: normalized})
	if err != nil {
		return result, err
	}
	result.Matched = len(matches)
	changes := []TextReplacement{}
	for _, target := range matches {
		if target.text == replacement {
			result.Skipped++
			continue
		}
		edits, err := s.planText(target, replacement)
		if err == nil {
			_, err = target.doc.Edit(edits, whitespaceAttributes(edits))
			if err != nil {
				err = editRefusal("unsupported_structure", err.Error())
			}
		}
		if err != nil {
			var refusal *packaging.Refusal
			if !errors.As(err, &refusal) || refusal.Kind == "stale_target" {
				return result, err
			}
			result.Refusals = append(result.Refusals, MatchRefusal{Text: target.text, Kind: refusal.Kind, Detail: refusal.Detail})
			continue
		}
		changes = append(changes, TextReplacement{Target: target, Text: replacement})
	}
	if err = s.ReplaceBatch(changes); err != nil {
		return result, err
	}
	result.Changed = len(changes)
	return result, nil
}
