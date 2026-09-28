package acceptance

import (
	"context"
	"fmt"
	"strings"
	"testing"

	"github.com/cucumber/godog"
	"github.com/rcarmo/go-ooxml/pkg/document"
	"github.com/rcarmo/go-ooxml/pkg/ooxml/common"
)

func documentPropertyGetterSteps(sc *godog.ScenarioContext, shared *tableTextReadbackState) {
	var props *common.CoreProperties
	var setErr, getErr error
	var got *common.CoreProperties
	var sectionDoc document.Document
	var section document.Section
	sc.Before(func(ctx context.Context, _ *godog.Scenario) (context.Context, error) {
		props, got, setErr, getErr, sectionDoc, section = nil, nil, nil, nil, nil, nil
		return ctx, nil
	})
	sc.After(func(ctx context.Context, _ *godog.Scenario, _ error) (context.Context, error) {
		if sectionDoc != nil {
			_ = sectionDoc.Close()
		}
		return ctx, nil
	})
	// The existing new-document step supplies shared.doc for the core case.
	sc.Step(`^its core properties are set to title Doc Title, creator Doc Author and subject Doc Subject$`, func() error {
		if shared.doc == nil {
			return fmt.Errorf("no Word document")
		}
		props = &common.CoreProperties{Title: "Doc Title", Creator: "Doc Author", Subject: "Doc Subject"}
		return nil
	})
	sc.Step(`^description Doc Description, keywords one;two, category Category and language en-US are supplied$`, func() error {
		if props == nil {
			return fmt.Errorf("core properties missing")
		}
		props.Description, props.Keywords, props.Category, props.Language = "Doc Description", "one;two", "Category", "en-US"
		return nil
	})
	sc.Step(`^content status Draft, identifier urn:example:doc, last modifier Reviewer, revision 2 and version 1.0 are supplied$`, func() error {
		if props == nil {
			return fmt.Errorf("core properties missing")
		}
		props.ContentStatus, props.Identifier, props.LastModifiedBy, props.Revision, props.Version = "Draft", "urn:example:doc", "Reviewer", "2", "1.0"
		return nil
	})
	sc.Step(`^created, modified and last-printed W3CDTF timestamps are supplied for 2026-02-03T00:00:00Z, 2026-02-03T01:00:00Z and 2026-02-03T02:00:00Z$`, func() error {
		if props == nil || shared.doc == nil {
			return fmt.Errorf("core properties or document missing")
		}
		props.Created = &common.DCDate{Type: "dcterms:W3CDTF", Value: "2026-02-03T00:00:00Z"}
		props.Modified = &common.DCDate{Type: "dcterms:W3CDTF", Value: "2026-02-03T01:00:00Z"}
		props.LastPrinted = &common.DCDate{Type: "dcterms:W3CDTF", Value: "2026-02-03T02:00:00Z"}
		setErr = shared.doc.SetCoreProperties(props)
		if setErr == nil {
			got, getErr = shared.doc.CoreProperties()
		}
		return nil
	})
	sc.Step(`^the setter and getter return no error$`, func() error {
		return checkCoreErrors(setErr, getErr, got)
	})
	sc.Step(`^only the in-memory title creator and subject are compared to Doc Title, Doc Author and Doc Subject$`, func() error {
		return checkCoreSelected(got)
	})
	sc.Step(`^a new Word document with a first section$`, func() error {
		var err error
		sectionDoc, err = document.New()
		if err != nil {
			return err
		}
		sections := sectionDoc.Sections()
		if len(sections) == 0 || sections[0] == nil {
			return fmt.Errorf("first section missing")
		}
		section = sections[0]
		return nil
	})
	sc.Step(`^TitlePage is set true on that section and BackgroundColor to EEEEEE$`, func() error {
		if sectionDoc == nil || section == nil {
			return fmt.Errorf("first section missing")
		}
		section.SetTitlePage(true)
		sectionDoc.SetBackgroundColor("EEEEEE")
		return nil
	})
	sc.Step(`^the section TitlePage getter is true and the document BackgroundColor getter equals EEEEEE$`, func() error {
		return checkSectionBackground(section, sectionDoc)
	})
}

func checkCoreErrors(setErr, getErr error, got *common.CoreProperties) error {
	if setErr != nil {
		return fmt.Errorf("SetCoreProperties: %w", setErr)
	}
	if getErr != nil {
		return fmt.Errorf("CoreProperties: %w", getErr)
	}
	if got == nil {
		return fmt.Errorf("CoreProperties returned nil")
	}
	return nil
}

func checkCoreSelected(got *common.CoreProperties) error {
	if got == nil {
		return fmt.Errorf("CoreProperties returned nil")
	}
	for _, field := range []struct{ name, value, want string }{
		{"Title", got.Title, "Doc Title"}, {"Creator", got.Creator, "Doc Author"}, {"Subject", got.Subject, "Doc Subject"},
	} {
		if field.value != field.want {
			return fmt.Errorf("%s = %q, want %q", field.name, field.value, field.want)
		}
	}
	return nil
}

func checkSectionBackground(section document.Section, doc document.Document) error {
	if section == nil || doc == nil {
		return fmt.Errorf("first section or document missing")
	}
	if !section.TitlePage() {
		return fmt.Errorf("first section TitlePage getter is false")
	}
	if got := doc.BackgroundColor(); got != "EEEEEE" {
		return fmt.Errorf("BackgroundColor getter = %q, want EEEEEE", got)
	}
	return nil
}

func TestDocumentPropertyGetterNegativeControls(t *testing.T) {
	for _, tc := range []struct {
		name   string
		change func(*common.CoreProperties)
	}{
		{"Title", func(p *common.CoreProperties) { p.Title = "wrong" }},
		{"Creator", func(p *common.CoreProperties) { p.Creator = "wrong" }},
		{"Subject", func(p *common.CoreProperties) { p.Subject = "wrong" }},
	} {
		t.Run(tc.name, func(t *testing.T) {
			doc, err := document.New()
			if err != nil {
				t.Fatal(err)
			}
			defer doc.Close()
			props := &common.CoreProperties{Title: "Doc Title", Creator: "Doc Author", Subject: "Doc Subject"}
			if err := doc.SetCoreProperties(props); err != nil {
				t.Fatal(err)
			}
			got, err := doc.CoreProperties()
			if err := checkCoreErrors(nil, err, got); err != nil {
				t.Fatal(err)
			}
			if err := checkCoreSelected(got); err != nil {
				t.Fatal(err)
			}
			tc.change(got)
			if err := doc.SetCoreProperties(got); err != nil {
				t.Fatal(err)
			}
			changed, err := doc.CoreProperties()
			if err := checkCoreErrors(nil, err, changed); err != nil {
				t.Fatal(err)
			}
			if err := checkCoreSelected(changed); err == nil || !strings.Contains(err.Error(), tc.name) {
				t.Fatalf("wrong %s passed: %v", tc.name, err)
			}
		})
	}
	if err := checkCoreErrors(fmt.Errorf("set failed"), nil, nil); err == nil || !strings.Contains(err.Error(), "SetCoreProperties") {
		t.Fatalf("setter error passed: %v", err)
	}
	if err := checkCoreErrors(nil, fmt.Errorf("get failed"), nil); err == nil || !strings.Contains(err.Error(), "CoreProperties") {
		t.Fatalf("getter error passed: %v", err)
	}
	if err := checkCoreErrors(nil, nil, nil); err == nil {
		t.Fatal("nil getter result passed")
	}
	doc, err := document.New()
	if err != nil {
		t.Fatal(err)
	}
	defer doc.Close()
	sections := doc.Sections()
	if len(sections) == 0 {
		t.Fatal("no first section")
	}
	section := sections[0]
	section.SetTitlePage(true)
	doc.SetBackgroundColor("EEEEEE")
	if err := checkSectionBackground(section, doc); err != nil {
		t.Fatal(err)
	}
	section.SetTitlePage(false)
	if err := checkSectionBackground(section, doc); err == nil || !strings.Contains(err.Error(), "TitlePage") {
		t.Fatalf("cleared TitlePage passed: %v", err)
	}
	section.SetTitlePage(true)
	doc.SetBackgroundColor("FFFFFF")
	if err := checkSectionBackground(section, doc); err == nil || !strings.Contains(err.Error(), "BackgroundColor") {
		t.Fatalf("changed background passed: %v", err)
	}
}
