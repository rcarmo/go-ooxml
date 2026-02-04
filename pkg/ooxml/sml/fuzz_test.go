package sml

import (
	"strings"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/utils"
)

func FuzzUnmarshalWorksheet(f *testing.F) {
	seed := `<worksheet xmlns="` + NS + `"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`
	f.Add(seed)

	f.Fuzz(func(t *testing.T, xmlInput string) {
		if !strings.Contains(xmlInput, "<worksheet") || !strings.Contains(xmlInput, NS) {
			return
		}
		var ws Worksheet
		if err := utils.UnmarshalXML([]byte(xmlInput), &ws); err != nil {
			return
		}
		out, err := utils.MarshalXMLWithHeader(&ws)
		if err != nil {
			t.Fatalf("MarshalXMLWithHeader error: %v", err)
		}
		if !strings.Contains(string(out), "worksheet") {
			t.Fatalf("round-trip missing worksheet")
		}
	})
}

func FuzzUnmarshalTable(f *testing.F) {
	seed := `<table xmlns="` + NS + `" id="1" name="Table1" displayName="Table1" ref="A1:B2"><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>`
	f.Add(seed)

	f.Fuzz(func(t *testing.T, xmlInput string) {
		if !strings.Contains(xmlInput, "<table") || !strings.Contains(xmlInput, NS) {
			return
		}
		var tbl Table
		if err := utils.UnmarshalXML([]byte(xmlInput), &tbl); err != nil {
			return
		}
		out, err := utils.MarshalXMLWithHeader(&tbl)
		if err != nil {
			t.Fatalf("MarshalXMLWithHeader error: %v", err)
		}
		if !strings.Contains(string(out), "table") {
			t.Fatalf("round-trip missing table")
		}
	})
}

