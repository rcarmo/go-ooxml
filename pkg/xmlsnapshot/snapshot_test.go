package xmlsnapshot_test

import (
	"bytes"
	"encoding/xml"
	"errors"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/xmlsnapshot"
)

func TestParseFailureClassificationAndCustody(t *testing.T) {
	t.Run("mismatched tag", func(t *testing.T) {
		source := []byte("<a></b>")
		original := bytes.Clone(source)
		result, err := xmlsnapshot.Parse(source)
		var classified *xmlsnapshot.ParseError
		if result != nil || !errors.As(err, &classified) || classified.Category != xmlsnapshot.CategoryMalformedXML || !bytes.Equal(source, original) {
			t.Fatalf("result=%v error=%v source=%q", result, err, source)
		}
	})
	t.Run("decoder syntax error", func(t *testing.T) {
		source := []byte("<a>&bogus;</a>")
		result, err := xmlsnapshot.Parse(source)
		var classified *xmlsnapshot.ParseError
		if result != nil || !errors.As(err, &classified) || classified.Category != xmlsnapshot.CategoryMalformedXML || string(source) != "<a>&bogus;</a>" {
			t.Fatalf("result=%v error=%v", result, err)
		}
	})
	t.Run("valid document", func(t *testing.T) {
		result, err := xmlsnapshot.Parse([]byte("<a/>"))
		if err != nil || result == nil || result.Root().Name() != (xml.Name{Local: "a"}) {
			t.Fatalf("result=%v error=%v", result, err)
		}
	})
	t.Run("unrelated parse failure remains unclassified", func(t *testing.T) {
		result, err := xmlsnapshot.Parse([]byte("<!DOCTYPE a><a/>"))
		var classified *xmlsnapshot.ParseError
		if result != nil || err == nil || errors.As(err, &classified) {
			t.Fatalf("result=%v error=%v", result, err)
		}
	})
}

func TestSpecialAttributeNamesAreLiteralAndIsolated(t *testing.T) {
	source := []byte(`<r __proto__="polluted" constructor="safe"/>`)
	original := bytes.Clone(source)
	doc, err := xmlsnapshot.Parse(source)
	if err != nil || doc == nil {
		t.Fatal(err)
	}
	source[1] = 'x'
	root := doc.Root()
	attrs := root.Attributes()
	if len(attrs) != 2 || attrs[0] != (xml.Attr{Name: xml.Name{Local: "__proto__"}, Value: "polluted"}) || attrs[1] != (xml.Attr{Name: xml.Name{Local: "constructor"}, Value: "safe"}) {
		t.Fatalf("attributes=%v", attrs)
	}
	attrs[0].Value = "corrupt"
	fresh := root.Attributes()
	text, _ := root.Text()
	if len(fresh) != 2 || fresh[0].Value != "polluted" || fresh[1].Value != "safe" || root.Name() != (xml.Name{Local: "r"}) || text != "" || len(root.Children()) != 0 || !bytes.Equal(doc.Source(), original) {
		t.Fatalf("fresh=%v root=%v text=%q", fresh, root.Name(), text)
	}
	returned := doc.Source()
	returned[0] = 'x'
	if !bytes.Equal(doc.Source(), original) {
		t.Fatal("returned source aliases model")
	}
}

func TestNamespaceMetadataCannotChangeModel(t *testing.T) {
	source := []byte(`<r xmlns:a="urn:a" a:id="outer"/>`)
	original := bytes.Clone(source)
	doc, err := xmlsnapshot.Parse(source)
	if err != nil || doc == nil {
		t.Fatal(err)
	}
	root := doc.Root()
	assertOriginal := func() {
		t.Helper()
		value, found := root.Attribute("urn:a", "id")
		metadata := root.AttributeNamespaces()
		text, _ := root.Text()
		if value != "outer" || !found || len(metadata) != 1 || metadata["a:id"] != "urn:a" || root.Name() != (xml.Name{Local: "r"}) || text != "" || len(root.Children()) != 0 || !bytes.Equal(doc.Source(), original) || !bytes.Equal(source, original) {
			t.Fatalf("value=%q found=%t metadata=%v", value, found, metadata)
		}
	}
	assertOriginal()
	returned := root.AttributeNamespaces()
	returned["a:id"] = "urn:changed"
	assertOriginal() // same model after replacement
	delete(returned, "a:id")
	assertOriginal() // same model after deletion
	attrs := root.Attributes()
	attrs[0].Value = "corrupt"
	assertOriginal()
	fresh, err := xmlsnapshot.Parse(source)
	if err != nil || fresh == nil || fresh.Root().AttributeNamespaces()["a:id"] != "urn:a" {
		t.Fatal("fresh parse did not preserve namespace", err)
	}
}
