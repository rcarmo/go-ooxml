package spreadsheet

import (
	"bytes"
	"image"
	"image/color"
	"image/jpeg"
	"image/png"
	"testing"

	"github.com/rcarmo/go-ooxml/pkg/packaging"
)

func imageEditFixture(t *testing.T) (*EditSession, []byte, []byte) {
	t.Helper()
	var original, replacement bytes.Buffer
	im := image.NewRGBA(image.Rect(0, 0, 2, 2))
	im.Set(0, 0, color.RGBA{R: 255, A: 255})
	if err := png.Encode(&original, im); err != nil {
		t.Fatal(err)
	}
	if err := jpeg.Encode(&replacement, im, nil); err != nil {
		t.Fatal(err)
	}
	q := packaging.New()
	_, _ = q.AddPart("xl/workbook.xml", packaging.ContentTypeWorkbook, []byte(`<workbook xmlns="`+packaging.NSSpreadsheetML+`" xmlns:r="`+packaging.NSDocumentRelationships+`"><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`))
	_, _ = q.AddPart("xl/sheet.xml", packaging.ContentTypeWorksheet, []byte(`<worksheet xmlns="`+packaging.NSSpreadsheetML+`" xmlns:r="`+packaging.NSDocumentRelationships+`"><sheetData/><drawing r:id="rId1"/></worksheet>`))
	_, _ = q.AddPart("xl/drawing.xml", packaging.ContentTypeDrawing, []byte(`<x:wsDr xmlns:x="`+xdrNS+`" xmlns:a="`+packaging.NSDrawingML+`" xmlns:r="`+packaging.NSDocumentRelationships+`"><x:absoluteAnchor><x:pos x="0" y="0"/><x:ext cx="1" cy="1"/><x:pic><x:nvPicPr><x:cNvPr id="1" name="Picture"/><x:cNvPicPr/></x:nvPicPr><x:blipFill><a:blip r:embed="rId1"/></x:blipFill><x:spPr/></x:pic><x:clientData/></x:absoluteAnchor></x:wsDr>`))
	_, _ = q.AddPart("xl/media/original.png", packaging.ContentTypePNG, original.Bytes())
	_, _ = q.AddPart("xl/media/IMAGE1.jpg", packaging.ContentTypeJPEG, replacement.Bytes())
	q.AddRelationship("", "xl/workbook.xml", packaging.RelTypeOfficeDocument)
	q.AddRelationship("xl/workbook.xml", "sheet.xml", packaging.RelTypeWorksheet)
	q.AddRelationship("xl/sheet.xml", "drawing.xml", packaging.RelTypeDrawing)
	q.AddRelationship("xl/drawing.xml", "media/original.png", packaging.RelTypeImage)
	var b bytes.Buffer
	if err := q.WriteTo(&b); err != nil {
		t.Fatal(err)
	}
	s, err := OpenEditing(b.Bytes(), packaging.Limits{})
	if err != nil {
		t.Fatal(err)
	}
	return s, original.Bytes(), replacement.Bytes()
}
func TestImageReplacementBatch(t *testing.T) {
	t.Run("PNG no-op JPEG replacement collision and handles", func(t *testing.T) {
		s, original, replacement := imageEditFixture(t)
		target, err := s.FindImage("S", 1)
		if err != nil {
			t.Fatal(err)
		}
		if err = s.ReplaceImage(target, original); err != nil {
			t.Fatal(err)
		}
		if len(s.pkg.Receipt().Changes) != 0 || target.consumed {
			t.Fatal("no-op mutated")
		}
		other, _, _ := imageEditFixture(t)
		if other.ReplaceImage(target, replacement) == nil {
			t.Fatal("foreign target accepted")
		}
		if err = s.ReplaceImage(target, replacement); err != nil {
			t.Fatal(err)
		}
		g, err := s.pkg.Graph()
		if err != nil {
			t.Fatal(err)
		}
		found := false
		for _, p := range g.Parts {
			if p.Name == "xl/media/image2.jpg" && p.ContentType == packaging.ContentTypeJPEG {
				found = true
			}
		}
		if !found {
			t.Fatal("case-safe allocation/MIME wrong")
		}
		if s.ReplaceImage(target, replacement) == nil {
			t.Fatal("consumed target reused")
		}
		fresh, err := s.FindImage("S", 1)
		if err != nil {
			t.Fatal(err)
		}
		if err = s.ReplaceImage(fresh, replacement); err != nil {
			t.Fatal(err)
		}
	})
	t.Run("truncated payload refusal keeps target reusable", func(t *testing.T) {
		s, _, replacement := imageEditFixture(t)
		target, err := s.FindImage("S", 1)
		if err != nil {
			t.Fatal(err)
		}
		if s.ReplaceImage(target, replacement[:len(replacement)/2]) == nil {
			t.Fatal("truncated JPEG accepted")
		}
		if target.consumed || len(s.pkg.Receipt().Changes) != 0 {
			t.Fatal("refusal mutated")
		}
		if err = s.ReplaceImage(target, replacement); err != nil {
			t.Fatal(err)
		}
	})
	t.Run("one replacement stales other held handles", func(t *testing.T) {
		s, _, replacement := imageEditFixture(t)
		a, err := s.FindImage("S", 1)
		if err != nil {
			t.Fatal(err)
		}
		b, err := s.FindImage("S", 1)
		if err != nil {
			t.Fatal(err)
		}
		if err = s.ReplaceImage(a, replacement); err != nil {
			t.Fatal(err)
		}
		if s.ReplaceImage(b, replacement) == nil {
			t.Fatal("other handle not stale")
		}
	})
}
