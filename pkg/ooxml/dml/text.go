package dml

import "encoding/xml"

// TxBody represents text body (text frame).
type TxBody struct {
	XMLName   xml.Name   `xml:"txBody"`
	BodyPr    *BodyPr    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main bodyPr"`
	LstStyle  *LstStyle  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lstStyle,omitempty"`
	P         []*P       `xml:"http://schemas.openxmlformats.org/drawingml/2006/main p,omitempty"`
}

// BodyPr represents text body properties.
type BodyPr struct {
	Rot         int64   `xml:"rot,attr,omitempty"`
	Vert        string  `xml:"vert,attr,omitempty"`        // horz, vert, vert270, wordArtVert
	Wrap        string  `xml:"wrap,attr,omitempty"`        // none, square
	Anchor      string  `xml:"anchor,attr,omitempty"`      // t, ctr, b
	AnchorCtr   *bool   `xml:"anchorCtr,attr,omitempty"`
	LIns        *int64  `xml:"lIns,attr,omitempty"`        // Left inset in EMUs
	TIns        *int64  `xml:"tIns,attr,omitempty"`        // Top inset
	RIns        *int64  `xml:"rIns,attr,omitempty"`        // Right inset
	BIns        *int64  `xml:"bIns,attr,omitempty"`        // Bottom inset
	NumCol      int     `xml:"numCol,attr,omitempty"`
	SpcCol      int64   `xml:"spcCol,attr,omitempty"`
	RtlCol      *bool   `xml:"rtlCol,attr,omitempty"`
	FromWordArt *bool   `xml:"fromWordArt,attr,omitempty"`
	
	NoAutofit    *NoAutofit    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main noAutofit,omitempty"`
	NormAutofit  *NormAutofit  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main normAutofit,omitempty"`
	SpAutoFit    *SpAutoFit    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main spAutoFit,omitempty"`
}

// NoAutofit represents no autofit.
type NoAutofit struct{}

// NormAutofit represents normal autofit.
type NormAutofit struct {
	FontScale int `xml:"fontScale,attr,omitempty"` // Percentage * 1000
	LnSpcReduction int `xml:"lnSpcReduction,attr,omitempty"`
}

// SpAutoFit represents shape autofit.
type SpAutoFit struct{}

// LstStyle represents list style.
type LstStyle struct {
	DefPPr  *PPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main defPPr,omitempty"`
	Lvl1pPr *PPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lvl1pPr,omitempty"`
	Lvl2pPr *PPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lvl2pPr,omitempty"`
	Lvl3pPr *PPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lvl3pPr,omitempty"`
	Lvl4pPr *PPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lvl4pPr,omitempty"`
	Lvl5pPr *PPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lvl5pPr,omitempty"`
	Lvl6pPr *PPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lvl6pPr,omitempty"`
	Lvl7pPr *PPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lvl7pPr,omitempty"`
	Lvl8pPr *PPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lvl8pPr,omitempty"`
	Lvl9pPr *PPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lvl9pPr,omitempty"`
}

// P represents a paragraph.
type P struct {
	PPr      *PPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main pPr,omitempty"`
	R        []*R   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main r,omitempty"`
	Br       []*Br  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main br,omitempty"`
	Fld      []*Fld `xml:"http://schemas.openxmlformats.org/drawingml/2006/main fld,omitempty"`
	EndParaRPr *RPr `xml:"http://schemas.openxmlformats.org/drawingml/2006/main endParaRPr,omitempty"`
}

// PPr represents paragraph properties.
type PPr struct {
	MarL        *int64    `xml:"marL,attr,omitempty"`       // Left margin in EMUs
	MarR        *int64    `xml:"marR,attr,omitempty"`       // Right margin
	Lvl         *int      `xml:"lvl,attr,omitempty"`        // Outline level (0-8)
	Indent      *int64    `xml:"indent,attr,omitempty"`     // First line indent
	Algn        string    `xml:"algn,attr,omitempty"`       // l, ctr, r, just, justLow, dist, thaiDist
	DefTabSz    *int64    `xml:"defTabSz,attr,omitempty"`   // Default tab size
	RtL         *bool     `xml:"rtl,attr,omitempty"`        // Right to left
	EaLnBrk     *bool     `xml:"eaLnBrk,attr,omitempty"`    // East Asian line break
	FontAlgn    string    `xml:"fontAlgn,attr,omitempty"`   // auto, t, ctr, base, b
	LatinLnBrk  *bool     `xml:"latinLnBrk,attr,omitempty"` // Latin line break
	HangingPunct *bool    `xml:"hangingPunct,attr,omitempty"`
	
	LnSpc       *LnSpc    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main lnSpc,omitempty"`
	SpcBef      *Spc      `xml:"http://schemas.openxmlformats.org/drawingml/2006/main spcBef,omitempty"`
	SpcAft      *Spc      `xml:"http://schemas.openxmlformats.org/drawingml/2006/main spcAft,omitempty"`
	BuClr       *BuClr    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main buClr,omitempty"`
	BuSzTx      *BuSzTx   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main buSzTx,omitempty"`
	BuSzPct     *BuSzPct  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main buSzPct,omitempty"`
	BuSzPts     *BuSzPts  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main buSzPts,omitempty"`
	BuFontTx    *BuFontTx `xml:"http://schemas.openxmlformats.org/drawingml/2006/main buFontTx,omitempty"`
	BuFont      *BuFont   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main buFont,omitempty"`
	BuNone      *BuNone   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main buNone,omitempty"`
	BuAutoNum   *BuAutoNum `xml:"http://schemas.openxmlformats.org/drawingml/2006/main buAutoNum,omitempty"`
	BuChar      *BuChar   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main buChar,omitempty"`
	BuBlip      *BuBlip   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main buBlip,omitempty"`
	DefRPr      *RPr      `xml:"http://schemas.openxmlformats.org/drawingml/2006/main defRPr,omitempty"`
}

// LnSpc represents line spacing.
type LnSpc struct {
	SpcPct *SpcPct `xml:"http://schemas.openxmlformats.org/drawingml/2006/main spcPct,omitempty"`
	SpcPts *SpcPts `xml:"http://schemas.openxmlformats.org/drawingml/2006/main spcPts,omitempty"`
}

// Spc represents spacing (before/after).
type Spc struct {
	SpcPct *SpcPct `xml:"http://schemas.openxmlformats.org/drawingml/2006/main spcPct,omitempty"`
	SpcPts *SpcPts `xml:"http://schemas.openxmlformats.org/drawingml/2006/main spcPts,omitempty"`
}

// SpcPct represents spacing percentage.
type SpcPct struct {
	Val int `xml:"val,attr"` // Percentage * 1000
}

// SpcPts represents spacing in points.
type SpcPts struct {
	Val int `xml:"val,attr"` // Points * 100
}

// Bullet-related types
type BuClr struct {
	SrgbClr   *SrgbClr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main srgbClr,omitempty"`
	SchemeClr *SchemeClr `xml:"http://schemas.openxmlformats.org/drawingml/2006/main schemeClr,omitempty"`
}

type BuSzTx struct{}
type BuSzPct struct {
	Val int `xml:"val,attr"` // Percentage * 1000
}
type BuSzPts struct {
	Val int `xml:"val,attr"` // Points * 100
}
type BuFontTx struct{}
type BuFont struct {
	Typeface string `xml:"typeface,attr"`
	Pitchfamily string `xml:"pitchFamily,attr,omitempty"`
	Charset int `xml:"charset,attr,omitempty"`
}
type BuNone struct{}
type BuAutoNum struct {
	Type    string `xml:"type,attr"` // arabicPeriod, alphaLcParenBoth, etc.
	StartAt *int   `xml:"startAt,attr,omitempty"`
}
type BuChar struct {
	Char string `xml:"char,attr"`
}
type BuBlip struct {
	// Blip reference
}

// R represents a text run.
type R struct {
	RPr *RPr  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main rPr,omitempty"`
	T   string `xml:"http://schemas.openxmlformats.org/drawingml/2006/main t"`
}

// RPr represents run properties.
type RPr struct {
	Lang      string     `xml:"lang,attr,omitempty"`
	AltLang   string     `xml:"altLang,attr,omitempty"`
	Sz        *int       `xml:"sz,attr,omitempty"`        // Font size in hundredths of a point
	B         *bool      `xml:"b,attr,omitempty"`         // Bold
	I         *bool      `xml:"i,attr,omitempty"`         // Italic
	U         string     `xml:"u,attr,omitempty"`         // Underline: none, sng, dbl, etc.
	Strike    string     `xml:"strike,attr,omitempty"`    // noStrike, sngStrike, dblStrike
	Kern      *int       `xml:"kern,attr,omitempty"`      // Kerning
	Cap       string     `xml:"cap,attr,omitempty"`       // none, small, all
	Spc       *int       `xml:"spc,attr,omitempty"`       // Character spacing
	Baseline  *int       `xml:"baseline,attr,omitempty"`  // Baseline shift (percentage * 1000)
	NoProof   *bool      `xml:"noProof,attr,omitempty"`
	Dirty     *bool      `xml:"dirty,attr,omitempty"`
	Err       *bool      `xml:"err,attr,omitempty"`
	SmtClean  *bool      `xml:"smtClean,attr,omitempty"`
	SmtId     *int       `xml:"smtId,attr,omitempty"`
	
	Ln        *Ln        `xml:"http://schemas.openxmlformats.org/drawingml/2006/main ln,omitempty"`
	NoFill    *NoFill    `xml:"http://schemas.openxmlformats.org/drawingml/2006/main noFill,omitempty"`
	SolidFill *SolidFill `xml:"http://schemas.openxmlformats.org/drawingml/2006/main solidFill,omitempty"`
	GradFill  *GradFill  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main gradFill,omitempty"`
	EffectLst *EffectLst `xml:"http://schemas.openxmlformats.org/drawingml/2006/main effectLst,omitempty"`
	Highlight *SrgbClr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main highlight,omitempty"`
	Latin     *TextFont  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main latin,omitempty"`
	Ea        *TextFont  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main ea,omitempty"`
	Cs        *TextFont  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main cs,omitempty"`
	Sym       *TextFont  `xml:"http://schemas.openxmlformats.org/drawingml/2006/main sym,omitempty"`
	HlinkClick *HlinkClick `xml:"http://schemas.openxmlformats.org/drawingml/2006/main hlinkClick,omitempty"`
}

// TextFont represents font specification.
type TextFont struct {
	Typeface    string `xml:"typeface,attr"`
	PitchFamily string `xml:"pitchFamily,attr,omitempty"`
	Charset     int    `xml:"charset,attr,omitempty"`
}

// Br represents a break.
type Br struct {
	RPr *RPr `xml:"http://schemas.openxmlformats.org/drawingml/2006/main rPr,omitempty"`
}

// Fld represents a text field.
type Fld struct {
	Type string `xml:"type,attr"` // slidenum, datetime, etc.
	UUID string `xml:"uuid,attr,omitempty"`
	RPr  *RPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main rPr,omitempty"`
	PPr  *PPr   `xml:"http://schemas.openxmlformats.org/drawingml/2006/main pPr,omitempty"`
	T    string `xml:"http://schemas.openxmlformats.org/drawingml/2006/main t,omitempty"`
}
