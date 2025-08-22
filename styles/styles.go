package styles

import (
	"encoding/xml"
)

// Styles structure for defining heading styles
type Styles struct {
	XMLName xml.Name `xml:"w:styles"`
	XmlnsW  string   `xml:"xmlns:w,attr"`
	XmlnsR  string   `xml:"xmlns:r,attr,omitempty"`
	Styles  []Style  `xml:"w:style"`
}

type Style struct {
	Type        string        `xml:"w:type,attr"`
	StyleId     string        `xml:"w:styleId,attr"`
	Default     string        `xml:"w:default,attr,omitempty"`
	CustomStyle string        `xml:"w:customStyle,attr,omitempty"`
	Name        StyleName     `xml:"w:name"`
	BasedOn     *StyleBasedOn `xml:"w:basedOn,omitempty"`
	Next        *StyleNext    `xml:"w:next,omitempty"`
	Link        *StyleLink    `xml:"w:link,omitempty"`
	UiPriority  *UiPriority   `xml:"w:uiPriority,omitempty"`
	QFormat     *QFormat      `xml:"w:qFormat,omitempty"`
	StylePPr    *StylePPr     `xml:"w:pPr,omitempty"`
	StyleRPr    *StyleRPr     `xml:"w:rPr,omitempty"`
}

type StyleName struct {
	Val string `xml:"w:val,attr"`
}

type StyleBasedOn struct {
	Val string `xml:"w:val,attr"`
}

type StyleNext struct {
	Val string `xml:"w:val,attr"`
}

type StyleLink struct {
	Val string `xml:"w:val,attr"`
}

type UiPriority struct {
	Val string `xml:"w:val,attr"`
}

type QFormat struct{}

type StylePPr struct {
	KeepNext      *KeepNext      `xml:"w:keepNext,omitempty"`
	KeepLines     *KeepLines     `xml:"w:keepLines,omitempty"`
	SpacingStyle  *SpacingStyle  `xml:"w:spacing,omitempty"`
	OutlineLevel  *OutlineLevel  `xml:"w:outlineLvl,omitempty"`
	Justification *Justification `xml:"w:jc,omitempty"`
	Ind           *Indentation   `xml:"w:ind,omitempty"`
}

type StyleRPr struct {
	RFonts    *RFonts    `xml:"w:rFonts,omitempty"`
	Bold      *Bold      `xml:"w:b,omitempty"`
	BoldCs    *Bold      `xml:"w:bCs,omitempty"`
	Italic    *Italic    `xml:"w:i,omitempty"`
	ItalicCs  *Italic    `xml:"w:iCs,omitempty"`
	Size      *Size      `xml:"w:sz,omitempty"`
	SizeCs    *Size      `xml:"w:szCs,omitempty"`
	Color     *Color     `xml:"w:color,omitempty"`
	Underline *Underline `xml:"w:u,omitempty"`
}

type KeepNext struct{}
type KeepLines struct{}
type Bold struct{}
type Italic struct{}

type SpacingStyle struct {
	Before   string `xml:"w:before,attr,omitempty"`
	After    string `xml:"w:after,attr,omitempty"`
	Line     string `xml:"w:line,attr,omitempty"`
	LineRule string `xml:"w:lineRule,attr,omitempty"`
}

type OutlineLevel struct {
	Val string `xml:"w:val,attr"`
}

type Size struct {
	Val string `xml:"w:val,attr"`
}

type Color struct {
	Val        string `xml:"w:val,attr"`
	ThemeColor string `xml:"w:themeColor,attr,omitempty"`
	ThemeShade string `xml:"w:themeShade,attr,omitempty"`
}

type RFonts struct {
	Ascii    string `xml:"w:ascii,attr,omitempty"`
	HAnsi    string `xml:"w:hAnsi,attr,omitempty"`
	Cs       string `xml:"w:cs,attr,omitempty"`
	EastAsia string `xml:"w:eastAsia,attr,omitempty"`
}

type Justification struct {
	Val string `xml:"w:val,attr"`
}

type Indentation struct {
	Left      string `xml:"w:left,attr,omitempty"`
	Right     string `xml:"w:right,attr,omitempty"`
	FirstLine string `xml:"w:firstLine,attr,omitempty"`
	Hanging   string `xml:"w:hanging,attr,omitempty"`
}

type Underline struct {
	Val string `xml:"w:val,attr"`
}

func normalStyle() Style {
	return Style{
		Type:    "paragraph",
		StyleId: "Normal",
		Default: "1",
		Name:    StyleName{Val: "Normal"},
		StylePPr: &StylePPr{
			SpacingStyle: &SpacingStyle{
				// After:    "160",
				Line:     "259",
				LineRule: "auto",
			},
		},
		StyleRPr: &StyleRPr{
			RFonts: &RFonts{
				Ascii: "Calibri",
				HAnsi: "Calibri",
				Cs:    "Calibri",
			},
			Size:   &Size{Val: "22"}, // 11pt
			SizeCs: &Size{Val: "22"},
		},
	}
}

// heading1Style
func heading1Style() Style {
	return Style{
		Type:       "paragraph",
		StyleId:    "Heading1",
		Name:       StyleName{Val: "Heading 1"},
		BasedOn:    &StyleBasedOn{Val: "Normal"},
		Next:       &StyleNext{Val: "Normal"},
		Link:       &StyleLink{Val: "Heading1Char"},
		UiPriority: &UiPriority{Val: "9"},
		QFormat:    &QFormat{},
		StylePPr: &StylePPr{
			KeepNext:     &KeepNext{},
			KeepLines:    &KeepLines{},
			SpacingStyle: &SpacingStyle{Before: "480", After: "240"},
			OutlineLevel: &OutlineLevel{Val: "0"},
		},
		StyleRPr: &StyleRPr{
			Bold:   &Bold{},
			BoldCs: &Bold{},
			Size:   &Size{Val: "32"}, // 16pt
			SizeCs: &Size{Val: "32"},
			Color:  &Color{Val: "2F5496"},
		},
	}
}

func heading2Style() Style {
	return Style{
		Type:       "paragraph",
		StyleId:    "Heading2",
		Name:       StyleName{Val: "Heading 2"},
		BasedOn:    &StyleBasedOn{Val: "Normal"},
		Next:       &StyleNext{Val: "Normal"},
		Link:       &StyleLink{Val: "Heading2Char"},
		UiPriority: &UiPriority{Val: "9"},
		QFormat:    &QFormat{},
		StylePPr: &StylePPr{
			KeepNext:     &KeepNext{},
			KeepLines:    &KeepLines{},
			SpacingStyle: &SpacingStyle{Before: "360", After: "180"},
			OutlineLevel: &OutlineLevel{Val: "1"},
		},
		StyleRPr: &StyleRPr{
			Bold:   &Bold{},
			BoldCs: &Bold{},
			Size:   &Size{Val: "28"}, // 14pt
			SizeCs: &Size{Val: "28"},
			Color:  &Color{Val: "2F5496"},
		},
	}
}

func heading3Style() Style {
	return Style{
		Type:       "paragraph",
		StyleId:    "Heading3",
		Name:       StyleName{Val: "Heading 3"},
		BasedOn:    &StyleBasedOn{Val: "Normal"},
		Next:       &StyleNext{Val: "Normal"},
		Link:       &StyleLink{Val: "Heading3Char"},
		UiPriority: &UiPriority{Val: "9"},
		QFormat:    &QFormat{},
		StylePPr: &StylePPr{
			KeepNext:     &KeepNext{},
			KeepLines:    &KeepLines{},
			SpacingStyle: &SpacingStyle{Before: "240", After: "120"},
			OutlineLevel: &OutlineLevel{Val: "2"},
		},
		StyleRPr: &StyleRPr{
			Bold:   &Bold{},
			BoldCs: &Bold{},
			Size:   &Size{Val: "24"}, // 12pt
			SizeCs: &Size{Val: "24"},
			Color:  &Color{Val: "1F3763"},
		},
	}
}

func heading4Style() Style {
	return Style{
		Type:    "paragraph",
		StyleId: "Heading4",
		Name:    StyleName{Val: "Heading 4"},
		BasedOn: &StyleBasedOn{Val: "Normal"},
		Next:    &StyleNext{Val: "Normal"},
		StylePPr: &StylePPr{
			KeepNext:     &KeepNext{},
			KeepLines:    &KeepLines{},
			SpacingStyle: &SpacingStyle{Before: "240", After: "120"},
			OutlineLevel: &OutlineLevel{Val: "3"},
		},
		StyleRPr: &StyleRPr{
			Bold:     &Bold{},
			BoldCs:   &Bold{},
			Italic:   &Italic{},
			ItalicCs: &Italic{},
			Size:     &Size{Val: "22"}, // 11pt
			SizeCs:   &Size{Val: "22"},
		},
	}
}

func heading5Style() Style {
	return Style{
		Type:    "paragraph",
		StyleId: "Heading5",
		Name:    StyleName{Val: "Heading 5"},
		BasedOn: &StyleBasedOn{Val: "Normal"},
		Next:    &StyleNext{Val: "Normal"},
		StylePPr: &StylePPr{
			SpacingStyle: &SpacingStyle{Before: "240", After: "120"},
			OutlineLevel: &OutlineLevel{Val: "4"},
		},
		StyleRPr: &StyleRPr{
			Bold:   &Bold{},
			BoldCs: &Bold{},
			Size:   &Size{Val: "20"}, // 10pt
			SizeCs: &Size{Val: "20"},
		},
	}
}

// NewDefaultStyles
func NewDefaultStyles() *Styles {
	styles := Styles{
		XmlnsW: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		XmlnsR: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
		Styles: []Style{
			// Normal style
			normalStyle(),
			// Heading 1
			heading1Style(),
			// Heading 2
			heading2Style(),
			// Heading 3
			heading3Style(),
			// Heading 4
			heading4Style(),
			// Heading 5
			heading5Style(),
		},
	}
	return &styles
}

// GetStyles
func (s *Styles) Get() *Styles {
	return s
}
