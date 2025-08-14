// writer/styles.go
package writer

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"strconv"

	"github.com/didikprabowo/mbadocx/properties"
	"github.com/didikprabowo/mbadocx/styles"
	"github.com/didikprabowo/mbadocx/types"
)

var _ zipWritable = (*Styles)(nil)

type Styles struct {
	document types.Document
}

// StylesXML represents the root styles XML structure
type StylesXML struct {
	XMLName      xml.Name         `xml:"w:styles"`
	XmlnsW       string           `xml:"xmlns:w,attr"`
	DocDefaults  *DocDefaultsXML  `xml:"w:docDefaults"`
	LatentStyles *LatentStylesXML `xml:"w:latentStyles,omitempty"`
	Styles       []StyleXML       `xml:"w:style"`
}

// DocDefaultsXML represents document defaults
type DocDefaultsXML struct {
	RPrDefault *RPrDefaultXML `xml:"w:rPrDefault,omitempty"`
	PPrDefault *PPrDefaultXML `xml:"w:pPrDefault,omitempty"`
}

type RPrDefaultXML struct {
	RPr *RunPropertiesXML `xml:"w:rPr"`
}

type PPrDefaultXML struct {
	PPr *ParagraphPropertiesXML `xml:"w:pPr"`
}

// LatentStylesXML represents latent styles
type LatentStylesXML struct {
	DefLockedState    string `xml:"w:defLockedState,attr"`
	DefUIPriority     string `xml:"w:defUIPriority,attr"`
	DefSemiHidden     string `xml:"w:defSemiHidden,attr"`
	DefUnhideWhenUsed string `xml:"w:defUnhideWhenUsed,attr"`
	DefQFormat        string `xml:"w:defQFormat,attr"`
	Count             string `xml:"w:count,attr"`
}

// StyleXML represents individual style
type StyleXML struct {
	Type    string                  `xml:"w:type,attr"`
	StyleId string                  `xml:"w:styleId,attr"`
	Default string                  `xml:"w:default,attr,omitempty"`
	Name    *NameXML                `xml:"w:name"`
	BasedOn *BasedOnXML             `xml:"w:basedOn,omitempty"`
	Next    *NextXML                `xml:"w:next,omitempty"`
	QFormat *QFormatXML             `xml:"w:qFormat,omitempty"`
	PPr     *ParagraphPropertiesXML `xml:"w:pPr,omitempty"`
	RPr     *RunPropertiesXML       `xml:"w:rPr,omitempty"`
}

type NameXML struct {
	Val string `xml:"w:val,attr"`
}

type BasedOnXML struct {
	Val string `xml:"w:val,attr"`
}

type NextXML struct {
	Val string `xml:"w:val,attr"`
}

type QFormatXML struct{}

// ParagraphPropertiesXML represents paragraph properties
type ParagraphPropertiesXML struct {
	KeepNext     *KeepNextXML     `xml:"w:keepNext,omitempty"`
	KeepLines    *KeepLinesXML    `xml:"w:keepLines,omitempty"`
	PageBreak    *PageBreakXML    `xml:"w:pageBreakBefore,omitempty"`
	WidowControl *WidowControlXML `xml:"w:widowControl,omitempty"`
	Spacing      *SpacingXML      `xml:"w:spacing,omitempty"`
	Ind          *IndXML          `xml:"w:ind,omitempty"`
	Jc           *JcXML           `xml:"w:jc,omitempty"`
	OutlineLvl   *OutlineLvlXML   `xml:"w:outlineLvl,omitempty"`
	NumPr        *NumPrXML        `xml:"w:numPr,omitempty"`
	PBdr         *PBdrXML         `xml:"w:pBdr,omitempty"`
	Shd          *ShdXML          `xml:"w:shd,omitempty"`
	Tabs         *TabsXML         `xml:"w:tabs,omitempty"`
}

type KeepNextXML struct{}
type KeepLinesXML struct{}
type PageBreakXML struct{}
type WidowControlXML struct{}

type SpacingXML struct {
	Before   string `xml:"w:before,attr,omitempty"`
	After    string `xml:"w:after,attr,omitempty"`
	Line     string `xml:"w:line,attr,omitempty"`
	LineRule string `xml:"w:lineRule,attr,omitempty"`
}

type IndXML struct {
	Left      string `xml:"w:left,attr,omitempty"`
	Right     string `xml:"w:right,attr,omitempty"`
	FirstLine string `xml:"w:firstLine,attr,omitempty"`
	Hanging   string `xml:"w:hanging,attr,omitempty"`
}

type JcXML struct {
	Val string `xml:"w:val,attr"`
}

type OutlineLvlXML struct {
	Val string `xml:"w:val,attr"`
}

type NumPrXML struct {
	NumId *NumIdXML `xml:"w:numId,omitempty"`
	ILvl  *ILvlXML  `xml:"w:ilvl,omitempty"`
}

type NumIdXML struct {
	Val string `xml:"w:val,attr"`
}

type ILvlXML struct {
	Val string `xml:"w:val,attr"`
}

type PBdrXML struct {
	Top    *BorderXML `xml:"w:top,omitempty"`
	Bottom *BorderXML `xml:"w:bottom,omitempty"`
	Left   *BorderXML `xml:"w:left,omitempty"`
	Right  *BorderXML `xml:"w:right,omitempty"`
}

type BorderXML struct {
	Val   string `xml:"w:val,attr"`
	Sz    string `xml:"w:sz,attr,omitempty"`
	Color string `xml:"w:color,attr,omitempty"`
	Space string `xml:"w:space,attr,omitempty"`
}

type ShdXML struct {
	Val   string `xml:"w:val,attr"`
	Fill  string `xml:"w:fill,attr,omitempty"`
	Color string `xml:"w:color,attr,omitempty"`
}

type TabsXML struct {
	Tab []TabXML `xml:"w:tab"`
}

type TabXML struct {
	Val    string `xml:"w:val,attr"`
	Pos    string `xml:"w:pos,attr"`
	Leader string `xml:"w:leader,attr,omitempty"`
}

// RunPropertiesXML represents run properties
type RunPropertiesXML struct {
	RFonts    *RFontsXML     `xml:"w:rFonts,omitempty"`
	Bold      *BoldXML       `xml:"w:b,omitempty"`
	BoldCs    *BoldCsXML     `xml:"w:bCs,omitempty"`
	Italic    *ItalicXML     `xml:"w:i,omitempty"`
	ItalicCs  *ItalicCsXML   `xml:"w:iCs,omitempty"`
	Underline *UnderlineXML  `xml:"w:u,omitempty"`
	Strike    *StrikeXML     `xml:"w:strike,omitempty"`
	DStrike   *DStrikeXML    `xml:"w:dstrike,omitempty"`
	SmallCaps *SmallCapsXML  `xml:"w:smallCaps,omitempty"`
	Caps      *CapsXML       `xml:"w:caps,omitempty"`
	Vanish    *VanishXML     `xml:"w:vanish,omitempty"`
	Emboss    *EmbossXML     `xml:"w:emboss,omitempty"`
	Imprint   *ImprintXML    `xml:"w:imprint,omitempty"`
	Shadow    *ShadowXML     `xml:"w:shadow,omitempty"`
	Outline   *OutlineXML    `xml:"w:outline,omitempty"`
	Sz        *SzXML         `xml:"w:sz,omitempty"`
	SzCs      *SzCsXML       `xml:"w:szCs,omitempty"`
	Color     *ColorXML      `xml:"w:color,omitempty"`
	Highlight *HighlightXML  `xml:"w:highlight,omitempty"`
	VertAlign *VertAlignXML  `xml:"w:vertAlign,omitempty"`
	Position  *PositionXML   `xml:"w:position,omitempty"`
	Spacing   *SpacingRunXML `xml:"w:spacing,omitempty"`
	Kern      *KernXML       `xml:"w:kern,omitempty"`
	Lang      *LangXML       `xml:"w:lang,omitempty"`
}

type RFontsXML struct {
	Ascii    string `xml:"w:ascii,attr,omitempty"`
	HAnsi    string `xml:"w:hAnsi,attr,omitempty"`
	EastAsia string `xml:"w:eastAsia,attr,omitempty"`
	Cs       string `xml:"w:cs,attr,omitempty"`
}

type BoldXML struct{}
type BoldCsXML struct{}
type ItalicXML struct{}
type ItalicCsXML struct{}

type UnderlineXML struct {
	Val   string `xml:"w:val,attr"`
	Color string `xml:"w:color,attr,omitempty"`
}

type StrikeXML struct{}
type DStrikeXML struct{}
type SmallCapsXML struct{}
type CapsXML struct{}
type VanishXML struct{}
type EmbossXML struct{}
type ImprintXML struct{}
type ShadowXML struct{}
type OutlineXML struct{}

type SzXML struct {
	Val string `xml:"w:val,attr"`
}

type SzCsXML struct {
	Val string `xml:"w:val,attr"`
}

type ColorXML struct {
	Val string `xml:"w:val,attr"`
}

type HighlightXML struct {
	Val string `xml:"w:val,attr"`
}

type VertAlignXML struct {
	Val string `xml:"w:val,attr"`
}

type PositionXML struct {
	Val string `xml:"w:val,attr"`
}

type SpacingRunXML struct {
	Val string `xml:"w:val,attr"`
}

type KernXML struct {
	Val string `xml:"w:val,attr"`
}

type LangXML struct {
	Val      string `xml:"w:val,attr,omitempty"`
	EastAsia string `xml:"w:eastAsia,attr,omitempty"`
	Bidi     string `xml:"w:bidi,attr,omitempty"`
}

func newStyles(document types.Document) *Styles {
	return &Styles{document: document}
}

// Path returns the location of the part inside the DOCX ZIP.
func (s *Styles) Path() string {
	return "word/styles.xml"
}

func (s *Styles) Byte() ([]byte, error) {
	// Get document styles - you'll need to add GetStyles() method to your document interface
	var docStyles *styles.DocumentStyles

	// Try to get styles from document, fallback to defaults
	if stylesGetter, ok := s.document.(interface{ GetStyles() *styles.DocumentStyles }); ok {
		docStyles = stylesGetter.GetStyles()
	}

	if docStyles == nil {
		docStyles = styles.DefaultDocumentStyles()
	}

	// Get document settings for font defaults
	docSettings := s.document.GetSettings()
	defaultFont := "Calibri"
	defaultSize := "22" // 11pt in half-points
	if docSettings != nil && docSettings.Font != nil {
		if docSettings.Font.DefaultFont != "" {
			defaultFont = docSettings.Font.DefaultFont
		}
		if docSettings.Font.DefaultSize > 0 {
			defaultSize = strconv.Itoa(docSettings.Font.DefaultSize)
		}
	}

	// Create styles XML structure
	stylesXML := &StylesXML{
		XmlnsW: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
		DocDefaults: &DocDefaultsXML{
			RPrDefault: &RPrDefaultXML{
				RPr: &RunPropertiesXML{
					RFonts: &RFontsXML{
						Ascii:    defaultFont,
						HAnsi:    defaultFont,
						EastAsia: defaultFont,
						Cs:       defaultFont,
					},
					Sz:   &SzXML{Val: defaultSize},
					SzCs: &SzCsXML{Val: defaultSize},
					Lang: &LangXML{
						Val:      "en-US",
						EastAsia: "en-US",
						Bidi:     "ar-SA",
					},
				},
			},
		},
		LatentStyles: &LatentStylesXML{
			DefLockedState:    "0",
			DefUIPriority:     "99",
			DefSemiHidden:     "0",
			DefUnhideWhenUsed: "0",
			DefQFormat:        "0",
			Count:             "376",
		},
	}

	// Convert styles to XML
	stylesXML.Styles = s.convertStylesToXML(docStyles)

	var buf bytes.Buffer

	// Write XML declaration
	buf.WriteString(xml.Header)

	// Encode the struct
	enc := xml.NewEncoder(&buf)
	enc.Indent("", "  ")

	if err := enc.Encode(stylesXML); err != nil {
		return nil, fmt.Errorf("encoding Styles XML: %w", err)
	}

	log.Printf("'%s' has been created.\n", s.Path())
	// log.Print(buf.String())

	return buf.Bytes(), nil
}

// convertStylesToXML converts DocumentStyles to XML format
func (s *Styles) convertStylesToXML(docStyles *styles.DocumentStyles) []StyleXML {
	var stylesList []StyleXML

	// Get all paragraph styles
	allStyles := docStyles.GetAllStyles()

	for _, style := range allStyles {
		if style == nil {
			continue
		}

		xmlStyle := StyleXML{
			Type:    style.Type,
			StyleId: style.StyleId,
			Name:    &NameXML{Val: style.Name},
		}

		// Set default flag
		if style.IsDefault {
			xmlStyle.Default = "1"
		}

		// Set based on
		if style.BasedOn != "" {
			xmlStyle.BasedOn = &BasedOnXML{Val: style.BasedOn}
		}

		// Set next style
		if style.NextStyle != "" {
			xmlStyle.Next = &NextXML{Val: style.NextStyle}
		}

		// Set qFormat
		if style.IsQFormat {
			xmlStyle.QFormat = &QFormatXML{}
		}

		// Convert paragraph properties
		if style.Paragraph != nil {
			xmlStyle.PPr = s.convertParagraphProperties(style.Paragraph, style.OutlineLevel)
		}

		// Convert run properties
		if style.Run != nil {
			xmlStyle.RPr = s.convertRunProperties(style.Run)
		}

		stylesList = append(stylesList, xmlStyle)
	}

	// Add character styles
	charStyles := []*styles.CharacterStyle{
		docStyles.DefaultCharacter,
		docStyles.Emphasis,
		docStyles.Strong,
		docStyles.Hyperlink,
	}

	for _, charStyle := range charStyles {
		if charStyle == nil {
			continue
		}

		xmlStyle := StyleXML{
			Type:    "character",
			StyleId: charStyle.StyleId,
			Name:    &NameXML{Val: charStyle.Name},
		}

		if charStyle.IsDefault {
			xmlStyle.Default = "1"
		}

		if charStyle.BasedOn != "" {
			xmlStyle.BasedOn = &BasedOnXML{Val: charStyle.BasedOn}
		}

		if charStyle.Run != nil {
			xmlStyle.RPr = s.convertRunProperties(charStyle.Run)
		}

		stylesList = append(stylesList, xmlStyle)
	}

	return stylesList
}

// convertParagraphProperties converts paragraph properties to XML
func (s *Styles) convertParagraphProperties(props *properties.ParagraphProperties, outlineLevel int) *ParagraphPropertiesXML {
	pPr := &ParagraphPropertiesXML{}

	// Spacing
	if props.SpacingBefore > 0 || props.SpacingAfter > 0 || props.LineSpacing > 0 {
		spacing := &SpacingXML{}

		if props.SpacingBefore > 0 {
			spacing.Before = fmt.Sprintf("%d", int(props.SpacingBefore))
		}
		if props.SpacingAfter > 0 {
			spacing.After = fmt.Sprintf("%d", int(props.SpacingAfter))
		}
		if props.LineSpacing > 0 {
			spacing.Line = fmt.Sprintf("%d", int(props.LineSpacing))
			if props.LineSpacingRule != "" {
				spacing.LineRule = props.LineSpacingRule
			}
		}

		pPr.Spacing = spacing
	}

	// Indentation
	if props.IndentLeft > 0 || props.IndentRight > 0 || props.IndentFirstLine != 0 {
		ind := &IndXML{}

		if props.IndentLeft > 0 {
			ind.Left = fmt.Sprintf("%d", int(props.IndentLeft))
		}
		if props.IndentRight > 0 {
			ind.Right = fmt.Sprintf("%d", int(props.IndentRight))
		}
		if props.IndentFirstLine > 0 {
			ind.FirstLine = fmt.Sprintf("%d", int(props.IndentFirstLine))
		} else if props.IndentFirstLine < 0 {
			ind.Hanging = fmt.Sprintf("%d", -int(props.IndentFirstLine))
		}

		pPr.Ind = ind
	}

	// Alignment
	if props.Alignment != "" && props.Alignment != "left" {
		pPr.Jc = &JcXML{Val: props.Alignment}
	}

	// Keep behaviors
	if props.KeepNext {
		pPr.KeepNext = &KeepNextXML{}
	}
	if props.KeepLines {
		pPr.KeepLines = &KeepLinesXML{}
	}
	if props.PageBreakBefore {
		pPr.PageBreak = &PageBreakXML{}
	}
	if props.WidowControl {
		pPr.WidowControl = &WidowControlXML{}
	}

	// Outline level
	if outlineLevel >= 0 && outlineLevel <= 9 {
		pPr.OutlineLvl = &OutlineLvlXML{Val: strconv.Itoa(outlineLevel)}
	}

	// Numbering
	if props.NumberingID != "" {
		pPr.NumPr = &NumPrXML{
			NumId: &NumIdXML{Val: props.NumberingID},
			ILvl:  &ILvlXML{Val: strconv.Itoa(props.NumberingLevel)},
		}
	}

	// Borders
	if props.Borders != nil {
		pPr.PBdr = s.convertBorders(props.Borders)
	}

	// Shading
	if props.Shading != nil {
		pPr.Shd = &ShdXML{
			Val:   props.Shading.Pattern,
			Fill:  props.Shading.Fill,
			Color: props.Shading.Color,
		}
	}

	// Tabs
	if len(props.Tabs) > 0 {
		tabs := &TabsXML{}
		for _, tab := range props.Tabs {
			tabs.Tab = append(tabs.Tab, TabXML{
				Val:    tab.Alignment,
				Pos:    strconv.Itoa(tab.Position),
				Leader: tab.Leader,
			})
		}
		pPr.Tabs = tabs
	}

	return pPr
}

// convertRunProperties converts run properties to XML
func (s *Styles) convertRunProperties(props *properties.RunProperties) *RunPropertiesXML {
	rPr := &RunPropertiesXML{}

	// Fonts
	if props.FontFamily != "" {
		fonts := &RFontsXML{}
		fonts.Ascii = props.FontFamily
		rPr.RFonts = fonts
	}

	// Font size
	if props.FontSize > 0 {
		fontSizeStr := fmt.Sprintf("%d", int(props.FontSize))
		rPr.Sz = &SzXML{Val: fontSizeStr}
		// rPr.SzCs = &SzCsXML{Val: fontSizeStr}

	}

	// Font styles
	if props.Bold != nil {
		rPr.Bold = &BoldXML{}
	}

	if props.Italic != nil {
		rPr.Italic = &ItalicXML{}
	}

	// Underline
	if props.Underline != "" && props.Underline != "none" {
		rPr.Underline = &UnderlineXML{Val: props.Underline}
	}

	// Strike through
	if props.Strike != nil && *props.Strike {
		rPr.Strike = &StrikeXML{}
	}
	if props.DoubleStrike != nil && *props.DoubleStrike {
		rPr.DStrike = &DStrikeXML{}
	}

	// Font effects
	if props.SmallCaps != nil && *props.SmallCaps {
		rPr.SmallCaps = &SmallCapsXML{}
	}
	if props.AllCaps != nil && *props.AllCaps {
		rPr.Caps = &CapsXML{}
	}
	if props.Hidden != nil && *props.Hidden {
		rPr.Vanish = &VanishXML{}
	}
	if props.Emboss != nil && *props.Emboss {
		rPr.Emboss = &EmbossXML{}
	}
	if props.Imprint != nil && *props.Imprint {
		rPr.Imprint = &ImprintXML{}
	}
	if props.Shadow != nil && *props.Shadow {
		rPr.Shadow = &ShadowXML{}
	}
	if props.Outline != nil && *props.Outline {
		rPr.Outline = &OutlineXML{}
	}

	// Color
	if props.Color != "" {
		rPr.Color = &ColorXML{Val: props.Color}
	}

	// Highlight
	if props.Highlight != "" {
		rPr.Highlight = &HighlightXML{Val: props.Highlight}
	}

	// Vertical alignment
	if props.VerticalAlign != "" {
		rPr.VertAlign = &VertAlignXML{Val: props.VerticalAlign}
	}

	// Position
	if props.Position != 0 {
		rPr.Position = &PositionXML{Val: strconv.Itoa(props.Position)}
	}

	// Spacing
	if props.Spacing != 0 {
		rPr.Spacing = &SpacingRunXML{Val: strconv.Itoa(props.Spacing)}
	}

	// Kerning
	if props.Kerning > 0 {
		rPr.Kern = &KernXML{Val: fmt.Sprintf("%d", int(props.Kerning))}
	}

	// Language
	if props.Language != "" {
		lang := &LangXML{}
		if props.Language != "" {
			lang.Val = props.Language
		}
		rPr.Lang = lang
	}

	return rPr
}

// convertBorders converts border properties to XML
func (s *Styles) convertBorders(borders *properties.ParagraphBorders) *PBdrXML {
	pBdr := &PBdrXML{}

	if borders.Top != nil {
		pBdr.Top = &BorderXML{
			Val:   borders.Top.Type,
			Sz:    strconv.Itoa(borders.Top.Width),
			Color: borders.Top.Color,
			Space: strconv.Itoa(borders.Top.Space),
		}
	}

	if borders.Bottom != nil {
		pBdr.Bottom = &BorderXML{
			Val:   borders.Bottom.Type,
			Sz:    strconv.Itoa(borders.Bottom.Width),
			Color: borders.Bottom.Color,
			Space: strconv.Itoa(borders.Bottom.Space),
		}
	}

	if borders.Left != nil {
		pBdr.Left = &BorderXML{
			Val:   borders.Left.Type,
			Sz:    strconv.Itoa(borders.Left.Width),
			Color: borders.Left.Color,
			Space: strconv.Itoa(borders.Left.Space),
		}
	}

	if borders.Right != nil {
		pBdr.Right = &BorderXML{
			Val:   borders.Right.Type,
			Sz:    strconv.Itoa(borders.Right.Width),
			Color: borders.Right.Color,
			Space: strconv.Itoa(borders.Right.Space),
		}
	}

	return pBdr
}

// WriteTo writes the XML content to the given writer.
func (s *Styles) WriteTo(w io.Writer) (int64, error) {
	data, err := s.Byte()
	if err != nil {
		return 0, err
	}

	n, err := w.Write(data)
	return int64(n), err
}
