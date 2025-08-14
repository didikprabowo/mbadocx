// File: elements/run.go
package elements

import (
	"bytes"
	"fmt"
	"strings"

	"github.com/didikprabowo/mbadocx/properties"
)

// Run represents a run of text with consistent formatting
type Run struct {
	Properties *properties.RunProperties
	Children   []RunChild
}

// RunChild interface for elements that can be children of a run
type RunChild interface {
	Type() string
	XML() ([]byte, error)
}

// NewRun creates a new run
func NewRun() *Run {
	return &Run{
		Properties: properties.NewRunProperties(),
		Children:   make([]RunChild, 0),
	}
}

// Type returns the element type
func (r *Run) Type() string {
	return "run"
}

// AddText adds text to the run
func (r *Run) AddText(text string) *Run {
	t := NewText(text)
	r.Children = append(r.Children, t)
	return r
}

// AddBreak adds a line break
func (r *Run) AddBreak() *Run {
	r.Children = append(r.Children, NewLineBreak())
	return r
}

// AddPageBreak
func (r *Run) AddPageBreak() *Run {
	r.Children = append(r.Children, NewPageBreak())
	return r
}

// AddTab adds a tab character
func (r *Run) AddTab() *Run {
	r.Children = append(r.Children, NewTab())
	return r
}

// AddSpace adds exactly N space characters (preserved)
func (r *Run) AddSpace(count int) *Run {
	spaces := strings.Repeat(" ", count)
	t := &Text{
		Value:         spaces,
		PreserveSpace: true, // Must preserve!
	}
	r.Children = append(r.Children, t)
	return r
}

// SetBold sets the bold property
func (r *Run) SetBold(bold bool) *Run {
	r.Properties.Bold = &bold
	return r
}

// SetItalic sets the italic property
func (r *Run) SetItalic(italic bool) *Run {
	r.Properties.Italic = &italic
	return r
}

// SetUnderline sets the underline property
// Values: "single", "double", "thick", "dotted", "dash", "dotDash", "dotDotDash", "wave"
func (r *Run) SetUnderline(underline string) *Run {
	r.Properties.Underline = underline
	return r
}

// SetStrike sets the strikethrough property
func (r *Run) SetStrike(strike bool) *Run {
	r.Properties.Strike = &strike
	return r
}

// SetFontSize sets the font size in points
func (r *Run) SetFontSize(size float64) *Run {
	r.Properties.FontSize = size
	return r
}

// SetFontFamily sets the font family
func (r *Run) SetFontFamily(font string) *Run {
	r.Properties.FontFamily = font
	return r
}

// SetColor sets the text color (hex format, e.g., "FF0000" for red)
func (r *Run) SetColor(color string) *Run {
	r.Properties.Color = color
	return r
}

// SetHighlight sets the highlight color
// Values: "yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue",
//
//	"darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow",
//	"darkGray", "lightGray", "black", "none"
func (r *Run) SetHighlight(color string) *Run {
	r.Properties.Highlight = color
	return r
}

// SetVerticalAlign sets the vertical alignment
// Values: "baseline", "superscript", "subscript"
func (r *Run) SetVerticalAlign(align string) *Run {
	r.Properties.VerticalAlign = align
	return r
}

// SetSpacing sets the character spacing in twips (1/20th of a point)
func (r *Run) SetSpacing(spacing int) *Run {
	r.Properties.Spacing = spacing
	return r
}

// SetKerning sets the kerning in points
func (r *Run) SetKerning(kerning float64) *Run {
	r.Properties.Kerning = kerning
	return r
}

// SetStyle sets the character style
func (r *Run) SetStyle(styleID string) *Run {
	r.Properties.StyleID = styleID
	return r
}

// SetAllCaps sets the all caps property
func (r *Run) SetAllCaps(allCaps bool) *Run {
	r.Properties.AllCaps = &allCaps
	return r
}

// SetSmallCaps sets the small caps property
func (r *Run) SetSmallCaps(smallCaps bool) *Run {
	r.Properties.SmallCaps = &smallCaps
	return r
}

// SetDoubleStrike sets the double strikethrough property
func (r *Run) SetDoubleStrike(doubleStrike bool) *Run {
	r.Properties.DoubleStrike = &doubleStrike
	return r
}

// SetEmboss sets the emboss property
func (r *Run) SetEmboss(emboss bool) *Run {
	r.Properties.Emboss = &emboss
	return r
}

// SetImprint sets the imprint/engrave property
func (r *Run) SetImprint(imprint bool) *Run {
	r.Properties.Imprint = &imprint
	return r
}

// SetOutline sets the outline property
func (r *Run) SetOutline(outline bool) *Run {
	r.Properties.Outline = &outline
	return r
}

// SetShadow sets the shadow property
func (r *Run) SetShadow(shadow bool) *Run {
	r.Properties.Shadow = &shadow
	return r
}

// SetVanish sets the vanish/hidden property
func (r *Run) SetVanish(vanish bool) *Run {
	r.Properties.Vanish = &vanish
	return r
}

// Clone creates a deep copy of the run
func (r *Run) Clone() *Run {
	newRun := &Run{
		Properties: r.Properties.Clone(),
		Children:   make([]RunChild, 0, len(r.Children)),
	}

	// Clone children
	for _, child := range r.Children {
		switch c := child.(type) {
		case *Text:
			newRun.Children = append(newRun.Children, &Text{
				Value:         c.Value,
				PreserveSpace: c.PreserveSpace,
			})
		case *LineBreak:
			newRun.Children = append(newRun.Children, NewLineBreak())
		case *PageBreak:
			newRun.Children = append(newRun.Children, NewPageBreak())
		case *Tab:
			newRun.Children = append(newRun.Children, NewTab())
		}
	}

	return newRun
}

// HasFormatting returns true if the run has any formatting applied
func (r *Run) HasFormatting() bool {
	if r.Properties == nil {
		return false
	}

	p := r.Properties
	return (p.Bold != nil && *p.Bold) ||
		(p.Italic != nil && *p.Italic) ||
		p.Underline != "" ||
		(p.Strike != nil && *p.Strike) ||
		p.FontSize != 0 ||
		p.FontFamily != "" ||
		p.Color != "" ||
		p.Highlight != "" ||
		p.VerticalAlign != "" ||
		p.Spacing != 0 ||
		p.Kerning != 0 ||
		p.StyleID != ""
}

// XML generates the XML representation of the run
func (r *Run) XML() ([]byte, error) {
	var buf bytes.Buffer

	// Start run tag
	buf.WriteString(`<w:r>`)

	// Add properties if they exist
	if r.Properties != nil && r.HasFormatting() {
		propXML, err := r.generatePropertiesXML()
		if err != nil {
			return nil, fmt.Errorf("generating run properties XML: %w", err)
		}
		buf.Write(propXML)
	}

	// Add children (text, breaks, tabs)
	for _, child := range r.Children {
		childXML, err := child.XML()
		if err != nil {
			return nil, fmt.Errorf("generating child XML: %w", err)
		}
		buf.Write(childXML)
	}

	// Close run tag
	buf.WriteString(`</w:r>`)

	return buf.Bytes(), nil
}

// generatePropertiesXML generates the run properties XML
func (r *Run) generatePropertiesXML() ([]byte, error) {
	rp := r.Properties
	if rp == nil {
		return nil, nil
	}

	var buf bytes.Buffer
	buf.WriteString(`<w:rPr>`)

	// Character style
	if rp.StyleID != "" {
		buf.WriteString(fmt.Sprintf(`<w:rStyle w:val="%s"/>`, rp.StyleID))
	}

	// Font family
	if rp.FontFamily != "" {
		buf.WriteString(fmt.Sprintf(`<w:rFonts w:ascii="%s" w:hAnsi="%s" w:eastAsia="%s" w:cs="%s"/>`,
			rp.FontFamily, rp.FontFamily, rp.FontFamily, rp.FontFamily))
	}

	// Bold
	if rp.Bold != nil {
		if *rp.Bold {
			buf.WriteString(`<w:b/>`)
			buf.WriteString(`<w:bCs/>`) // Complex script bold
		} else {
			buf.WriteString(`<w:b w:val="false"/>`)
			buf.WriteString(`<w:bCs w:val="false"/>`)
		}
	}

	// Italic
	if rp.Italic != nil {
		if *rp.Italic {
			buf.WriteString(`<w:i/>`)
			buf.WriteString(`<w:iCs/>`) // Complex script italic
		} else {
			buf.WriteString(`<w:i w:val="false"/>`)
			buf.WriteString(`<w:iCs w:val="false"/>`)
		}
	}

	// All caps
	if rp.AllCaps != nil && *rp.AllCaps {
		buf.WriteString(`<w:caps/>`)
	}

	// Small caps
	if rp.SmallCaps != nil && *rp.SmallCaps {
		buf.WriteString(`<w:smallCaps/>`)
	}

	// Strike
	if rp.Strike != nil && *rp.Strike {
		buf.WriteString(`<w:strike/>`)
	}

	// Double strike
	if rp.DoubleStrike != nil && *rp.DoubleStrike {
		buf.WriteString(`<w:dstrike/>`)
	}

	// Outline
	if rp.Outline != nil && *rp.Outline {
		buf.WriteString(`<w:outline/>`)
	}

	// Shadow
	if rp.Shadow != nil && *rp.Shadow {
		buf.WriteString(`<w:shadow/>`)
	}

	// Emboss
	if rp.Emboss != nil && *rp.Emboss {
		buf.WriteString(`<w:emboss/>`)
	}

	// Imprint
	if rp.Imprint != nil && *rp.Imprint {
		buf.WriteString(`<w:imprint/>`)
	}

	// Font size
	if rp.FontSize > 0 {
		// Convert points to half-points
		halfPoints := int(rp.FontSize * 2)
		buf.WriteString(fmt.Sprintf(`<w:sz w:val="%d"/>`, halfPoints))
		buf.WriteString(fmt.Sprintf(`<w:szCs w:val="%d"/>`, halfPoints)) // Complex script size
	}

	// Underline
	if rp.Underline != "" && rp.Underline != "none" {
		buf.WriteString(fmt.Sprintf(`<w:u w:val="%s"/>`, rp.Underline))
	}

	// Vanish/hidden
	if rp.Vanish != nil && *rp.Vanish {
		buf.WriteString(`<w:vanish/>`)
	}

	// Color
	if rp.Color != "" {
		// Remove # if present
		color := rp.Color
		if strings.HasPrefix(color, "#") {
			color = color[1:]
		}
		buf.WriteString(fmt.Sprintf(`<w:color w:val="%s"/>`, color))
	}

	// Character spacing
	if rp.Spacing != 0 {
		buf.WriteString(fmt.Sprintf(`<w:spacing w:val="%d"/>`, rp.Spacing))
	}

	// Kerning
	if rp.Kerning > 0 {
		buf.WriteString(fmt.Sprintf(`<w:kern w:val="%d"/>`, int(rp.Kerning*2))) // Convert to half-points
	}

	// Highlight
	if rp.Highlight != "" && rp.Highlight != "none" {
		buf.WriteString(fmt.Sprintf(`<w:highlight w:val="%s"/>`, rp.Highlight))
	}

	// Vertical alignment
	if rp.VerticalAlign != "" && rp.VerticalAlign != "baseline" {
		buf.WriteString(fmt.Sprintf(`<w:vertAlign w:val="%s"/>`, rp.VerticalAlign))
	}

	buf.WriteString(`</w:rPr>`)

	return buf.Bytes(), nil
}

// Validate checks if the run is valid
func (r *Run) Validate() error {
	if r.Properties != nil {
		// Validate underline values
		validUnderlines := map[string]bool{
			"single": true, "double": true, "thick": true, "dotted": true,
			"dash": true, "dotDash": true, "dotDotDash": true, "wave": true,
			"none": true, "": true,
		}
		if !validUnderlines[r.Properties.Underline] {
			return fmt.Errorf("invalid underline value: %s", r.Properties.Underline)
		}

		// Validate vertical alignment
		validVertAlign := map[string]bool{
			"baseline": true, "superscript": true, "subscript": true, "": true,
		}
		if !validVertAlign[r.Properties.VerticalAlign] {
			return fmt.Errorf("invalid vertical alignment: %s", r.Properties.VerticalAlign)
		}

		// Validate highlight colors
		validHighlights := map[string]bool{
			"yellow": true, "green": true, "cyan": true, "magenta": true,
			"blue": true, "red": true, "darkBlue": true, "darkCyan": true,
			"darkGreen": true, "darkMagenta": true, "darkRed": true,
			"darkYellow": true, "darkGray": true, "lightGray": true,
			"black": true, "none": true, "": true,
		}
		if !validHighlights[r.Properties.Highlight] {
			return fmt.Errorf("invalid highlight color: %s", r.Properties.Highlight)
		}
	}

	return nil
}
