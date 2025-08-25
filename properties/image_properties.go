package properties

import (
	"fmt"
	"strings"
)

// ImageProperties contains all configurable properties for an image
type ImageProperties struct {
	// Positioning
	Inline   bool      // true for inline, false for floating
	WrapType WrapStyle // inline, square, tight, through, topAndBottom

	// Alignment (for inline images)
	Alignment ImageAlignment // left, center, right, justify

	// Floating position (when not inline)
	HorizontalPosition HorizontalAnchor
	VerticalPosition   VerticalAnchor
	HorizontalOffset   int64 // Offset in EMUs
	VerticalOffset     int64 // Offset in EMUs

	// Borders and effects
	BorderWidth int    // Border width in points
	BorderColor string // Hex color (e.g., "FF0000" for red)
	BorderStyle string // solid, dashed, dotted, double

	// Shadow effects
	Shadow       bool
	ShadowColor  string // Hex color
	ShadowOffset int    // Offset in points

	// Rotation
	Rotation int // Degrees of rotation (0-360)

	// Cropping (percentage values 0-100)
	CropTop    float64
	CropBottom float64
	CropLeft   float64
	CropRight  float64

	// Image adjustments
	Brightness float64 // -100 to 100 (0 is normal)
	Contrast   float64 // -100 to 100 (0 is normal)
	Saturation float64 // 0 to 200 (100 is normal)

	// Transparency
	Transparency float64 // 0 to 100 (0 is opaque)

	// Alternative text for accessibility
	AltText string

	// Lock aspect ratio
	LockAspectRatio bool

	// Allow overlap with other objects
	AllowOverlap bool

	// Z-order (layering)
	ZOrder int
}

// WrapStyle defines how text wraps around the image
type WrapStyle string

const (
	WrapInline       WrapStyle = "inline"
	WrapSquare       WrapStyle = "square"
	WrapTight        WrapStyle = "tight"
	WrapThrough      WrapStyle = "through"
	WrapTopAndBottom WrapStyle = "topAndBottom"
	WrapBehindText   WrapStyle = "behindText"
	WrapInFrontText  WrapStyle = "inFrontOfText"
)

// ImageAlignment defines the horizontal alignment of inline images
type ImageAlignment string

const (
	AlignLeft    ImageAlignment = "left"
	AlignCenter  ImageAlignment = "center"
	AlignRight   ImageAlignment = "right"
	AlignJustify ImageAlignment = "justify"
)

// HorizontalAnchor defines the horizontal anchor point for floating images
type HorizontalAnchor string

const (
	HorizontalAnchorMargin        HorizontalAnchor = "margin"
	HorizontalAnchorPage          HorizontalAnchor = "page"
	HorizontalAnchorColumn        HorizontalAnchor = "column"
	HorizontalAnchorCharacter     HorizontalAnchor = "character"
	HorizontalAnchorLeftMargin    HorizontalAnchor = "leftMargin"
	HorizontalAnchorRightMargin   HorizontalAnchor = "rightMargin"
	HorizontalAnchorInsideMargin  HorizontalAnchor = "insideMargin"
	HorizontalAnchorOutsideMargin HorizontalAnchor = "outsideMargin"
)

// VerticalAnchor defines the vertical anchor point for floating images
type VerticalAnchor string

const (
	VerticalAnchorMargin        VerticalAnchor = "margin"
	VerticalAnchorPage          VerticalAnchor = "page"
	VerticalAnchorParagraph     VerticalAnchor = "paragraph"
	VerticalAnchorLine          VerticalAnchor = "line"
	VerticalAnchorTopMargin     VerticalAnchor = "topMargin"
	VerticalAnchorBottomMargin  VerticalAnchor = "bottomMargin"
	VerticalAnchorInsideMargin  VerticalAnchor = "insideMargin"
	VerticalAnchorOutsideMargin VerticalAnchor = "outsideMargin"
)

// NewImageProperties creates a new ImageProperties with default values
func NewImageProperties() *ImageProperties {
	return &ImageProperties{
		Inline:          true,
		WrapType:        WrapInline,
		Alignment:       AlignLeft,
		LockAspectRatio: true,
		AllowOverlap:    true,
		Brightness:      0,
		Contrast:        0,
		Saturation:      100,
		Transparency:    0,
	}
}

// SetBorder sets the border properties
func (p *ImageProperties) SetBorder(width int, color string, style string) *ImageProperties {
	p.BorderWidth = width
	p.BorderColor = strings.TrimPrefix(color, "#")
	p.BorderStyle = style
	return p
}

// SetShadow sets the shadow properties
func (p *ImageProperties) SetShadow(enabled bool, color string, offset int) *ImageProperties {
	p.Shadow = enabled
	if enabled {
		p.ShadowColor = strings.TrimPrefix(color, "#")
		p.ShadowOffset = offset
	}
	return p
}

// SetCrop sets the cropping values (in percentages)
func (p *ImageProperties) SetCrop(top, bottom, left, right float64) *ImageProperties {
	p.CropTop = clamp(top, 0, 100)
	p.CropBottom = clamp(bottom, 0, 100)
	p.CropLeft = clamp(left, 0, 100)
	p.CropRight = clamp(right, 0, 100)
	return p
}

// SetImageAdjustments sets brightness, contrast, and saturation
func (p *ImageProperties) SetImageAdjustments(brightness, contrast, saturation float64) *ImageProperties {
	p.Brightness = clamp(brightness, -100, 100)
	p.Contrast = clamp(contrast, -100, 100)
	p.Saturation = clamp(saturation, 0, 200)
	return p
}

// SetFloatingPosition sets the position for floating images
func (p *ImageProperties) SetFloatingPosition(hAnchor HorizontalAnchor, vAnchor VerticalAnchor, hOffset, vOffset int64) *ImageProperties {
	p.Inline = false
	p.HorizontalPosition = hAnchor
	p.VerticalPosition = vAnchor
	p.HorizontalOffset = hOffset
	p.VerticalOffset = vOffset
	return p
}

// GeneratePositionXML generates the XML for image positioning
func (p *ImageProperties) GeneratePositionXML() string {
	var xml strings.Builder

	if !p.Inline {
		// Floating image positioning
		xml.WriteString(fmt.Sprintf(`<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="%d" behindDoc="%s" locked="0" layoutInCell="1" allowOverlap="%s">`,
			p.ZOrder,
			boolToString(p.WrapType == WrapBehindText),
			boolToString(p.AllowOverlap)))

		// Position settings
		xml.WriteString(`<wp:simplePos x="0" y="0"/>`)
		xml.WriteString(fmt.Sprintf(`<wp:positionH relativeFrom="%s">`, p.HorizontalPosition))
		xml.WriteString(fmt.Sprintf(`<wp:posOffset>%d</wp:posOffset>`, p.HorizontalOffset))
		xml.WriteString(`</wp:positionH>`)
		xml.WriteString(fmt.Sprintf(`<wp:positionV relativeFrom="%s">`, p.VerticalPosition))
		xml.WriteString(fmt.Sprintf(`<wp:posOffset>%d</wp:posOffset>`, p.VerticalOffset))
		xml.WriteString(`</wp:positionV>`)
	}

	return xml.String()
}

// GenerateWrapXML generates the XML for text wrapping
func (p *ImageProperties) GenerateWrapXML() string {
	switch p.WrapType {
	case WrapSquare:
		return `<wp:wrapSquare wrapText="bothSides"/>`
	case WrapTight:
		return `<wp:wrapTight wrapText="bothSides"/>`
	case WrapThrough:
		return `<wp:wrapThrough wrapText="bothSides"/>`
	case WrapTopAndBottom:
		return `<wp:wrapTopAndBottom/>`
	case WrapBehindText:
		return `<wp:wrapNone/>`
	case WrapInFrontText:
		return `<wp:wrapNone/>`
	default:
		return ""
	}
}

// GenerateEffectsXML generates the XML for image effects
func (p *ImageProperties) GenerateEffectsXML() string {
	var xml strings.Builder

	// Border
	if p.BorderWidth > 0 && p.BorderColor != "" {
		xml.WriteString(`<a:ln w="`)
		xml.WriteString(fmt.Sprintf("%d", p.BorderWidth*12700)) // Convert points to EMUs
		xml.WriteString(`">`)
		xml.WriteString(`<a:solidFill>`)
		xml.WriteString(fmt.Sprintf(`<a:srgbClr val="%s"/>`, p.BorderColor))
		xml.WriteString(`</a:solidFill>`)

		// Border style
		switch p.BorderStyle {
		case "dashed":
			xml.WriteString(`<a:prstDash val="dash"/>`)
		case "dotted":
			xml.WriteString(`<a:prstDash val="dot"/>`)
		case "double":
			xml.WriteString(`<a:prstDash val="lgDash"/>`)
		default:
			xml.WriteString(`<a:prstDash val="solid"/>`)
		}

		xml.WriteString(`</a:ln>`)
	}

	// Shadow
	if p.Shadow {
		xml.WriteString(`<a:effectLst>`)
		xml.WriteString(`<a:outerShdw blurRad="50800" dist="`)
		xml.WriteString(fmt.Sprintf("%d", p.ShadowOffset*12700))
		xml.WriteString(`" dir="2700000" algn="tl">`)
		xml.WriteString(fmt.Sprintf(`<a:srgbClr val="%s">`, p.ShadowColor))
		xml.WriteString(`<a:alpha val="50000"/>`)
		xml.WriteString(`</a:srgbClr>`)
		xml.WriteString(`</a:outerShdw>`)
		xml.WriteString(`</a:effectLst>`)
	}

	return xml.String()
}

// GenerateTransformXML generates the XML for image transformations
func (p *ImageProperties) GenerateTransformXML(width, height int64) string {
	var xml strings.Builder

	xml.WriteString(`<a:xfrm`)
	if p.Rotation != 0 {
		// Convert degrees to 60000ths of a degree
		xml.WriteString(fmt.Sprintf(` rot="%d"`, p.Rotation*60000))
	}
	xml.WriteString(`>`)
	xml.WriteString(`<a:off x="0" y="0"/>`)
	xml.WriteString(fmt.Sprintf(`<a:ext cx="%d" cy="%d"/>`, width, height))
	xml.WriteString(`</a:xfrm>`)

	return xml.String()
}

// GenerateCropXML generates the XML for image cropping
func (p *ImageProperties) GenerateCropXML() string {
	if p.CropTop == 0 && p.CropBottom == 0 && p.CropLeft == 0 && p.CropRight == 0 {
		return ""
	}

	// Convert percentages to per-thousandths
	top := int(p.CropTop * 1000)
	bottom := int(p.CropBottom * 1000)
	left := int(p.CropLeft * 1000)
	right := int(p.CropRight * 1000)

	return fmt.Sprintf(`<a:srcRect l="%d" t="%d" r="%d" b="%d"/>`, left, top, right, bottom)
}

// GenerateImageAdjustmentsXML generates the XML for image adjustments
func (p *ImageProperties) GenerateImageAdjustmentsXML() string {
	var xml strings.Builder

	if p.Brightness != 0 || p.Contrast != 0 || p.Saturation != 100 {
		xml.WriteString(`<a:extLst>`)
		xml.WriteString(`<a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}">`)
		xml.WriteString(`<a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">`)
		xml.WriteString(`<a14:imgLayer r:embed="rId1">`)
		xml.WriteString(`<a14:imgEffect>`)

		// Brightness
		if p.Brightness != 0 {
			bright := int((p.Brightness + 100) / 200 * 100000) // Convert to 0-100000 scale
			xml.WriteString(fmt.Sprintf(`<a14:brightnessContrast bright="%d"/>`, bright))
		}

		// Contrast
		if p.Contrast != 0 {
			contrast := int((p.Contrast + 100) / 200 * 100000) // Convert to 0-100000 scale
			xml.WriteString(fmt.Sprintf(`<a14:brightnessContrast contrast="%d"/>`, contrast))
		}

		// Saturation
		if p.Saturation != 100 {
			sat := int(p.Saturation * 1000) // Convert to per-thousandths
			xml.WriteString(fmt.Sprintf(`<a14:saturation sat="%d"/>`, sat))
		}

		xml.WriteString(`</a14:imgEffect>`)
		xml.WriteString(`</a14:imgLayer>`)
		xml.WriteString(`</a14:imgProps>`)
		xml.WriteString(`</a:ext>`)
		xml.WriteString(`</a:extLst>`)
	}

	// Transparency
	if p.Transparency > 0 {
		alpha := int((100 - p.Transparency) * 1000) // Convert to alpha value
		xml.WriteString(fmt.Sprintf(`<a:alphaModFix amt="%d"/>`, alpha))
	}

	return xml.String()
}

// Validate checks if the properties are valid
func (p *ImageProperties) Validate() error {
	if p.BorderWidth < 0 {
		return fmt.Errorf("border width cannot be negative")
	}

	if p.Rotation < 0 || p.Rotation > 360 {
		return fmt.Errorf("rotation must be between 0 and 360 degrees")
	}

	if p.CropTop < 0 || p.CropTop > 100 ||
		p.CropBottom < 0 || p.CropBottom > 100 ||
		p.CropLeft < 0 || p.CropLeft > 100 ||
		p.CropRight < 0 || p.CropRight > 100 {
		return fmt.Errorf("crop values must be between 0 and 100")
	}

	if p.CropTop+p.CropBottom >= 100 || p.CropLeft+p.CropRight >= 100 {
		return fmt.Errorf("total crop cannot exceed 100%%")
	}

	return nil
}

// Helper functions

func clamp(value, min, max float64) float64 {
	if value < min {
		return min
	}
	if value > max {
		return max
	}
	return value
}

func boolToString(b bool) string {
	if b {
		return "1"
	}
	return "0"
}
