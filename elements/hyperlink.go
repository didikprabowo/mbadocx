// File: elements/hyperlink.go
package elements

import (
	"bytes"
	"fmt"
	"strings"

	"github.com/didikprabowo/mbadocx/properties"
	"github.com/google/uuid"
)

// Hyperlink represents a hyperlink element
type Hyperlink struct {
	ID          string                    // Relationship ID
	URL         string                    // Target URL or anchor
	Tooltip     string                    // Tooltip text
	Children    []ParagraphChild          // Child elements (usually runs)
	Properties  *properties.RunProperties // Default properties for children
	Typ         string                    // Type: "external" or "internal"
	Anchor      string                    // Anchor for internal links
	DocLocation string                    // Document location for internal links
	TargetFrame string                    // Target frame
	History     bool                      // Add to history
	ScreenTip   string                    // Extended tooltip
}

// HyperlinkType constants
const (
	HyperlinkTypeExternal = "external"
	HyperlinkTypeInternal = "internal"
	HyperlinkTypeEmail    = "email"
	HyperlinkTypeBookmark = "bookmark"
)

// NewHyperlink creates a new external hyperlink
func NewHyperlink(text, url string) *Hyperlink {
	h := &Hyperlink{
		ID:         generateRelationshipID(),
		URL:        url,
		Typ:        HyperlinkTypeExternal,
		History:    true,
		Properties: properties.NewRunProperties(),
		Children:   make([]ParagraphChild, 0),
	}

	// Set default hyperlink formatting
	h.Properties.Color = "0563C1"     // Blue
	h.Properties.Underline = "single" // Underlined

	// Add the text as a run
	if text != "" {
		r := NewRun()
		r.Properties = h.Properties.Clone()
		r.AddText(text)
		h.Children = append(h.Children, r)
	}

	return h
}

// NewInternalHyperlink creates a hyperlink to a bookmark
func NewInternalHyperlink(text, anchor string) *Hyperlink {
	h := &Hyperlink{
		ID:         "", // Internal links don't need relationship ID
		Anchor:     anchor,
		Typ:        HyperlinkTypeInternal,
		History:    true,
		Properties: properties.NewRunProperties(),
		Children:   make([]ParagraphChild, 0),
	}

	// Set default hyperlink formatting
	h.Properties.Color = "0563C1"
	h.Properties.Underline = "single"

	// Add the text as a run
	if text != "" {
		r := NewRun()
		r.Properties = h.Properties.Clone()
		r.AddText(text)
		h.Children = append(h.Children, r)
	}

	return h
}

// NewEmailHyperlink creates an email hyperlink
func NewEmailHyperlink(text, email string) *Hyperlink {
	url := email
	if !strings.HasPrefix(email, "mailto:") {
		url = "mailto:" + email
	}

	h := NewHyperlink(text, url)
	h.Typ = HyperlinkTypeEmail
	return h
}

// NewBookmarkHyperlink creates a hyperlink to a bookmark
func NewBookmarkHyperlink(text, bookmarkName string) *Hyperlink {
	h := NewInternalHyperlink(text, bookmarkName)
	h.Typ = HyperlinkTypeBookmark
	return h
}

// Type returns the element type
func (h *Hyperlink) Type() string {
	return "hyperlink"
}

// SetTooltip sets the hyperlink tooltip
func (h *Hyperlink) SetTooltip(tooltip string) *Hyperlink {
	h.Tooltip = tooltip
	h.ScreenTip = tooltip
	return h
}

// SetScreenTip sets extended tooltip text
func (h *Hyperlink) SetScreenTip(tip string) *Hyperlink {
	h.ScreenTip = tip
	return h
}

// SetTargetFrame sets the target frame
func (h *Hyperlink) SetTargetFrame(frame string) *Hyperlink {
	h.TargetFrame = frame
	return h
}

// SetHistory sets whether to add to history
func (h *Hyperlink) SetHistory(history bool) *Hyperlink {
	h.History = history
	return h
}

// SetDocLocation sets document location for internal links
func (h *Hyperlink) SetDocLocation(location string) *Hyperlink {
	h.DocLocation = location
	return h
}

// AddRun adds a run to the hyperlink
func (h *Hyperlink) AddRun() *Run {
	r := NewRun()
	// Apply hyperlink properties to the run
	if h.Properties != nil {
		r.Properties = h.Properties.Clone()
	}
	h.Children = append(h.Children, r)
	return r
}

// AddText adds text to the hyperlink
func (h *Hyperlink) AddText(text string) *Run {
	r := h.AddRun()
	r.AddText(text)
	return r
}

// AddFormattedText adds text with additional formatting
func (h *Hyperlink) AddFormattedText(text string, format func(*Run)) *Run {
	r := h.AddRun()
	r.AddText(text)
	if format != nil {
		format(r)
	}
	return r
}

// SetStyle sets the hyperlink style
func (h *Hyperlink) SetStyle(styleID string) *Hyperlink {
	if h.Properties == nil {
		h.Properties = properties.NewRunProperties()
	}
	h.Properties.StyleID = styleID

	// Apply to all children
	for _, child := range h.Children {
		if run, ok := child.(*Run); ok {
			run.Properties.StyleID = styleID
		}
	}
	return h
}

// RemoveUnderline removes the default underline
func (h *Hyperlink) RemoveUnderline() *Hyperlink {
	if h.Properties == nil {
		h.Properties = properties.NewRunProperties()
	}
	h.Properties.Underline = "none"

	// Apply to all children
	for _, child := range h.Children {
		if run, ok := child.(*Run); ok {
			run.Properties.Underline = "none"
		}
	}
	return h
}

// SetColor sets the hyperlink color
func (h *Hyperlink) SetColor(color string) *Hyperlink {
	if h.Properties == nil {
		h.Properties = properties.NewRunProperties()
	}
	h.Properties.Color = color

	// Apply to all children
	for _, child := range h.Children {
		if run, ok := child.(*Run); ok {
			run.Properties.Color = color
		}
	}
	return h
}

// Clone creates a deep copy of the hyperlink
func (h *Hyperlink) Clone() *Hyperlink {
	clone := &Hyperlink{
		ID:          h.ID,
		URL:         h.URL,
		Tooltip:     h.Tooltip,
		Typ:         h.Typ,
		Anchor:      h.Anchor,
		DocLocation: h.DocLocation,
		TargetFrame: h.TargetFrame,
		History:     h.History,
		ScreenTip:   h.ScreenTip,
		Children:    make([]ParagraphChild, 0, len(h.Children)),
	}

	if h.Properties != nil {
		clone.Properties = h.Properties.Clone()
	}

	// Clone children
	for _, child := range h.Children {
		if run, ok := child.(*Run); ok {
			clone.Children = append(clone.Children, run.Clone())
		}
	}

	return clone
}

// Validate validates the hyperlink
func (h *Hyperlink) Validate() error {
	if h.Typ == HyperlinkTypeExternal && h.URL == "" {
		return fmt.Errorf("external hyperlink must have a URL")
	}

	if h.Typ == HyperlinkTypeInternal && h.Anchor == "" && h.DocLocation == "" {
		return fmt.Errorf("internal hyperlink must have an anchor or document location")
	}

	if h.Typ == HyperlinkTypeExternal && h.ID == "" {
		return fmt.Errorf("external hyperlink must have a relationship ID")
	}

	// Validate URL format for external links
	if h.Typ == HyperlinkTypeExternal {
		if !isValidURL(h.URL) && !isValidEmail(h.URL) {
			return fmt.Errorf("invalid URL format: %s", h.URL)
		}
	}

	// Validate children
	if len(h.Children) == 0 {
		return fmt.Errorf("hyperlink must have at least one child element")
	}

	return nil
}

// XML generates the XML representation of the hyperlink
func (h *Hyperlink) XML() ([]byte, error) {
	var buf bytes.Buffer

	// Start hyperlink tag
	buf.WriteString(`<w:hyperlink`)

	// Add attributes based on type
	if h.Typ == HyperlinkTypeExternal && h.ID != "" {
		buf.WriteString(fmt.Sprintf(` r:id="%s"`, h.ID))
	}

	if h.Anchor != "" {
		buf.WriteString(fmt.Sprintf(` w:anchor="%s"`, h.Anchor))
	}

	if h.DocLocation != "" {
		buf.WriteString(fmt.Sprintf(` w:docLocation="%s"`, h.DocLocation))
	}

	if h.Tooltip != "" {
		buf.WriteString(fmt.Sprintf(` w:tooltip="%s"`, escapeXMLAttribute(h.Tooltip)))
	}

	if h.ScreenTip != "" && h.ScreenTip != h.Tooltip {
		buf.WriteString(fmt.Sprintf(` w:screenTip="%s"`, escapeXMLAttribute(h.ScreenTip)))
	}

	if h.TargetFrame != "" {
		buf.WriteString(fmt.Sprintf(` w:tgtFrame="%s"`, h.TargetFrame))
	}

	if h.History {
		buf.WriteString(` w:history="1"`)
	} else {
		buf.WriteString(` w:history="0"`)
	}

	buf.WriteString(`>`)

	// Add children (usually runs with hyperlink formatting)
	for _, child := range h.Children {
		childXML, err := child.XML()
		if err != nil {
			return nil, fmt.Errorf("generating hyperlink child XML: %w", err)
		}
		buf.Write(childXML)
	}

	// Close hyperlink tag
	buf.WriteString(`</w:hyperlink>`)

	return buf.Bytes(), nil
}

// Helper functions

// generateRelationshipID generates a unique relationship ID
func generateRelationshipID() string {
	return "rId" + strings.ReplaceAll(uuid.New().String(), "-", "")[:8]
}

// isValidURL checks if a URL is valid
func isValidURL(url string) bool {
	// Simple validation - check for common URL patterns
	validPrefixes := []string{
		"http://",
		"https://",
		"ftp://",
		"file://",
		"mailto:",
	}

	for _, prefix := range validPrefixes {
		if strings.HasPrefix(strings.ToLower(url), prefix) {
			return true
		}
	}

	return false
}

// isValidEmail checks if an email address is valid
func isValidEmail(email string) bool {
	// Remove mailto: prefix if present
	email = strings.TrimPrefix(email, "mailto:")

	// Very simple email validation
	return strings.Contains(email, "@") && strings.Contains(email, ".")
}

// escapeXMLAttribute escapes special characters in XML attributes
func escapeXMLAttribute(s string) string {
	s = strings.ReplaceAll(s, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, "\"", "&quot;")
	s = strings.ReplaceAll(s, "'", "&apos;")
	return s
}

// HyperlinkRelationship represents a hyperlink relationship for the document
type HyperlinkRelationship struct {
	ID         string
	Type       string
	Target     string
	TargetMode string
}

// NewHyperlinkRelationship creates a new hyperlink relationship
func NewHyperlinkRelationship(id, target string) *HyperlinkRelationship {
	return &HyperlinkRelationship{
		ID:         id,
		Type:       "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
		Target:     target,
		TargetMode: "External",
	}
}

// CreateHyperlinkWithRelationship creates a hyperlink and its relationship
func CreateHyperlinkWithRelationship(text, url string) (*Hyperlink, *HyperlinkRelationship) {
	h := NewHyperlink(text, url)
	rel := NewHyperlinkRelationship(h.ID, url)
	return h, rel
}
