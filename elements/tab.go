package elements

// Tab represents a tab character
type Tab struct {
	Position  int    // Position in twips (optional, for positional tabs)
	Alignment string // Alignment: "left", "center", "right", "decimal", "bar"
	Leader    string // Leader character: "dot", "hyphen", "underscore", "heavy", "middleDot"
}

// NewTab creates a new tab
func NewTab() *Tab {
	return &Tab{}
}

// NewPositionalTab creates a tab at a specific position
func NewPositionalTab(position int, alignment string) *Tab {
	return &Tab{
		Position:  position,
		Alignment: alignment,
	}
}

// Type returns the element type
func (t *Tab) Type() string {
	return "tab"
}

// SetPosition sets the tab position in twips
func (t *Tab) SetPosition(position int) *Tab {
	t.Position = position
	return t
}

// SetAlignment sets the tab alignment
func (t *Tab) SetAlignment(alignment string) *Tab {
	t.Alignment = alignment
	return t
}

// SetLeader sets the tab leader character
func (t *Tab) SetLeader(leader string) *Tab {
	t.Leader = leader
	return t
}

// XML generates the XML representation
func (t *Tab) XML() ([]byte, error) {
	// Simple tab (most common case)
	if t.Position == 0 && t.Alignment == "" && t.Leader == "" {
		return []byte(`<w:tab/>`), nil
	}

	// Positional tab (requires run properties)
	// This would typically be handled at the paragraph level
	// For now, return a simple tab
	return []byte(`<w:tab/>`), nil
}
