package elements

import (
	"bytes"
	"fmt"
	"strings"

	"github.com/google/uuid"
)

// Field represents a field element (e.g., page numbers, dates, references)
type Field struct {
	ID          string
	Typ         string // Field type (PAGE, DATE, TIME, etc.)
	Instruction string // Full field instruction
	Format      string // Format string
	Result      string // Cached result
	IsDirty     bool   // Needs recalculation
	IsLocked    bool   // Prevent updates
	Properties  *FieldProperties
}

// FieldProperties contains field-specific properties
type FieldProperties struct {
	UpdateOnOpen   bool
	PreserveFormat bool
	DoNotShadeForm bool
}

// Common field types
const (
	FieldTypePage        = "PAGE"
	FieldTypeNumPages    = "NUMPAGES"
	FieldTypeDate        = "DATE"
	FieldTypeTime        = "TIME"
	FieldTypeFileName    = "FILENAME"
	FieldTypeAuthor      = "AUTHOR"
	FieldTypeTitle       = "TITLE"
	FieldTypeSubject     = "SUBJECT"
	FieldTypeTOC         = "TOC"
	FieldTypeRef         = "REF"
	FieldTypePageRef     = "PAGEREF"
	FieldTypeHyperlink   = "HYPERLINK"
	FieldTypeSeq         = "SEQ"
	FieldTypeMergeField  = "MERGEFIELD"
	FieldTypeIf          = "IF"
	FieldTypeIndex       = "INDEX"
	FieldTypeQuote       = "QUOTE"
	FieldTypeInclude     = "INCLUDE"
	FieldTypeStyleRef    = "STYLEREF"
	FieldTypeDocVariable = "DOCVARIABLE"
)

// NewField creates a new field
func NewField(fieldType, instruction string) *Field {
	return &Field{
		ID:          generateFieldID(),
		Typ:         fieldType,
		Instruction: instruction,
		Properties:  &FieldProperties{},
	}
}

// NewPageField creates a page number field
func NewPageField(format string) *Field {
	instruction := "PAGE"
	if format != "" {
		instruction += " \\* " + format
	}
	return NewField(FieldTypePage, instruction)
}

// NewNumPagesField creates a total pages field
func NewNumPagesField(format string) *Field {
	instruction := "NUMPAGES"
	if format != "" {
		instruction += " \\* " + format
	}
	return NewField(FieldTypeNumPages, instruction)
}

// NewDateField creates a date field
func NewDateField(format string) *Field {
	instruction := "DATE"
	if format != "" {
		instruction += ` \@ "` + format + `"`
	}
	return NewField(FieldTypeDate, instruction)
}

// NewTimeField creates a time field
func NewTimeField(format string) *Field {
	instruction := "TIME"
	if format != "" {
		instruction += ` \@ "` + format + `"`
	}
	return NewField(FieldTypeTime, instruction)
}

// NewTOCField creates a table of contents field
func NewTOCField(options ...string) *Field {
	instruction := "TOC"
	for _, opt := range options {
		instruction += " " + opt
	}
	return NewField(FieldTypeTOC, instruction)
}

// NewRefField creates a cross-reference field
func NewRefField(bookmarkName string, options ...string) *Field {
	instruction := "REF " + bookmarkName
	for _, opt := range options {
		instruction += " " + opt
	}
	return NewField(FieldTypeRef, instruction)
}

// NewMergeField creates a mail merge field
func NewMergeField(fieldName string, format string) *Field {
	instruction := "MERGEFIELD " + fieldName
	if format != "" {
		instruction += " \\* " + format
	}
	return NewField(FieldTypeMergeField, instruction)
}

// NewIfField creates a conditional field
func NewIfField(condition, trueText, falseText string) *Field {
	instruction := fmt.Sprintf(`IF %s "%s" "%s"`, condition, trueText, falseText)
	return NewField(FieldTypeIf, instruction)
}

// Type returns the element type
func (f *Field) Type() string {
	return "field"
}

// SetFormat sets the field format
func (f *Field) SetFormat(format string) *Field {
	f.Format = format
	return f
}

// SetResult sets the cached result
func (f *Field) SetResult(result string) *Field {
	f.Result = result
	return f
}

// SetDirty marks the field as needing recalculation
func (f *Field) SetDirty(dirty bool) *Field {
	f.IsDirty = dirty
	return f
}

// SetLocked sets whether the field is locked
func (f *Field) SetLocked(locked bool) *Field {
	f.IsLocked = locked
	return f
}

// GetText returns the field result or instruction
func (f *Field) GetText() string {
	if f.Result != "" {
		return f.Result
	}
	return "{" + f.Instruction + "}"
}

// XML generates the XML representation
func (f *Field) XML() ([]byte, error) {
	var buf bytes.Buffer

	// Field character begin
	buf.WriteString(`<w:r>`)
	buf.WriteString(`<w:fldChar w:fldCharType="begin"`)
	if f.IsLocked {
		buf.WriteString(` w:fldLock="1"`)
	}
	if f.IsDirty {
		buf.WriteString(` w:dirty="1"`)
	}
	buf.WriteString(`/>`)
	buf.WriteString(`</w:r>`)

	// Field instruction
	buf.WriteString(`<w:r>`)
	buf.WriteString(`<w:instrText xml:space="preserve"> `)
	buf.WriteString(escapeXML(f.Instruction))
	buf.WriteString(` </w:instrText>`)
	buf.WriteString(`</w:r>`)

	// Field character separate
	buf.WriteString(`<w:r>`)
	buf.WriteString(`<w:fldChar w:fldCharType="separate"/>`)
	buf.WriteString(`</w:r>`)

	// Field result
	if f.Result != "" {
		buf.WriteString(`<w:r>`)
		buf.WriteString(`<w:t>`)
		buf.WriteString(escapeXML(f.Result))
		buf.WriteString(`</w:t>`)
		buf.WriteString(`</w:r>`)
	}

	// Field character end
	buf.WriteString(`<w:r>`)
	buf.WriteString(`<w:fldChar w:fldCharType="end"/>`)
	buf.WriteString(`</w:r>`)

	return buf.Bytes(), nil
}

// generateFieldID generates a unique field ID
func generateFieldID() string {
	return "fld" + strings.ReplaceAll(uuid.New().String(), "-", "")[:8]
}
