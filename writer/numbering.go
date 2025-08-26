package writer

import (
	"bytes"
	"fmt"
	"io"
	"log"
	"strings"
	"sync"

	"github.com/didikprabowo/mbadocx/types"
)

// bufferPool provides a pool of reusable byte buffers to reduce memory allocations
var bufferPool = sync.Pool{
	New: func() interface{} {
		return &bytes.Buffer{}
	},
}

// getBuffer gets a buffer from the pool
func getBuffer() *bytes.Buffer {
	return bufferPool.Get().(*bytes.Buffer)
}

// putBuffer returns a buffer to the pool after resetting it
func putBuffer(buf *bytes.Buffer) {
	buf.Reset()
	bufferPool.Put(buf)
}

var _ zipWritable = (*NumberingDefinitions)(nil)

// NumberingDefinitions contains all numbering definitions for the document
type NumberingDefinitions struct {
	AbstractNums []AbstractNum
	Nums         []Num
	document     types.Document // Add document reference for lazy loading
	initialized  bool           // Track if numbering has been initialized
}

// AbstractNum defines an abstract numbering definition
type AbstractNum struct {
	ID         int
	MultiLevel bool
	Levels     []Level
	Name       string
}

// Level defines a numbering level
type Level struct {
	Level         int
	Start         int
	NumFormat     string // bullet, decimal, upperRoman, lowerRoman, upperLetter, lowerLetter, etc.
	LevelText     string
	LevelJc       string // left, center, right, justify
	PStyle        string
	IsLegalNum    bool
	Suffix        string // tab, space, nothing
	BulletChar    string // For bullet lists
	Font          string // Font for bullet
	IndentLeft    int    // In twips
	IndentHanging int    // In twips
}

// Num defines a concrete numbering instance
type Num struct {
	ID         int
	AbstractID int
	Overrides  []LevelOverride
}

// LevelOverride allows overriding specific levels
type LevelOverride struct {
	Level         int
	StartOverride int
}

// NewNumberingDefinitions creates numbering definitions with lazy loading
func newNumberingDefinitions() *NumberingDefinitions {
	return &NumberingDefinitions{
		initialized: false,
	}
}

// NewNumberingDefinitionsForDocument creates numbering definitions with document reference
func newNumberingDefinitionsForDocument(doc types.Document) *NumberingDefinitions {
	return &NumberingDefinitions{
		document:    doc,
		initialized: false,
	}
}

// ensureInitialized initializes numbering definitions only when needed
func (num *NumberingDefinitions) ensureInitialized() {
	if num.initialized {
		return
	}
	
	// Only create default numbering if the document actually uses lists
	if num.document != nil && !num.hasNumberedElements() {
		// Create minimal numbering for compatibility
		num.AbstractNums = []AbstractNum{}
		num.Nums = []Num{}
	} else {
		// Create full default numbering
		num.AbstractNums = createDefaultAbstractNums()
		num.Nums = createDefaultNums()
	}
	
	num.initialized = true
}

// hasNumberedElements checks if the document contains any numbered or bulleted elements
func (num *NumberingDefinitions) hasNumberedElements() bool {
	if num.document == nil {
		return true // Default to creating full numbering if no document reference
	}
	
	// Check document elements for list usage
	for _, elem := range num.document.Body().GetElements() {
		// This would need to be implemented based on your element types
		// For now, we'll assume lists might be used
		_ = elem
	}
	
	// For safety, assume numbering might be needed
	// In a real implementation, you'd check the actual element types
	return true
}

func createDefaultAbstractNums() []AbstractNum {
	return []AbstractNum{
		// Abstract Num 0: Standard Bullet List
		{
			ID:         0,
			MultiLevel: true,
			Name:       "Standard Bullet List",
			Levels: []Level{
				// Level 0
				{
					Level:         0,
					Start:         1,
					NumFormat:     "bullet",
					LevelText:     "•",
					BulletChar:    "•",
					LevelJc:       "left",
					Suffix:        "tab",
					Font:          "Symbol",
					IndentLeft:    720, // 0.5 inch
					IndentHanging: 360, // 0.25 inch
				},
				// Level 1
				{
					Level:         1,
					Start:         1,
					NumFormat:     "bullet",
					LevelText:     "○",
					BulletChar:    "○",
					LevelJc:       "left",
					Suffix:        "tab",
					Font:          "Symbol",
					IndentLeft:    1440, // 1 inch
					IndentHanging: 360,
				},
				// Level 2
				{
					Level:         2,
					Start:         1,
					NumFormat:     "bullet",
					LevelText:     "▪",
					BulletChar:    "▪",
					LevelJc:       "left",
					Suffix:        "tab",
					Font:          "Symbol",
					IndentLeft:    2160, // 1.5 inch
					IndentHanging: 360,
				},
				// Level 3
				{
					Level:         3,
					Start:         1,
					NumFormat:     "bullet",
					LevelText:     "▫",
					BulletChar:    "▫",
					LevelJc:       "left",
					Suffix:        "tab",
					Font:          "Symbol",
					IndentLeft:    2880, // 2 inch
					IndentHanging: 360,
				},
			},
		},
		// Abstract Num 1: Decimal Numbering
		{
			ID:         1,
			MultiLevel: true,
			Name:       "Decimal Numbering",
			Levels: []Level{
				// Level 0: 1. 2. 3.
				{
					Level:         0,
					Start:         1,
					NumFormat:     "decimal",
					LevelText:     "%1.",
					LevelJc:       "left",
					Suffix:        "tab",
					IndentLeft:    720,
					IndentHanging: 360,
				},
				// Level 1: 1.1 1.2 1.3
				{
					Level:         1,
					Start:         1,
					NumFormat:     "decimal",
					LevelText:     "%1.%2",
					LevelJc:       "left",
					Suffix:        "tab",
					IndentLeft:    1440,
					IndentHanging: 540,
				},
				// Level 2: a. b. c.
				{
					Level:         2,
					Start:         1,
					NumFormat:     "lowerLetter",
					LevelText:     "%3.",
					LevelJc:       "left",
					Suffix:        "tab",
					IndentLeft:    2160,
					IndentHanging: 360,
				},
				// Level 3: i. ii. iii.
				{
					Level:         3,
					Start:         1,
					NumFormat:     "lowerRoman",
					LevelText:     "%4.",
					LevelJc:       "left",
					Suffix:        "tab",
					IndentLeft:    2880,
					IndentHanging: 360,
				},
			},
		},
		// Abstract Num 2: Legal Style Numbering
		{
			ID:         2,
			MultiLevel: true,
			Name:       "Legal Style",
			Levels: []Level{
				// Level 0: 1.
				{
					Level:         0,
					Start:         1,
					NumFormat:     "decimal",
					LevelText:     "%1.",
					LevelJc:       "left",
					Suffix:        "tab",
					IsLegalNum:    true,
					IndentLeft:    360,
					IndentHanging: 360,
				},
				// Level 1: 1.1
				{
					Level:         1,
					Start:         1,
					NumFormat:     "decimal",
					LevelText:     "%1.%2",
					LevelJc:       "left",
					Suffix:        "tab",
					IsLegalNum:    true,
					IndentLeft:    720,
					IndentHanging: 432,
				},
				// Level 2: 1.1.1
				{
					Level:         2,
					Start:         1,
					NumFormat:     "decimal",
					LevelText:     "%1.%2.%3",
					LevelJc:       "left",
					Suffix:        "tab",
					IsLegalNum:    true,
					IndentLeft:    1080,
					IndentHanging: 504,
				},
			},
		},
		// Abstract Num 3: Roman Numerals
		{
			ID:         3,
			MultiLevel: false,
			Name:       "Roman Numerals",
			Levels: []Level{
				{
					Level:         0,
					Start:         1,
					NumFormat:     "upperRoman",
					LevelText:     "%1.",
					LevelJc:       "left",
					Suffix:        "tab",
					IndentLeft:    720,
					IndentHanging: 360,
				},
			},
		},
		// Abstract Num 4: Custom Bullet Symbols
		{
			ID:         4,
			MultiLevel: true,
			Name:       "Custom Symbols",
			Levels: []Level{
				// Level 0: ➤
				{
					Level:         0,
					Start:         1,
					NumFormat:     "bullet",
					LevelText:     "➤",
					BulletChar:    "➤",
					LevelJc:       "left",
					Suffix:        "tab",
					Font:          "Wingdings",
					IndentLeft:    720,
					IndentHanging: 360,
				},
				// Level 1: ✓
				{
					Level:         1,
					Start:         1,
					NumFormat:     "bullet",
					LevelText:     "✓",
					BulletChar:    "✓",
					LevelJc:       "left",
					Suffix:        "tab",
					Font:          "Wingdings",
					IndentLeft:    1440,
					IndentHanging: 360,
				},
				// Level 2: ★
				{
					Level:         2,
					Start:         1,
					NumFormat:     "bullet",
					LevelText:     "★",
					BulletChar:    "★",
					LevelJc:       "left",
					Suffix:        "tab",
					Font:          "Wingdings",
					IndentLeft:    2160,
					IndentHanging: 360,
				},
			},
		},
	}
}

// createDefaultNums creates concrete numbering instances
func createDefaultNums() []Num {
	return []Num{
		{ID: 1, AbstractID: 0}, // Bullet list
		{ID: 2, AbstractID: 1}, // Decimal numbering
		{ID: 3, AbstractID: 2}, // Legal numbering
		{ID: 4, AbstractID: 3}, // Roman numerals
		{ID: 5, AbstractID: 4}, // Custom symbols
	}
}

func (num *NumberingDefinitions) Path() string {
	return "word/numbering.xml"
}

func (num *NumberingDefinitions) Byte() ([]byte, error) {
	// Ensure numbering is initialized before generating XML
	num.ensureInitialized()
	
	// If no numbering definitions are needed, return minimal XML
	if len(num.AbstractNums) == 0 && len(num.Nums) == 0 {
		buf := getBuffer()
		defer putBuffer(buf)
		
		buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
		buf.WriteString("\n")
		buf.WriteString(`<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`)
		
		result := make([]byte, buf.Len())
		copy(result, buf.Bytes())
		return result, nil
	}
	
	buf := getBuffer()
	defer putBuffer(buf)

	// XML declaration
	buf.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	buf.WriteString("\n")

	// Numbering root element with namespaces
	buf.WriteString(`<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`)
	buf.WriteString(` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`)
	buf.WriteString(` xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">`)
	buf.WriteString("\n")

	// Generate abstract numbering definitions
	for _, abstractNum := range num.AbstractNums {
		buf.WriteString(abstractNum.GenerateXML())
		buf.WriteString("\n")
	}

	// Generate concrete numbering instances
	for _, num := range num.Nums {
		buf.WriteString(num.GenerateXML())
		buf.WriteString("\n")
	}

	// Close numbering element
	buf.WriteString(`</w:numbering>`)

	log.Printf("'%s' has been created.\n", num.Path())
	// log.Print(buf.String())

	// Make a copy of the bytes before returning the buffer to the pool
	result := make([]byte, buf.Len())
	copy(result, buf.Bytes())
	return result, nil
}

// WriteTo writes the XML content to the given writer.
func (num *NumberingDefinitions) WriteTo(w io.Writer) (int64, error) {
	data, err := num.Byte()
	if err != nil {
		return 0, err
	}

	n, err := w.Write(data)
	return int64(n), err
}

// GenerateXML generates XML for an abstract numbering definition
func (an *AbstractNum) GenerateXML() string {
	var builder strings.Builder
	// Pre-allocate capacity to reduce reallocations
	builder.Grow(1024)

	builder.WriteString(fmt.Sprintf(`  <w:abstractNum w:abstractNumId="%d">`, an.ID))
	builder.WriteString("\n")

	// Multi-level type
	if an.MultiLevel {
		builder.WriteString(`    <w:multiLevelType w:val="multilevel"/>`)
	} else {
		builder.WriteString(`    <w:multiLevelType w:val="singleLevel"/>`)
	}
	builder.WriteString("\n")

	// Name
	if an.Name != "" {
		builder.WriteString(fmt.Sprintf(`    <w:name w:val="%s"/>`, an.Name))
		builder.WriteString("\n")
	}

	// Generate levels
	for _, level := range an.Levels {
		builder.WriteString(level.GenerateXML())
		builder.WriteString("\n")
	}

	builder.WriteString(`  </w:abstractNum>`)

	return builder.String()
}

// GenerateXML generates XML for a numbering level
func (l *Level) GenerateXML() string {
	var builder strings.Builder
	// Pre-allocate capacity to reduce reallocations
	builder.Grow(512)

	builder.WriteString(fmt.Sprintf(`    <w:lvl w:ilvl="%d">`, l.Level))
	builder.WriteString("\n")

	// Start value
	builder.WriteString(fmt.Sprintf(`      <w:start w:val="%d"/>`, l.Start))
	builder.WriteString("\n")

	// Number format
	if l.NumFormat == "bullet" {
		builder.WriteString(`      <w:numFmt w:val="bullet"/>`)
		builder.WriteString("\n")

		// Level text for bullet
		builder.WriteString(fmt.Sprintf(`      <w:lvlText w:val="%s"/>`, l.BulletChar))
		builder.WriteString("\n")

		// Font for bullet
		if l.Font != "" {
			builder.WriteString(`      <w:rPr>`)
			builder.WriteString("\n")
			builder.WriteString(fmt.Sprintf(`        <w:rFonts w:ascii="%s" w:hAnsi="%s" w:hint="default"/>`, l.Font, l.Font))
			builder.WriteString("\n")
			builder.WriteString(`      </w:rPr>`)
			builder.WriteString("\n")
		}
	} else {
		// Number format for numbered lists
		builder.WriteString(fmt.Sprintf(`      <w:numFmt w:val="%s"/>`, l.NumFormat))
		builder.WriteString("\n")

		// Level text (e.g., "%1.", "%1.%2")
		builder.WriteString(fmt.Sprintf(`      <w:lvlText w:val="%s"/>`, l.LevelText))
		builder.WriteString("\n")

		// Legal numbering
		if l.IsLegalNum {
			builder.WriteString(`      <w:isLgl/>`)
			builder.WriteString("\n")
		}
	}

	// Level justification
	builder.WriteString(fmt.Sprintf(`      <w:lvlJc w:val="%s"/>`, l.LevelJc))
	builder.WriteString("\n")

	// Paragraph properties
	builder.WriteString(`      <w:pPr>`)
	builder.WriteString("\n")

	// Indentation
	builder.WriteString(fmt.Sprintf(`        <w:ind w:left="%d" w:hanging="%d"/>`, l.IndentLeft, l.IndentHanging))
	builder.WriteString("\n")

	builder.WriteString(`      </w:pPr>`)
	builder.WriteString("\n")

	// Suffix (tab, space, or nothing)
	if l.Suffix != "" {
		builder.WriteString(fmt.Sprintf(`      <w:suff w:val="%s"/>`, l.Suffix))
		builder.WriteString("\n")
	}

	// Paragraph style reference
	if l.PStyle != "" {
		builder.WriteString(fmt.Sprintf(`      <w:pStyle w:val="%s"/>`, l.PStyle))
		builder.WriteString("\n")
	}

	builder.WriteString(`    </w:lvl>`)

	return builder.String()
}

// GenerateXML generates XML for a concrete numbering instance
func (n *Num) GenerateXML() string {
	var builder strings.Builder
	// Pre-allocate capacity to reduce reallocations
	builder.Grow(256)

	builder.WriteString(fmt.Sprintf(`  <w:num w:numId="%d">`, n.ID))
	builder.WriteString("\n")

	// Reference to abstract numbering
	builder.WriteString(fmt.Sprintf(`    <w:abstractNumId w:val="%d"/>`, n.AbstractID))
	builder.WriteString("\n")

	// Level overrides
	for _, override := range n.Overrides {
		builder.WriteString(fmt.Sprintf(`    <w:lvlOverride w:ilvl="%d">`, override.Level))
		builder.WriteString("\n")
		builder.WriteString(fmt.Sprintf(`      <w:startOverride w:val="%d"/>`, override.StartOverride))
		builder.WriteString("\n")
		builder.WriteString(`    </w:lvlOverride>`)
		builder.WriteString("\n")
	}

	builder.WriteString(`  </w:num>`)

	return builder.String()
}
