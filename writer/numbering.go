package writer

import (
	"bytes"
	"fmt"
	"io"
	"log"
)

var _ zipWritable = (*NumberingDefinitions)(nil)

// NumberingDefinitions contains all numbering definitions for the document
type NumberingDefinitions struct {
	AbstractNums []AbstractNum
	Nums         []Num
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

// NewNumberingDefinitions creates default numbering definitions
func newNumberingDefinitions() *NumberingDefinitions {
	return &NumberingDefinitions{
		AbstractNums: createDefaultAbstractNums(),
		Nums:         createDefaultNums(),
	}
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
	var buf bytes.Buffer

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

	return buf.Bytes(), nil
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
	var buf bytes.Buffer

	buf.WriteString(fmt.Sprintf(`  <w:abstractNum w:abstractNumId="%d">`, an.ID))
	buf.WriteString("\n")

	// Multi-level type
	if an.MultiLevel {
		buf.WriteString(`    <w:multiLevelType w:val="multilevel"/>`)
	} else {
		buf.WriteString(`    <w:multiLevelType w:val="singleLevel"/>`)
	}
	buf.WriteString("\n")

	// Name
	if an.Name != "" {
		buf.WriteString(fmt.Sprintf(`    <w:name w:val="%s"/>`, an.Name))
		buf.WriteString("\n")
	}

	// Generate levels
	for _, level := range an.Levels {
		buf.WriteString(level.GenerateXML())
		buf.WriteString("\n")
	}

	buf.WriteString(`  </w:abstractNum>`)

	return buf.String()
}

// GenerateXML generates XML for a numbering level
func (l *Level) GenerateXML() string {
	var buf bytes.Buffer

	buf.WriteString(fmt.Sprintf(`    <w:lvl w:ilvl="%d">`, l.Level))
	buf.WriteString("\n")

	// Start value
	buf.WriteString(fmt.Sprintf(`      <w:start w:val="%d"/>`, l.Start))
	buf.WriteString("\n")

	// Number format
	if l.NumFormat == "bullet" {
		buf.WriteString(`      <w:numFmt w:val="bullet"/>`)
		buf.WriteString("\n")

		// Level text for bullet
		buf.WriteString(fmt.Sprintf(`      <w:lvlText w:val="%s"/>`, l.BulletChar))
		buf.WriteString("\n")

		// Font for bullet
		if l.Font != "" {
			buf.WriteString(`      <w:rPr>`)
			buf.WriteString("\n")
			buf.WriteString(fmt.Sprintf(`        <w:rFonts w:ascii="%s" w:hAnsi="%s" w:hint="default"/>`, l.Font, l.Font))
			buf.WriteString("\n")
			buf.WriteString(`      </w:rPr>`)
			buf.WriteString("\n")
		}
	} else {
		// Number format for numbered lists
		buf.WriteString(fmt.Sprintf(`      <w:numFmt w:val="%s"/>`, l.NumFormat))
		buf.WriteString("\n")

		// Level text (e.g., "%1.", "%1.%2")
		buf.WriteString(fmt.Sprintf(`      <w:lvlText w:val="%s"/>`, l.LevelText))
		buf.WriteString("\n")

		// Legal numbering
		if l.IsLegalNum {
			buf.WriteString(`      <w:isLgl/>`)
			buf.WriteString("\n")
		}
	}

	// Level justification
	buf.WriteString(fmt.Sprintf(`      <w:lvlJc w:val="%s"/>`, l.LevelJc))
	buf.WriteString("\n")

	// Paragraph properties
	buf.WriteString(`      <w:pPr>`)
	buf.WriteString("\n")

	// Indentation
	buf.WriteString(fmt.Sprintf(`        <w:ind w:left="%d" w:hanging="%d"/>`, l.IndentLeft, l.IndentHanging))
	buf.WriteString("\n")

	buf.WriteString(`      </w:pPr>`)
	buf.WriteString("\n")

	// Suffix (tab, space, or nothing)
	if l.Suffix != "" {
		buf.WriteString(fmt.Sprintf(`      <w:suff w:val="%s"/>`, l.Suffix))
		buf.WriteString("\n")
	}

	// Paragraph style reference
	if l.PStyle != "" {
		buf.WriteString(fmt.Sprintf(`      <w:pStyle w:val="%s"/>`, l.PStyle))
		buf.WriteString("\n")
	}

	buf.WriteString(`    </w:lvl>`)

	return buf.String()
}

// GenerateXML generates XML for a concrete numbering instance
func (n *Num) GenerateXML() string {
	var buf bytes.Buffer

	buf.WriteString(fmt.Sprintf(`  <w:num w:numId="%d">`, n.ID))
	buf.WriteString("\n")

	// Reference to abstract numbering
	buf.WriteString(fmt.Sprintf(`    <w:abstractNumId w:val="%d"/>`, n.AbstractID))
	buf.WriteString("\n")

	// Level overrides
	for _, override := range n.Overrides {
		buf.WriteString(fmt.Sprintf(`    <w:lvlOverride w:ilvl="%d">`, override.Level))
		buf.WriteString("\n")
		buf.WriteString(fmt.Sprintf(`      <w:startOverride w:val="%d"/>`, override.StartOverride))
		buf.WriteString("\n")
		buf.WriteString(`    </w:lvlOverride>`)
		buf.WriteString("\n")
	}

	buf.WriteString(`  </w:num>`)

	return buf.String()
}
