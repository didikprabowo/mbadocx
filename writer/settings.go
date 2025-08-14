package writer

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"strconv"

	"github.com/didikprabowo/mbadocx/settings"
	"github.com/didikprabowo/mbadocx/types"
)

type Settings struct {
	document types.Document
}

// SettingsXML represents the root settings XML structure
type SettingsXML struct {
	XMLName                   xml.Name                      `xml:"w:settings"`
	XmlnsW                    string                        `xml:"xmlns:w,attr"`
	Zoom                      *ZoomXML                      `xml:"w:zoom,omitempty"`
	DefaultTabStop            *DefaultTabStopXML            `xml:"w:defaultTabStop,omitempty"`
	CharacterSpacingControl   *CharacterSpacingControlXML   `xml:"w:characterSpacingControl,omitempty"`
	Compat                    *CompatXML                    `xml:"w:compat,omitempty"`
	ThemeFontLang             *ThemeFontLangXML             `xml:"w:themeFontLang,omitempty"`
	ProofState                *ProofStateXML                `xml:"w:proofState,omitempty"`
	DefaultView               *ViewXML                      `xml:"w:defaultView,omitempty"`
	TrackRevisions            *TrackRevisionsXML            `xml:"w:trackRevisions,omitempty"`
	TrackFormatting           *TrackFormattingXML           `xml:"w:trackFormatting,omitempty"`
	DocumentProtection        *DocumentProtectionXML        `xml:"w:documentProtection,omitempty"`
	DocVars                   *DocVarsXML                   `xml:"w:docVars,omitempty"`
	MailMerge                 *MailMergeXML                 `xml:"w:mailMerge,omitempty"`
	DoNotPromptForConvert     *DoNotPromptForConvertXML     `xml:"w:doNotPromptForConvert,omitempty"`
	DoNotAutoCompressPictures *DoNotAutoCompressPicturesXML `xml:"w:doNotAutoCompressPictures,omitempty"`
	// Additional settings
	DrawingGridHorizontalSpacing *DrawingGridHorizontalSpacingXML `xml:"w:drawingGridHorizontalSpacing,omitempty"`
	DrawingGridVerticalSpacing   *DrawingGridVerticalSpacingXML   `xml:"w:drawingGridVerticalSpacing,omitempty"`
	HideSpellingErrors           *HideSpellingErrorsXML           `xml:"w:hideSpellingErrors,omitempty"`
	HideGrammaticalErrors        *HideGrammaticalErrorsXML        `xml:"w:hideGrammaticalErrors,omitempty"`
}

// ZoomXML represents zoom settings
type ZoomXML struct {
	Percent int    `xml:"w:percent,attr"`
	Val     string `xml:"w:val,attr,omitempty"`
}

// DefaultTabStopXML represents default tab stop settings
type DefaultTabStopXML struct {
	Val int `xml:"w:val,attr"`
}

// CharacterSpacingControlXML represents character spacing control
type CharacterSpacingControlXML struct {
	Val string `xml:"w:val,attr"`
}

// CompatXML represents compatibility settings
type CompatXML struct {
	CompatSettings []CompatSettingXML `xml:"w:compatSetting"`
}

// CompatSettingXML represents individual compatibility setting
type CompatSettingXML struct {
	Name string `xml:"w:name,attr"`
	URI  string `xml:"w:uri,attr"`
	Val  string `xml:"w:val,attr"`
}

// ThemeFontLangXML represents theme font language settings
type ThemeFontLangXML struct {
	Val      string `xml:"w:val,attr,omitempty"`
	EastAsia string `xml:"w:eastAsia,attr,omitempty"`
	Bidi     string `xml:"w:bidi,attr,omitempty"`
}

// ProofStateXML represents proofing state
type ProofStateXML struct {
	Spelling string `xml:"w:spelling,attr"`
	Grammar  string `xml:"w:grammar,attr"`
}

// ViewXML represents view settings
type ViewXML struct {
	Val string `xml:"w:val,attr"`
}

// TrackRevisionsXML represents track revisions setting
type TrackRevisionsXML struct{}

// TrackFormattingXML represents track formatting setting
type TrackFormattingXML struct{}

// DocumentProtectionXML represents document protection settings
type DocumentProtectionXML struct {
	Edit        string `xml:"w:edit,attr"`
	Enforcement string `xml:"w:enforcement,attr"`
	Hash        string `xml:"w:hash,attr,omitempty"`
	Salt        string `xml:"w:salt,attr,omitempty"`
}

// DocVarsXML represents document variables
type DocVarsXML struct {
	Variables []DocVarXML `xml:"w:docVar"`
}

// DocVarXML represents individual document variable
type DocVarXML struct {
	Name  string `xml:"w:name,attr"`
	Value string `xml:"w:val,attr"`
}

// MailMergeXML represents mail merge settings
type MailMergeXML struct {
	MainDocumentType *MainDocumentTypeXML `xml:"w:mainDocumentType,omitempty"`
	DataType         *DataTypeXML         `xml:"w:dataType,omitempty"`
	ConnectString    *ConnectStringXML    `xml:"w:connectString,omitempty"`
}

// MainDocumentTypeXML represents main document type for mail merge
type MainDocumentTypeXML struct {
	Val string `xml:"w:val,attr"`
}

// DataTypeXML represents data type for mail merge
type DataTypeXML struct {
	Val string `xml:"w:val,attr"`
}

// ConnectStringXML represents connection string for mail merge
type ConnectStringXML struct {
	Val string `xml:"w:val,attr"`
}

// DoNotPromptForConvertXML represents do not prompt for convert setting
type DoNotPromptForConvertXML struct{}

// DoNotAutoCompressPicturesXML represents do not auto compress pictures setting
type DoNotAutoCompressPicturesXML struct{}

// Additional XML types for grid and proofing settings
type DrawingGridHorizontalSpacingXML struct {
	Val int `xml:"w:val,attr"`
}

type DrawingGridVerticalSpacingXML struct {
	Val int `xml:"w:val,attr"`
}

type HideSpellingErrorsXML struct{}

type HideGrammaticalErrorsXML struct{}

// WebSettingsXML represents web settings
type WebSettingsXML struct {
	XMLName                          xml.Name                             `xml:"w:webSettings"`
	XmlnsW                           string                               `xml:"xmlns:w,attr"`
	OptimizeForBrowser               *OptimizeForBrowserXML               `xml:"w:optimizeForBrowser,omitempty"`
	DoNotUseHTMLParagraphAutoSpacing *DoNotUseHTMLParagraphAutoSpacingXML `xml:"w:doNotUseHTMLParagraphAutoSpacing,omitempty"`
	DoNotAutofit                     *DoNotAutofitXML                     `xml:"w:doNotAutofit,omitempty"`
}

// OptimizeForBrowserXML represents optimize for browser setting
type OptimizeForBrowserXML struct{}

// DoNotUseHTMLParagraphAutoSpacingXML represents HTML paragraph auto spacing setting
type DoNotUseHTMLParagraphAutoSpacingXML struct{}

// DoNotAutofitXML represents do not autofit setting
type DoNotAutofitXML struct{}

func NewSetting(document types.Document) *Settings {
	return &Settings{document: document}
}

// Path returns the location of the part inside the DOCX ZIP.
func (st *Settings) Path() string {
	return "word/settings.xml"
}

func (st *Settings) Byte() ([]byte, error) {
	// Get document settings
	docSettings := st.document.GetSettings()
	if docSettings == nil {
		// Use empty settings if none provided
		docSettings = &settings.DocumentSettings{}
	}

	// Create settings XML structure
	settingsXML := &SettingsXML{
		XmlnsW: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
	}

	// Map DocumentSettings to XML structures
	st.mapViewSettings(settingsXML, docSettings)
	st.mapFontSettings(settingsXML, docSettings)
	st.mapLanguageSettings(settingsXML, docSettings)
	st.mapCompatibilitySettings(settingsXML, docSettings)
	st.mapProtectionSettings(settingsXML, docSettings)
	st.mapTrackingSettings(settingsXML, docSettings)
	st.mapGridSettings(settingsXML, docSettings)
	st.mapProofingSettings(settingsXML, docSettings)
	st.mapMailMergeSettings(settingsXML, docSettings)
	st.mapDocumentVariables(settingsXML, docSettings)

	var buf bytes.Buffer

	// Write XML declaration
	buf.WriteString(xml.Header)

	// Encode the struct
	enc := xml.NewEncoder(&buf)
	enc.Indent("", "  ")

	if err := enc.Encode(settingsXML); err != nil {
		return nil, fmt.Errorf("encoding Settings XML: %w", err)
	}

	log.Printf("'%s' has been created.\n", st.Path())
	// Uncomment for debugging:
	// log.Print(buf.String())

	return buf.Bytes(), nil
}

// mapViewSettings maps view settings from DocumentSettings to XML
func (st *Settings) mapViewSettings(settingsXML *SettingsXML, docSettings *settings.DocumentSettings) {
	if docSettings.View == nil {
		return
	}

	view := docSettings.View

	// Zoom settings
	if view.ZoomPercent > 0 {
		settingsXML.Zoom = &ZoomXML{
			Percent: view.ZoomPercent,
			Val:     view.ZoomType,
		}
	}

	// Default view
	if view.DefaultView != "" && view.DefaultView != "print" {
		settingsXML.DefaultView = &ViewXML{
			Val: view.DefaultView,
		}
	}
}

// mapFontSettings maps font settings from DocumentSettings to XML
func (st *Settings) mapFontSettings(settingsXML *SettingsXML, docSettings *settings.DocumentSettings) {
	if docSettings.Font == nil {
		return
	}

	font := docSettings.Font

	// Default tab stop
	if font.DefaultTabStop > 0 {
		settingsXML.DefaultTabStop = &DefaultTabStopXML{
			Val: font.DefaultTabStop,
		}
	}

	// Character spacing control
	if font.CharacterSpacingControl != "" {
		settingsXML.CharacterSpacingControl = &CharacterSpacingControlXML{
			Val: font.CharacterSpacingControl,
		}
	}
}

// mapLanguageSettings maps language settings from DocumentSettings to XML
func (st *Settings) mapLanguageSettings(settingsXML *SettingsXML, docSettings *settings.DocumentSettings) {
	if docSettings.Language == nil {
		return
	}

	lang := docSettings.Language

	// Theme font language
	if lang.Default != "" || lang.EastAsian != "" || lang.BiDirectional != "" {
		settingsXML.ThemeFontLang = &ThemeFontLangXML{
			Val:      lang.Default,
			EastAsia: lang.EastAsian,
			Bidi:     lang.BiDirectional,
		}
	}
}

// mapCompatibilitySettings maps compatibility settings from DocumentSettings to XML
func (st *Settings) mapCompatibilitySettings(settingsXML *SettingsXML, docSettings *settings.DocumentSettings) {
	if docSettings.Compatibility == nil {
		return
	}

	compat := docSettings.Compatibility

	// Compatibility settings
	if compat.CompatibilityMode > 0 {
		settingsXML.Compat = &CompatXML{
			CompatSettings: []CompatSettingXML{
				{
					Name: "compatibilityMode",
					URI:  compat.URI,
					Val:  strconv.Itoa(compat.CompatibilityMode),
				},
			},
		}

		// Add additional compatibility settings
		if compat.AlignBordersAndEdges {
			settingsXML.Compat.CompatSettings = append(settingsXML.Compat.CompatSettings,
				CompatSettingXML{
					Name: "alignBordersAndEdges",
					URI:  compat.URI,
					Val:  "1",
				})
		}

		if compat.BordersDoNotSurroundHeader {
			settingsXML.Compat.CompatSettings = append(settingsXML.Compat.CompatSettings,
				CompatSettingXML{
					Name: "bordersDoNotSurroundHeader",
					URI:  compat.URI,
					Val:  "1",
				})
		}

		if compat.BordersDoNotSurroundFooter {
			settingsXML.Compat.CompatSettings = append(settingsXML.Compat.CompatSettings,
				CompatSettingXML{
					Name: "bordersDoNotSurroundFooter",
					URI:  compat.URI,
					Val:  "1",
				})
		}
	}
}

// mapProtectionSettings maps protection settings from DocumentSettings to XML
func (st *Settings) mapProtectionSettings(settingsXML *SettingsXML, docSettings *settings.DocumentSettings) {
	if docSettings.Protection == nil || !docSettings.Protection.EnforceProtection {
		return
	}

	protection := docSettings.Protection

	settingsXML.DocumentProtection = &DocumentProtectionXML{
		Edit:        protection.ProtectionType,
		Enforcement: "1",
		Hash:        protection.Password,
		Salt:        protection.Salt,
	}
}

// mapTrackingSettings maps tracking settings from DocumentSettings to XML
func (st *Settings) mapTrackingSettings(settingsXML *SettingsXML, docSettings *settings.DocumentSettings) {
	if docSettings.Tracking == nil {
		return
	}

	tracking := docSettings.Tracking

	// Track revisions
	if tracking.TrackRevisions {
		settingsXML.TrackRevisions = &TrackRevisionsXML{}
	}

	// Track formatting
	if tracking.TrackFormatting {
		settingsXML.TrackFormatting = &TrackFormattingXML{}
	}
}

// mapGridSettings maps grid settings from DocumentSettings to XML
func (st *Settings) mapGridSettings(settingsXML *SettingsXML, docSettings *settings.DocumentSettings) {
	if docSettings.Grid == nil {
		return
	}

	grid := docSettings.Grid

	// Drawing grid spacing
	if grid.DrawingGridHorizontalSpacing > 0 {
		settingsXML.DrawingGridHorizontalSpacing = &DrawingGridHorizontalSpacingXML{
			Val: grid.DrawingGridHorizontalSpacing,
		}
	}

	if grid.DrawingGridVerticalSpacing > 0 {
		settingsXML.DrawingGridVerticalSpacing = &DrawingGridVerticalSpacingXML{
			Val: grid.DrawingGridVerticalSpacing,
		}
	}
}

// mapProofingSettings maps proofing settings from DocumentSettings to XML
func (st *Settings) mapProofingSettings(settingsXML *SettingsXML, docSettings *settings.DocumentSettings) {
	if docSettings.Proofing == nil {
		return
	}

	proofing := docSettings.Proofing

	// Proof state
	if proofing.SpellingState != "" || proofing.GrammarState != "" {
		settingsXML.ProofState = &ProofStateXML{
			Spelling: proofing.SpellingState,
			Grammar:  proofing.GrammarState,
		}
	}

	// Hide errors
	if proofing.HideSpellingErrors {
		settingsXML.HideSpellingErrors = &HideSpellingErrorsXML{}
	}

	if proofing.HideGrammaticalErrors {
		settingsXML.HideGrammaticalErrors = &HideGrammaticalErrorsXML{}
	}

	// Additional proofing settings
	if proofing.DoNotPromptForConvert {
		settingsXML.DoNotPromptForConvert = &DoNotPromptForConvertXML{}
	}

	if proofing.DoNotAutoCompressPictures {
		settingsXML.DoNotAutoCompressPictures = &DoNotAutoCompressPicturesXML{}
	}
}

// mapMailMergeSettings maps mail merge settings from DocumentSettings to XML
func (st *Settings) mapMailMergeSettings(settingsXML *SettingsXML, docSettings *settings.DocumentSettings) {
	if docSettings.MailMerge == nil {
		return
	}

	mailMerge := docSettings.MailMerge
	xmlMailMerge := &MailMergeXML{}

	if mailMerge.MainDocumentType != "" {
		xmlMailMerge.MainDocumentType = &MainDocumentTypeXML{
			Val: mailMerge.MainDocumentType,
		}
	}

	if mailMerge.DataType != "" {
		xmlMailMerge.DataType = &DataTypeXML{
			Val: mailMerge.DataType,
		}
	}

	if mailMerge.ConnectString != "" {
		xmlMailMerge.ConnectString = &ConnectStringXML{
			Val: mailMerge.ConnectString,
		}
	}

	// Only add if at least one field is set
	if xmlMailMerge.MainDocumentType != nil || xmlMailMerge.DataType != nil || xmlMailMerge.ConnectString != nil {
		settingsXML.MailMerge = xmlMailMerge
	}
}

// mapDocumentVariables maps document variables from DocumentSettings to XML
func (st *Settings) mapDocumentVariables(settingsXML *SettingsXML, docSettings *settings.DocumentSettings) {
	if len(docSettings.Variables) == 0 {
		return
	}

	docVars := &DocVarsXML{}
	for name, value := range docSettings.Variables {
		docVars.Variables = append(docVars.Variables, DocVarXML{
			Name:  name,
			Value: value,
		})
	}

	settingsXML.DocVars = docVars
}

// WriteTo writes the XML content to the given writer.
func (st *Settings) WriteTo(w io.Writer) (int64, error) {
	data, err := st.Byte()
	if err != nil {
		return 0, err
	}

	n, err := w.Write(data)
	return int64(n), err
}
