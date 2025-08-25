// elements/image.go
package elements

import (
	"bytes"
	"encoding/base64"
	"fmt"
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"os"
	"path/filepath"
	"strings"

	"github.com/didikprabowo/mbadocx/properties"
	"github.com/didikprabowo/mbadocx/relationships"
	"github.com/didikprabowo/mbadocx/types"
)

var (
	_ types.Element = (*Image)(nil)
	_ types.Media   = (*Image)(nil)
)

// Image represents an image element in a Word document
type Image struct {
	document       types.Document
	RelationshipID string
	Width          int64 // Width in EMUs (English Metric Units)
	Height         int64 // Height in EMUs
	Name           string
	Description    string
	Data           []byte
	ContentType    string
	Extension      string
	props          properties.ImageProperties
}

const (
	// EMUs per inch
	EmusPerInch = 914400
	// EMUs per centimeter
	EmusPerCm = 360000
	// EMUs per pixel at 96 DPI
	EmusPerPixel = 9525
)

// Image content types
const (
	ContentTypeJPEG = "image/jpeg"
	ContentTypePNG  = "image/png"
	ContentTypeGIF  = "image/gif"
	ContentTypeBMP  = "image/bmp"
	ContentTypeTIFF = "image/tiff"
	ContentTypeSVG  = "image/svg+xml"
)

// NewImage creates a new image from file path
func NewImage(document types.Document, filePath string) (*Image, error) {
	// Read file
	data, err := os.ReadFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to read image file: %w", err)
	}

	// Get file extension
	ext := strings.ToLower(filepath.Ext(filePath))
	if ext == "" {
		return nil, fmt.Errorf("image file must have an extension")
	}
	ext = strings.TrimPrefix(ext, ".")

	// Determine content type
	contentType := getContentType(ext)
	if contentType == "" {
		return nil, fmt.Errorf("unsupported image format: %s", ext)
	}

	// Get image dimensions
	width, height, err := getImageDimensions(data)
	if err != nil {
		return nil, fmt.Errorf("failed to get image dimensions: %w", err)
	}

	// Create image with default properties
	img := &Image{
		document:    document,
		Name:        filepath.Base(filePath),
		Description: fmt.Sprintf("Image: %s", filepath.Base(filePath)),
		Data:        data,
		ContentType: contentType,
		Extension:   ext,
		Width:       int64(width) * EmusPerPixel,
		Height:      int64(height) * EmusPerPixel,
		props:       *properties.NewImageProperties(),
	}

	// Register with relationships
	if document != nil {
		rel := document.Relationships().AddImage(img.Name)
		img.RelationshipID = rel.ID
	}

	return img, nil
}

// NewImageFromBytes creates a new image from byte data
func NewImageFromBytes(document types.Document, data []byte, name string, contentType string) (*Image, error) {
	// Get image dimensions
	width, height, err := getImageDimensions(data)
	if err != nil {
		return nil, fmt.Errorf("failed to get image dimensions: %w", err)
	}

	// Determine extension from content type
	ext := getExtensionFromContentType(contentType)
	if ext == "" {
		return nil, fmt.Errorf("unsupported content type: %s", contentType)
	}

	img := &Image{
		document:    document,
		Name:        name,
		Description: fmt.Sprintf("Image: %s", name),
		Data:        data,
		ContentType: contentType,
		Extension:   ext,
		Width:       int64(width) * EmusPerPixel,
		Height:      int64(height) * EmusPerPixel,
		props:       *properties.NewImageProperties(),
	}

	// Register with relationships
	if document != nil {
		rel := document.Relationships().AddImage(img.Name + "." + img.Extension)
		img.RelationshipID = rel.ID
	}

	return img, nil
}

// NewImageFromReader creates a new image from an io.Reader
func NewImageFromReader(document types.Document, reader io.Reader, name string) (*Image, error) {
	data, err := io.ReadAll(reader)
	if err != nil {
		return nil, fmt.Errorf("failed to read image data: %w", err)
	}

	// Detect content type from data
	contentType := detectContentType(data)
	if contentType == "" {
		return nil, fmt.Errorf("unable to detect image format")
	}

	return NewImageFromBytes(document, data, name, contentType)
}

// Type returns the element type
func (img *Image) Type() string {
	return "image"
}

// RelID returns the relationship ID
func (img *Image) RelID() string {
	return img.RelationshipID
}

// RelType returns the relationship type
func (img *Image) RelType() string {
	return relationships.TypeImage
}

// TargetPath returns the target path for the image
func (img *Image) TargetPath() string {
	return "word/media/"
}

// FileName returns the file name
func (img *Image) FileName() string {
	return img.Name
}

// RawContent returns the raw image data
func (img *Image) RawContent() []byte {
	return img.Data
}

// SetSize sets the image size in inches
func (img *Image) SetSize(widthInches, heightInches float64) *Image {
	img.Width = int64(widthInches * float64(EmusPerInch))
	img.Height = int64(heightInches * float64(EmusPerInch))
	return img
}

// SetSizeInCm sets the image size in centimeters
func (img *Image) SetSizeInCm(widthCm, heightCm float64) *Image {
	img.Width = int64(widthCm * float64(EmusPerCm))
	img.Height = int64(heightCm * float64(EmusPerCm))
	return img
}

// SetSizeInPixels sets the image size in pixels (at 96 DPI)
func (img *Image) SetSizeInPixels(widthPx, heightPx int) *Image {
	img.Width = int64(widthPx) * EmusPerPixel
	img.Height = int64(heightPx) * EmusPerPixel
	return img
}

// ScaleToWidth scales the image to a specific width while maintaining aspect ratio
func (img *Image) ScaleToWidth(widthInches float64) *Image {
	targetWidth := int64(widthInches * float64(EmusPerInch))
	ratio := float64(targetWidth) / float64(img.Width)
	img.Width = targetWidth
	img.Height = int64(float64(img.Height) * ratio)
	return img
}

// ScaleToHeight scales the image to a specific height while maintaining aspect ratio
func (img *Image) ScaleToHeight(heightInches float64) *Image {
	targetHeight := int64(heightInches * float64(EmusPerInch))
	ratio := float64(targetHeight) / float64(img.Height)
	img.Height = targetHeight
	img.Width = int64(float64(img.Width) * ratio)
	return img
}

// Scale scales the image by a percentage (e.g., 0.5 for 50%)
func (img *Image) Scale(factor float64) *Image {
	img.Width = int64(float64(img.Width) * factor)
	img.Height = int64(float64(img.Height) * factor)
	return img
}

// FitToWidth scales the image to fit within a maximum width while maintaining aspect ratio
func (img *Image) FitToWidth(maxWidthInches float64) *Image {
	maxWidth := int64(maxWidthInches * float64(EmusPerInch))
	if img.Width > maxWidth {
		ratio := float64(maxWidth) / float64(img.Width)
		img.Width = maxWidth
		img.Height = int64(float64(img.Height) * ratio)
	}
	return img
}

// FitToHeight scales the image to fit within a maximum height while maintaining aspect ratio
func (img *Image) FitToHeight(maxHeightInches float64) *Image {
	maxHeight := int64(maxHeightInches * float64(EmusPerInch))
	if img.Height > maxHeight {
		ratio := float64(maxHeight) / float64(img.Height)
		img.Height = maxHeight
		img.Width = int64(float64(img.Width) * ratio)
	}
	return img
}

// FitToBox scales the image to fit within a bounding box while maintaining aspect ratio
func (img *Image) FitToBox(maxWidthInches, maxHeightInches float64) *Image {
	maxWidth := int64(maxWidthInches * float64(EmusPerInch))
	maxHeight := int64(maxHeightInches * float64(EmusPerInch))

	widthRatio := float64(maxWidth) / float64(img.Width)
	heightRatio := float64(maxHeight) / float64(img.Height)

	// Use the smaller ratio to ensure the image fits within both dimensions
	ratio := widthRatio
	if heightRatio < widthRatio {
		ratio = heightRatio
	}

	if ratio < 1 {
		img.Width = int64(float64(img.Width) * ratio)
		img.Height = int64(float64(img.Height) * ratio)
	}

	return img
}

// SetProperties sets the image properties
func (img *Image) SetProperties(props properties.ImageProperties) *Image {
	img.props = props
	return img
}

// SetAlignment sets the horizontal alignment for inline images
func (img *Image) SetAlignment(align properties.ImageAlignment) *Image {
	img.props.Alignment = align
	return img
}

// SetWrapStyle sets how text wraps around the image
func (img *Image) SetWrapStyle(wrap properties.WrapStyle) *Image {
	img.props.WrapType = wrap
	if wrap != properties.WrapInline {
		img.props.Inline = false
	}
	return img
}

// SetBorder adds a border to the image
func (img *Image) SetBorder(width int, color string) *Image {
	img.props.BorderWidth = width
	img.props.BorderColor = strings.TrimPrefix(color, "#")
	img.props.BorderStyle = "solid"
	return img
}

// SetShadow adds a shadow effect to the image
func (img *Image) SetShadow(enabled bool) *Image {
	img.props.Shadow = enabled
	if enabled && img.props.ShadowColor == "" {
		img.props.ShadowColor = "000000" // Default to black
		img.props.ShadowOffset = 3       // Default offset
	}
	return img
}

// SetRotation sets the rotation angle in degrees
func (img *Image) SetRotation(degrees int) *Image {
	img.props.Rotation = degrees % 360
	return img
}

// SetCropping sets the crop percentages for each side
func (img *Image) SetCropping(top, bottom, left, right float64) *Image {
	img.props.SetCrop(top, bottom, left, right)
	return img
}

// SetTransparency sets the image transparency (0 to 100)
func (img *Image) SetTransparency(transparency float64) *Image {
	img.props.Transparency = clamp(transparency, 0, 100)
	return img
}

// SetAltText sets the alternative text for accessibility
func (img *Image) SetAltText(text string) *Image {
	img.props.AltText = text
	if img.Description == "" || img.Description == fmt.Sprintf("Image: %s", img.Name) {
		img.Description = text
	}
	return img
}

// SetFloating makes the image float with specified anchoring
func (img *Image) SetFloating(hAnchor properties.HorizontalAnchor, vAnchor properties.VerticalAnchor) *Image {
	img.props.Inline = false
	img.props.HorizontalPosition = hAnchor
	img.props.VerticalPosition = vAnchor
	return img
}

// SetOffset sets the offset for floating images in EMUs
func (img *Image) SetOffset(horizontal, vertical int64) *Image {
	img.props.HorizontalOffset = horizontal
	img.props.VerticalOffset = vertical
	return img
}

// SetOffsetInches sets the offset for floating images in inches
func (img *Image) SetOffsetInches(horizontal, vertical float64) *Image {
	img.props.HorizontalOffset = int64(horizontal * float64(EmusPerInch))
	img.props.VerticalOffset = int64(vertical * float64(EmusPerInch))
	return img
}

// XML generates the XML representation of the image
func (img *Image) XML() ([]byte, error) {
	// Validate properties first
	if err := img.props.Validate(); err != nil {
		return nil, fmt.Errorf("invalid image properties: %w", err)
	}

	var buf bytes.Buffer

	// Generate unique IDs
	docPrID := generateID()
	picID := generateID()

	// Start drawing
	buf.WriteString(`<w:drawing>`)

	if img.props.Inline {
		// Inline image
		buf.WriteString(`<wp:inline distT="0" distB="0" distL="0" distR="0">`)

		// Size
		buf.WriteString(fmt.Sprintf(`<wp:extent cx="%d" cy="%d"/>`, img.Width, img.Height))

		// Effect extent
		buf.WriteString(`<wp:effectExtent l="0" t="0" r="0" b="0"/>`)

		// Document properties
		buf.WriteString(fmt.Sprintf(`<wp:docPr id="%d" name="%s" descr="%s"`,
			docPrID, img.Name, img.props.AltText))
		if img.props.AltText != "" {
			buf.WriteString(fmt.Sprintf(` title="%s"`, img.props.AltText))
		}
		buf.WriteString(`/>`)

		// Non-visual graphic properties
		buf.WriteString(`<wp:cNvGraphicFramePr>`)
		buf.WriteString(fmt.Sprintf(`<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="%s"/>`,
			boolToString(img.props.LockAspectRatio)))
		buf.WriteString(`</wp:cNvGraphicFramePr>`)
	} else {
		// Floating image - use anchor
		buf.WriteString(img.props.GeneratePositionXML())

		// Size
		buf.WriteString(fmt.Sprintf(`<wp:extent cx="%d" cy="%d"/>`, img.Width, img.Height))

		// Effect extent
		buf.WriteString(`<wp:effectExtent l="0" t="0" r="0" b="0"/>`)

		// Wrap settings
		wrapXML := img.props.GenerateWrapXML()
		if wrapXML != "" {
			buf.WriteString(wrapXML)
		}

		// Document properties
		buf.WriteString(fmt.Sprintf(`<wp:docPr id="%d" name="%s" descr="%s"`,
			docPrID, img.Name, img.props.AltText))
		if img.props.AltText != "" {
			buf.WriteString(fmt.Sprintf(` title="%s"`, img.props.AltText))
		}
		buf.WriteString(`/>`)

		// Non-visual graphic properties
		buf.WriteString(`<wp:cNvGraphicFramePr>`)
		buf.WriteString(fmt.Sprintf(`<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="%s"/>`,
			boolToString(img.props.LockAspectRatio)))
		buf.WriteString(`</wp:cNvGraphicFramePr>`)
	}

	// Graphic
	buf.WriteString(`<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">`)
	buf.WriteString(`<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">`)

	// Picture
	buf.WriteString(`<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">`)

	// Non-visual picture properties
	buf.WriteString(`<pic:nvPicPr>`)
	buf.WriteString(fmt.Sprintf(`<pic:cNvPr id="%d" name="%s" descr="%s"`,
		picID, img.Name, img.props.AltText))
	if img.props.AltText != "" {
		buf.WriteString(fmt.Sprintf(` title="%s"`, img.props.AltText))
	}
	buf.WriteString(`/>`)
	buf.WriteString(`<pic:cNvPicPr>`)
	buf.WriteString(fmt.Sprintf(`<a:picLocks noChangeAspect="%s"/>`, boolToString(img.props.LockAspectRatio)))
	buf.WriteString(`</pic:cNvPicPr>`)
	buf.WriteString(`</pic:nvPicPr>`)

	// Blip (image reference) with effects
	buf.WriteString(`<pic:blipFill>`)
	buf.WriteString(fmt.Sprintf(`<a:blip r:embed="%s">`, img.RelationshipID))

	// Add image adjustments if any
	adjustmentsXML := img.props.GenerateImageAdjustmentsXML()
	if adjustmentsXML != "" {
		buf.WriteString(adjustmentsXML)
	}

	buf.WriteString(`</a:blip>`)

	// Add cropping if specified
	cropXML := img.props.GenerateCropXML()
	if cropXML != "" {
		buf.WriteString(cropXML)
	}

	buf.WriteString(`<a:stretch>`)
	buf.WriteString(`<a:fillRect/>`)
	buf.WriteString(`</a:stretch>`)
	buf.WriteString(`</pic:blipFill>`)

	// Shape properties
	buf.WriteString(`<pic:spPr>`)

	// Transform (including rotation)
	buf.WriteString(img.props.GenerateTransformXML(img.Width, img.Height))

	// Preset geometry
	buf.WriteString(`<a:prstGeom prst="rect">`)
	buf.WriteString(`<a:avLst/>`)
	buf.WriteString(`</a:prstGeom>`)

	// Effects (border and shadow)
	effectsXML := img.props.GenerateEffectsXML()
	if effectsXML != "" {
		buf.WriteString(effectsXML)
	}

	buf.WriteString(`</pic:spPr>`)

	buf.WriteString(`</pic:pic>`)
	buf.WriteString(`</a:graphicData>`)
	buf.WriteString(`</a:graphic>`)

	// Close inline or anchor tag
	if img.props.Inline {
		buf.WriteString(`</wp:inline>`)
	} else {
		buf.WriteString(`</wp:anchor>`)
	}

	buf.WriteString(`</w:drawing>`)

	return buf.Bytes(), nil
}

// GetBase64Data returns the image data as base64 encoded string
func (img *Image) GetBase64Data() string {
	return base64.StdEncoding.EncodeToString(img.Data)
}

// SaveToFile saves the image data to a file
func (img *Image) SaveToFile(path string) error {
	return os.WriteFile(path, img.Data, 0644)
}

// GetDimensionsInInches returns the image dimensions in inches
func (img *Image) GetDimensionsInInches() (width, height float64) {
	width = float64(img.Width) / float64(EmusPerInch)
	height = float64(img.Height) / float64(EmusPerInch)
	return
}

// GetDimensionsInCm returns the image dimensions in centimeters
func (img *Image) GetDimensionsInCm() (width, height float64) {
	width = float64(img.Width) / float64(EmusPerCm)
	height = float64(img.Height) / float64(EmusPerCm)
	return
}

// GetDimensionsInPixels returns the image dimensions in pixels (at 96 DPI)
func (img *Image) GetDimensionsInPixels() (width, height int) {
	width = int(img.Width / EmusPerPixel)
	height = int(img.Height / EmusPerPixel)
	return
}

// GetAspectRatio returns the aspect ratio of the image
func (img *Image) GetAspectRatio() float64 {
	if img.Height == 0 {
		return 0
	}
	return float64(img.Width) / float64(img.Height)
}

// Clone creates a deep copy of the image
func (img *Image) Clone() *Image {
	dataCopy := make([]byte, len(img.Data))
	copy(dataCopy, img.Data)

	return &Image{
		document:       img.document,
		RelationshipID: img.RelationshipID,
		Width:          img.Width,
		Height:         img.Height,
		Name:           img.Name,
		Description:    img.Description,
		Data:           dataCopy,
		ContentType:    img.ContentType,
		Extension:      img.Extension,
		props:          img.props,
	}
}

// Helper functions

func getContentType(ext string) string {
	switch ext {
	case "jpg", "jpeg":
		return ContentTypeJPEG
	case "png":
		return ContentTypePNG
	case "gif":
		return ContentTypeGIF
	case "bmp":
		return ContentTypeBMP
	case "tiff", "tif":
		return ContentTypeTIFF
	case "svg":
		return ContentTypeSVG
	default:
		return ""
	}
}

func getExtensionFromContentType(contentType string) string {
	switch contentType {
	case ContentTypeJPEG:
		return "jpg"
	case ContentTypePNG:
		return "png"
	case ContentTypeGIF:
		return "gif"
	case ContentTypeBMP:
		return "bmp"
	case ContentTypeTIFF:
		return "tiff"
	case ContentTypeSVG:
		return "svg"
	default:
		return ""
	}
}

func detectContentType(data []byte) string {
	if len(data) < 4 {
		return ""
	}

	// Check magic numbers
	if bytes.HasPrefix(data, []byte{0xFF, 0xD8, 0xFF}) {
		return ContentTypeJPEG
	}
	if bytes.HasPrefix(data, []byte{0x89, 0x50, 0x4E, 0x47}) {
		return ContentTypePNG
	}
	if bytes.HasPrefix(data, []byte("GIF87a")) || bytes.HasPrefix(data, []byte("GIF89a")) {
		return ContentTypeGIF
	}
	if bytes.HasPrefix(data, []byte("BM")) {
		return ContentTypeBMP
	}

	return ""
}

func getImageDimensions(data []byte) (width, height int, err error) {
	reader := bytes.NewReader(data)
	config, _, err := image.DecodeConfig(reader)
	if err != nil {
		return 0, 0, err
	}
	return config.Width, config.Height, nil
}

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

var idCounter int64 = 1

func generateID() int64 {
	id := idCounter
	idCounter++
	return id
}
