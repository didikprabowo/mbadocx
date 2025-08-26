package mbadocx

import (
	"github.com/didikprabowo/mbadocx/elements"
)

// AddImage inserts an image into the document from the specified file path.
// The image is automatically embedded in the document and wrapped in a paragraph element.
//
// This method handles the complete image insertion process:
//  1. Creates a new Image element from the file
//  2. Wraps it in a paragraph for proper document structure
//  3. Registers it with the document's media manager
//  4. Returns the Image element for further manipulation (sizing, positioning, etc.)
//
// Parameters:
//   - imagePath: The file system path to the image file.
//     Supports common formats: JPEG, PNG, GIF, BMP, TIFF, SVG
//
// Returns:
//   - *elements.Image: A pointer to the created Image element that can be used
//     to set properties like width, height, alignment, etc.
//   - error: An error if the image file cannot be read, is in an unsupported format,
//     or if there are issues creating the image element
//
// Example:
//
//	doc := mbadocx.New()
//
//	// Add an image with error handling
//	img, err := doc.AddImage("./assets/logo.png")
//	if err != nil {
//	    log.Fatalf("Failed to add image: %v", err)
//	}
//
//	// Optionally configure the image after adding
//	img.SetWidth(300).SetHeight(200)
//	img.SetAlignment(elements.AlignCenter)
//
// Note: The image is embedded in the DOCX file, so the original image file
// is not needed once the document is saved. The image data is stored in the
// document's media collection and referenced via relationships.
//
// Common error scenarios:
//   - File not found at the specified path
//   - Insufficient permissions to read the file
//   - Unsupported or corrupted image format
//   - Out of memory for very large images
func (d *Document) AddImage(imagePath string) (*elements.Image, error) {
	// Create a new Image element from the file path
	// This validates the file exists and reads its contents
	img, err := elements.NewImage(d, imagePath)
	if err != nil {
		// Return early if image creation fails, maintaining clean error handling
		return nil, err
	}

	// Create a paragraph container for the image
	// Word documents require images to be contained within paragraph elements
	p := elements.NewParagraph(d)

	// Add the image as a child of the paragraph
	// This establishes the parent-child relationship in the document tree
	p.AddChildren(img)

	// Add the paragraph (containing the image) to the document body
	// This makes the image visible in the document flow
	d.body.AddElement(p)

	// Register the image with the document's media manager
	// This handles the internal DOCX packaging and relationship management
	d.media.AddMedia(img)

	// Return the image element for optional further configuration
	return img, nil
}
