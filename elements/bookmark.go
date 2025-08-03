package elements

import (
	"bytes"
	"fmt"

	"github.com/google/uuid"
)

// Bookmark represents a bookmark in the document
type Bookmark struct {
	ID   string
	Name string
}

// BookmarkStart represents the start of a bookmark
type BookmarkStart struct {
	ID   string
	Name string
}

// BookmarkEnd represents the end of a bookmark
type BookmarkEnd struct {
	ID string
}

// NewBookmark creates a new bookmark with start and end markers
func NewBookmark(name string) *Bookmark {
	return &Bookmark{
		ID:   generateBookmarkID(),
		Name: name,
	}
}

// NewBookmarkStart creates a bookmark start marker
func NewBookmarkStart(id, name string) *BookmarkStart {
	return &BookmarkStart{
		ID:   id,
		Name: name,
	}
}

// NewBookmarkEnd creates a bookmark end marker
func NewBookmarkEnd(id string) *BookmarkEnd {
	return &BookmarkEnd{
		ID: id,
	}
}

// Type returns the element type
func (b *Bookmark) Type() string {
	return "bookmark"
}

// Type returns the element type
func (bs *BookmarkStart) Type() string {
	return "bookmarkStart"
}

// Type returns the element type
func (be *BookmarkEnd) Type() string {
	return "bookmarkEnd"
}

// GetStart returns the bookmark start element
func (b *Bookmark) GetStart() *BookmarkStart {
	return &BookmarkStart{
		ID:   b.ID,
		Name: b.Name,
	}
}

// GetEnd returns the bookmark end element
func (b *Bookmark) GetEnd() *BookmarkEnd {
	return &BookmarkEnd{
		ID: b.ID,
	}
}

// XML generates the XML for bookmark (both start and end)
func (b *Bookmark) XML() ([]byte, error) {
	var buf bytes.Buffer

	// Bookmark start
	buf.WriteString(fmt.Sprintf(`<w:bookmarkStart w:id="%s" w:name="%s"/>`, b.ID, b.Name))

	// Bookmark end
	buf.WriteString(fmt.Sprintf(`<w:bookmarkEnd w:id="%s"/>`, b.ID))

	return buf.Bytes(), nil
}

// XML generates the XML for bookmark start
func (bs *BookmarkStart) XML() ([]byte, error) {
	return []byte(fmt.Sprintf(`<w:bookmarkStart w:id="%s" w:name="%s"/>`, bs.ID, bs.Name)), nil
}

// XML generates the XML for bookmark end
func (be *BookmarkEnd) XML() ([]byte, error) {
	return []byte(fmt.Sprintf(`<w:bookmarkEnd w:id="%s"/>`, be.ID)), nil
}

// generateBookmarkID generates a unique bookmark ID
func generateBookmarkID() string {
	// Bookmark IDs are typically numeric
	return fmt.Sprintf("%d", uuid.New().ID())
}
