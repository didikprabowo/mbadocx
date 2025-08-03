package elements

import "fmt"

type CommentRangeStart struct {
	ID string
}

// CommentRangeEnd marks the end of a commented range
type CommentRangeEnd struct {
	ID string
}

// CommentReference references a comment
type CommentReference struct {
	ID string
}

// Comment represents a comment (stored separately in comments.xml)
type Comment struct {
	ID       string
	Author   string
	Date     string
	Initials string
	Text     string
}

// NewCommentRangeStart creates a comment range start marker
func NewCommentRangeStart(id string) *CommentRangeStart {
	return &CommentRangeStart{
		ID: id,
	}
}

// NewCommentRangeEnd creates a comment range end marker
func NewCommentRangeEnd(id string) *CommentRangeEnd {
	return &CommentRangeEnd{
		ID: id,
	}
}

// NewCommentReference creates a comment reference
func NewCommentReference(id string) *CommentReference {
	return &CommentReference{
		ID: id,
	}
}

// NewComment creates a new comment
func NewComment(id, author, text string) *Comment {
	return &Comment{
		ID:     id,
		Author: author,
		Text:   text,
		Date:   "", // Will be set when saving
	}
}

// Type returns the element type
func (c *CommentRangeStart) Type() string {
	return "commentRangeStart"
}

// Type returns the element type
func (c *CommentRangeEnd) Type() string {
	return "commentRangeEnd"
}

// Type returns the element type
func (c *CommentReference) Type() string {
	return "commentReference"
}

// XML generates the XML for comment range start
func (c *CommentRangeStart) XML() ([]byte, error) {
	return []byte(fmt.Sprintf(`<w:commentRangeStart w:id="%s"/>`, c.ID)), nil
}

// XML generates the XML for comment range end
func (c *CommentRangeEnd) XML() ([]byte, error) {
	return []byte(fmt.Sprintf(`<w:commentRangeEnd w:id="%s"/>`, c.ID)), nil
}

// XML generates the XML for comment reference
func (c *CommentReference) XML() ([]byte, error) {
	return []byte(fmt.Sprintf(`<w:r><w:commentReference w:id="%s"/></w:r>`, c.ID)), nil
}
