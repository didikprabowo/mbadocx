package types

import (
	contenttypes "github.com/didikprabowo/mbadocx/content_types"
	"github.com/didikprabowo/mbadocx/metadata"
	"github.com/didikprabowo/mbadocx/relationships"
	"github.com/didikprabowo/mbadocx/styles"
)

// DocumentInterface provides access to document data for the writer
type Document interface {
	Body() Body
	Relationships() Relationships
	Metadata() Metadata
	Styles() Styles
	ContentTypes() ContentTypes
}

type ContentTypes interface {
	Get() *contenttypes.ContentTypes
}
type Styles interface {
	Get() *styles.Styles
}

type Metadata interface {
	Get() *metadata.Metadata
}

// BodyInterface provides access to body elements
type Body interface {
	GetElements() []Element
}

// Element interface for all document elements
type Element interface {
	Type() string
	XML() ([]byte, error)
}

type Relationships interface {
	PackageXML() ([]byte, error)
	DocumentXML() ([]byte, error)
	GetOrCreateHyperlink(url string) *relationships.Relationship
}
