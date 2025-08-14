package types

import (
	"time"

	"github.com/didikprabowo/mbadocx/relationships"
	"github.com/didikprabowo/mbadocx/settings"
)

// DocumentInterface provides access to document data for the writer
type Document interface {
	GetBody() Body
	GetRelationships() Relationships
	GetMetadata() *Metadata
	GetSettings() *settings.DocumentSettings
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

type Relationship struct {
	ID         string
	Type       string
	Target     string
	TargetMode string
}

// Metadata contains document metadata
type Metadata struct {
	Title          string
	Subject        string
	Creator        string
	Keywords       string
	Description    string
	LastModifiedBy string
	Revision       string
	Created        time.Time
	Modified       time.Time
	Category       string
	ContentStatus  string
	Language       string
	Version        string
	Company        string
	Manager        string
}
