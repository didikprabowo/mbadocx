package metadata

import (
	"time"
)

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

// NewDefaultMetadata creates default metadata
func NewDefaultMetadata() *Metadata {
	return &Metadata{
		Creator:  "Go DOCX Library",
		Created:  time.Now(),
		Modified: time.Now(),
		Revision: "1",
		Language: "en-US",
	}
}

func (md *Metadata) Get() *Metadata {
	return md
}
