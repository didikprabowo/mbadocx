package mbadocx

import (
	"github.com/didikprabowo/mbadocx/elements"
	"github.com/didikprabowo/mbadocx/types"
)

type Media struct {
	Media []types.Media
}

// AddMedia
func (m *Media) AddMedia(img *elements.Image) {
	m.Media = append(m.Media, img)
}
