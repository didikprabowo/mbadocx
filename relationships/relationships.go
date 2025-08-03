// File: relationships/relationships.go
package relationships

import (
	"encoding/xml"
	"fmt"
	"path"
	"strings"

	"github.com/google/uuid"
)

// Relationships manages all document relationships
type Relationships struct {
	items        map[string]*Relationship
	counter      int
	nextID       int
	byType       map[string][]*Relationship
	byTarget     map[string]*Relationship
	external     map[string]*Relationship
	packageRels  []*Relationship // Package-level relationships (_rels/.rels)
	documentRels []*Relationship // Document relationships (word/_rels/document.xml.rels)
}

// Relationship represents a single relationship
type Relationship struct {
	ID         string
	Type       string
	Target     string
	TargetMode string // Internal or External
	TargetKey  string // Unique key for lookups
}

// Relationship types constants
const (
	// Document relationships
	TypeDocument       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	TypeStyles         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
	TypeNumbering      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
	TypeSettings       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
	TypeWebSettings    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
	TypeFontTable      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
	TypeTheme          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
	TypeFootnotes      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
	TypeEndnotes       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
	TypeComments       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
	TypeHeader         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
	TypeFooter         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
	TypeImage          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
	TypeHyperlink      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
	TypeChart          = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
	TypeDiagram        = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
	TypeCustomXML      = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
	TypeCustomXMLProps = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"

	// Package relationships
	TypeCoreProperties   = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
	TypeExtProperties    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
	TypeCustomProperties = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
	TypeThumbnail        = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"
	TypeDigitalSignature = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"

	// Target modes
	TargetModeInternal = "Internal"
	TargetModeExternal = "External"
)

// RelationshipsXML represents the XML structure
type RelationshipsXML struct {
	XMLName       xml.Name          `xml:"Relationships"`
	Xmlns         string            `xml:"xmlns,attr"`
	Relationships []RelationshipXML `xml:"Relationship"`
}

// RelationshipXML represents a relationship in XML
type RelationshipXML struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

// New creates a new Relationships manager
func New() *Relationships {
	return &Relationships{
		items:        make(map[string]*Relationship),
		counter:      0,
		nextID:       1,
		byType:       make(map[string][]*Relationship),
		byTarget:     make(map[string]*Relationship),
		external:     make(map[string]*Relationship),
		packageRels:  make([]*Relationship, 0),
		documentRels: make([]*Relationship, 0),
	}
}

// NewDefault creates relationships with default entries
func NewDefault() *Relationships {
	rels := New()

	// Add default package relationships
	rels.AddPackageRelationship(TypeDocument, "word/document.xml", TargetModeInternal)
	rels.AddPackageRelationship(TypeCoreProperties, "docProps/core.xml", TargetModeInternal)
	rels.AddPackageRelationship(TypeExtProperties, "docProps/app.xml", TargetModeInternal)

	// Add default document relationships
	rels.AddDocumentRelationship(TypeStyles, "styles.xml", TargetModeInternal)
	rels.AddDocumentRelationship(TypeSettings, "settings.xml", TargetModeInternal)
	rels.AddDocumentRelationship(TypeWebSettings, "webSettings.xml", TargetModeInternal)
	rels.AddDocumentRelationship(TypeFontTable, "fontTable.xml", TargetModeInternal)
	rels.AddDocumentRelationship(TypeTheme, "theme/theme1.xml", TargetModeInternal)

	return rels
}

// generateID generates a new relationship ID
func (r *Relationships) generateID() string {
	id := fmt.Sprintf("rId%d", r.nextID)
	r.nextID++
	return id
}

// AddRelationship adds a generic relationship
func (r *Relationships) AddRelationship(relType, target, targetMode string) *Relationship {
	rel := &Relationship{
		ID:         r.generateID(),
		Type:       relType,
		Target:     target,
		TargetMode: targetMode,
	}

	// Generate unique key
	if targetMode == TargetModeExternal {
		rel.TargetKey = target
	} else {
		rel.TargetKey = path.Join(relType, target)
	}

	r.items[rel.ID] = rel
	r.byType[relType] = append(r.byType[relType], rel)
	r.byTarget[rel.TargetKey] = rel

	if targetMode == TargetModeExternal {
		r.external[target] = rel
	}

	return rel
}

// AddPackageRelationship adds a package-level relationship
func (r *Relationships) AddPackageRelationship(relType, target, targetMode string) *Relationship {
	rel := r.AddRelationship(relType, target, targetMode)
	r.packageRels = append(r.packageRels, rel)
	return rel
}

// AddDocumentRelationship adds a document-level relationship
func (r *Relationships) AddDocumentRelationship(relType, target, targetMode string) *Relationship {
	rel := r.AddRelationship(relType, target, targetMode)
	r.documentRels = append(r.documentRels, rel)
	return rel
}

// AddImage adds an image relationship
func (r *Relationships) AddImage(filename string) *Relationship {
	target := fmt.Sprintf("media/%s", filename)

	// Check if already exists
	if existing := r.GetByTarget(target); existing != nil {
		return existing
	}

	return r.AddDocumentRelationship(TypeImage, target, TargetModeInternal)
}

// AddHyperlink adds a hyperlink relationship
func (r *Relationships) AddHyperlink(url string) *Relationship {
	// Check if already exists
	if existing := r.GetExternal(url); existing != nil {
		return existing
	}

	return r.AddDocumentRelationship(TypeHyperlink, url, TargetModeExternal)
}

// AddHeader adds a header relationship
func (r *Relationships) AddHeader(headerFile string) *Relationship {
	return r.AddDocumentRelationship(TypeHeader, headerFile, TargetModeInternal)
}

// AddFooter adds a footer relationship
func (r *Relationships) AddFooter(footerFile string) *Relationship {
	return r.AddDocumentRelationship(TypeFooter, footerFile, TargetModeInternal)
}

// AddChart adds a chart relationship
func (r *Relationships) AddChart(chartFile string) *Relationship {
	return r.AddDocumentRelationship(TypeChart, chartFile, TargetModeInternal)
}

// AddCustomXML adds a custom XML relationship
func (r *Relationships) AddCustomXML(xmlFile string) *Relationship {
	return r.AddDocumentRelationship(TypeCustomXML, xmlFile, TargetModeInternal)
}

// GetByID returns a relationship by ID
func (r *Relationships) GetByID(id string) *Relationship {
	return r.items[id]
}

// GetByType returns all relationships of a specific type
func (r *Relationships) GetByType(relType string) []*Relationship {
	return r.byType[relType]
}

// GetByTarget returns a relationship by target
func (r *Relationships) GetByTarget(target string) *Relationship {
	// First try direct lookup
	if rel := r.byTarget[target]; rel != nil {
		return rel
	}

	// Try with different path combinations
	for _, rel := range r.items {
		if rel.Target == target {
			return rel
		}
	}

	return nil
}

// GetExternal returns an external relationship by URL
func (r *Relationships) GetExternal(url string) *Relationship {
	return r.external[url]
}

// GetImages returns all image relationships
func (r *Relationships) GetImages() []*Relationship {
	return r.GetByType(TypeImage)
}

// GetHyperlinks returns all hyperlink relationships
func (r *Relationships) GetHyperlinks() []*Relationship {
	return r.GetByType(TypeHyperlink)
}

// GetHeaders returns all header relationships
func (r *Relationships) GetHeaders() []*Relationship {
	return r.GetByType(TypeHeader)
}

// GetFooters returns all footer relationships
func (r *Relationships) GetFooters() []*Relationship {
	return r.GetByType(TypeFooter)
}

// Remove removes a relationship by ID
func (r *Relationships) Remove(id string) bool {
	rel, exists := r.items[id]
	if !exists {
		return false
	}

	// Remove from all indexes
	delete(r.items, id)
	delete(r.byTarget, rel.TargetKey)

	if rel.TargetMode == TargetModeExternal {
		delete(r.external, rel.Target)
	}

	// Remove from type index
	if rels := r.byType[rel.Type]; rels != nil {
		newRels := make([]*Relationship, 0, len(rels)-1)
		for _, r := range rels {
			if r.ID != id {
				newRels = append(newRels, r)
			}
		}
		r.byType[rel.Type] = newRels
	}

	// Remove from package/document rels
	r.packageRels = removeFromSlice(r.packageRels, id)
	r.documentRels = removeFromSlice(r.documentRels, id)

	return true
}

// removeFromSlice removes a relationship from a slice
func removeFromSlice(slice []*Relationship, id string) []*Relationship {
	newSlice := make([]*Relationship, 0, len(slice))
	for _, rel := range slice {
		if rel.ID != id {
			newSlice = append(newSlice, rel)
		}
	}
	return newSlice
}

// Clear removes all relationships
func (r *Relationships) Clear() {
	r.items = make(map[string]*Relationship)
	r.byType = make(map[string][]*Relationship)
	r.byTarget = make(map[string]*Relationship)
	r.external = make(map[string]*Relationship)
	r.packageRels = make([]*Relationship, 0)
	r.documentRels = make([]*Relationship, 0)
	r.counter = 0
	r.nextID = 1
}

// Count returns the total number of relationships
func (r *Relationships) Count() int {
	return len(r.items)
}

// CountByType returns the number of relationships of a specific type
func (r *Relationships) CountByType(relType string) int {
	return len(r.byType[relType])
}

// GetPackageRelationships returns all package-level relationships
func (r *Relationships) GetPackageRelationships() []*Relationship {
	return r.packageRels
}

// GetDocumentRelationships returns all document-level relationships
func (r *Relationships) GetDocumentRelationships() []*Relationship {
	return r.documentRels
}

// PackageXML generates the package relationships XML
func (r *Relationships) PackageXML() ([]byte, error) {
	relsXML := &RelationshipsXML{
		Xmlns:         "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationships: make([]RelationshipXML, 0, len(r.packageRels)),
	}

	for _, rel := range r.packageRels {
		relXML := RelationshipXML{
			ID:     rel.ID,
			Type:   rel.Type,
			Target: rel.Target,
		}
		if rel.TargetMode == TargetModeExternal {
			relXML.TargetMode = rel.TargetMode
		}
		relsXML.Relationships = append(relsXML.Relationships, relXML)
	}

	return xml.MarshalIndent(relsXML, "", "  ")
}

// DocumentXML generates the document relationships XML
func (r *Relationships) DocumentXML() ([]byte, error) {
	relsXML := &RelationshipsXML{
		Xmlns:         "http://schemas.openxmlformats.org/package/2006/relationships",
		Relationships: make([]RelationshipXML, 0, len(r.documentRels)),
	}

	for _, rel := range r.documentRels {
		relXML := RelationshipXML{
			ID:     rel.ID,
			Type:   rel.Type,
			Target: rel.Target,
		}
		if rel.TargetMode == TargetModeExternal {
			relXML.TargetMode = rel.TargetMode
		}
		relsXML.Relationships = append(relsXML.Relationships, relXML)
	}

	return xml.MarshalIndent(relsXML, "", "  ")
}

// Clone creates a deep copy of the relationships
func (r *Relationships) Clone() *Relationships {
	clone := New()

	// Clone all relationships
	for _, rel := range r.items {
		clonedRel := &Relationship{
			ID:         rel.ID,
			Type:       rel.Type,
			Target:     rel.Target,
			TargetMode: rel.TargetMode,
			TargetKey:  rel.TargetKey,
		}

		clone.items[clonedRel.ID] = clonedRel
		clone.byType[clonedRel.Type] = append(clone.byType[clonedRel.Type], clonedRel)
		clone.byTarget[clonedRel.TargetKey] = clonedRel

		if clonedRel.TargetMode == TargetModeExternal {
			clone.external[clonedRel.Target] = clonedRel
		}
	}

	// Clone package relationships
	for _, rel := range r.packageRels {
		if clonedRel := clone.items[rel.ID]; clonedRel != nil {
			clone.packageRels = append(clone.packageRels, clonedRel)
		}
	}

	// Clone document relationships
	for _, rel := range r.documentRels {
		if clonedRel := clone.items[rel.ID]; clonedRel != nil {
			clone.documentRels = append(clone.documentRels, clonedRel)
		}
	}

	clone.counter = r.counter
	clone.nextID = r.nextID

	return clone
}

// Merge merges another Relationships into this one
func (r *Relationships) Merge(other *Relationships) {
	if other == nil {
		return
	}

	// Map old IDs to new IDs
	idMap := make(map[string]string)

	// Merge all relationships
	for oldID, rel := range other.items {
		// Check if relationship already exists
		existing := r.findDuplicate(rel)
		if existing != nil {
			idMap[oldID] = existing.ID
			continue
		}

		// Create new relationship
		newRel := &Relationship{
			ID:         r.generateID(),
			Type:       rel.Type,
			Target:     rel.Target,
			TargetMode: rel.TargetMode,
			TargetKey:  rel.TargetKey,
		}

		idMap[oldID] = newRel.ID

		r.items[newRel.ID] = newRel
		r.byType[newRel.Type] = append(r.byType[newRel.Type], newRel)
		r.byTarget[newRel.TargetKey] = newRel

		if newRel.TargetMode == TargetModeExternal {
			r.external[newRel.Target] = newRel
		}

		// Add to appropriate list
		if isPackageRel(other.packageRels, rel) {
			r.packageRels = append(r.packageRels, newRel)
		} else if isDocumentRel(other.documentRels, rel) {
			r.documentRels = append(r.documentRels, newRel)
		}
	}
}

// findDuplicate finds a duplicate relationship
func (r *Relationships) findDuplicate(rel *Relationship) *Relationship {
	// Check by target key
	if existing := r.byTarget[rel.TargetKey]; existing != nil {
		return existing
	}

	// Check external URLs
	if rel.TargetMode == TargetModeExternal {
		return r.external[rel.Target]
	}

	return nil
}

// isPackageRel checks if a relationship is in the package relationships
func isPackageRel(packageRels []*Relationship, rel *Relationship) bool {
	for _, pRel := range packageRels {
		if pRel == rel {
			return true
		}
	}
	return false
}

// isDocumentRel checks if a relationship is in the document relationships
func isDocumentRel(documentRels []*Relationship, rel *Relationship) bool {
	for _, dRel := range documentRels {
		if dRel == rel {
			return true
		}
	}
	return false
}

// Validate validates all relationships
func (r *Relationships) Validate() error {
	// Check for duplicate IDs
	seen := make(map[string]bool)
	for id := range r.items {
		if seen[id] {
			return fmt.Errorf("duplicate relationship ID: %s", id)
		}
		seen[id] = true
	}

	// Validate each relationship
	for _, rel := range r.items {
		if err := rel.Validate(); err != nil {
			return fmt.Errorf("invalid relationship %s: %w", rel.ID, err)
		}
	}

	return nil
}

// Validate validates a single relationship
func (rel *Relationship) Validate() error {
	if rel.ID == "" {
		return fmt.Errorf("relationship ID is required")
	}

	if rel.Type == "" {
		return fmt.Errorf("relationship type is required")
	}

	if rel.Target == "" {
		return fmt.Errorf("relationship target is required")
	}

	if rel.TargetMode != "" && rel.TargetMode != TargetModeInternal && rel.TargetMode != TargetModeExternal {
		return fmt.Errorf("invalid target mode: %s", rel.TargetMode)
	}

	return nil
}

// String returns a string representation of the relationship
func (rel *Relationship) String() string {
	mode := rel.TargetMode
	if mode == "" {
		mode = TargetModeInternal
	}
	return fmt.Sprintf("Relationship{ID: %s, Type: %s, Target: %s, Mode: %s}",
		rel.ID, getTypeName(rel.Type), rel.Target, mode)
}

// getTypeName returns a human-readable name for a relationship type
func getTypeName(relType string) string {
	parts := strings.Split(relType, "/")
	if len(parts) > 0 {
		return parts[len(parts)-1]
	}
	return relType
}

// Helper function to create a unique image relationship
func (r *Relationships) AddUniqueImage(filename string) *Relationship {
	// Generate unique filename if needed
	target := fmt.Sprintf("media/%s", filename)

	if existing := r.GetByTarget(target); existing != nil {
		// Generate unique name
		ext := path.Ext(filename)
		base := strings.TrimSuffix(filename, ext)
		target = fmt.Sprintf("media/%s_%s%s", base, uuid.New().String()[:8], ext)
	}

	return r.AddDocumentRelationship(TypeImage, target, TargetModeInternal)
}

// GetOrCreateHyperlink gets an existing hyperlink relationship or creates a new one
func (r *Relationships) GetOrCreateHyperlink(url string) *Relationship {
	if existing := r.GetExternal(url); existing != nil {
		return existing
	}
	return r.AddHyperlink(url)
}
