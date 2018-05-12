package models

import "encoding/xml"

// XmlxWorkbookRels contains xmlxWorkbookRelations which maps sheet id and sheet XML.
type xlsxRelationships struct {
	XMLName       xml.Name           `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []xlsxRelationship `xml:"Relationship"`
}

// XmlxWorkbookRelation maps sheet id and xl/worksheets/_rels/sheet%d.xml.rels
type xlsxRelationship struct {
	ID         string `xml:"Id,attr"`
	Target     string `xml:",attr"`
	Type       string `xml:",attr"`
	TargetMode string `xml:",attr,omitempty"`
}

// workbookRelsReader provides function to read and unmarshal workbook
// relationships of XLSX file.
func marshalRelationships(fileContent string) (content xlsxRelationships) {
	xml.Unmarshal([]byte(fileContent), &content)
	return
}
