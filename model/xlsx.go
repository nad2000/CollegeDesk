// Package model it the core of the system for accessing DB.
//go:generate sh -c "echo 'package xlsx; import \"encoding/xml\"' >xlsx/worksheet.go; zek -e <../assets/sheet.xml >>xlsx/worksheet.go"
//go:generate sh -c "echo 'package xlsx; import \"encoding/xml\"' >xlsx/shared_strings.go; zek -e <../assets/sharedStrings.xml >>xlsx/shared_strings.go"
//go:generate sh -c "echo 'package xlsx; import \"encoding/xml\"' >xlsx/pivot_table.go; zek -e <../assets/pivotTable.xml >>xlsx/pivot_table.go"
//go:generate sh -c "echo 'package xlsx; import \"encoding/xml\"' >xlsx/pivot_cache_definition.go; zek -e <../assets/pivotCacheDefinition.xml >>xlsx/pivot_cache_definition.go"
//go:generate sh -c "echo 'package xlsx; import \"encoding/xml\"' >xlsx/style_sheet.go; zek -e <../assets/styles.xml >>xlsx/style_sheet.go"
//go:generate sh -c "echo 'package xlsx; import \"encoding/xml\"' >xlsx/workbook.go; zek -e <../assets/workbook.xml >>xlsx/workbook.go"
package model

import (
	"encoding/xml"
	"extract-blocks/model/xlsx"
	"strings"

	log "github.com/Sirupsen/logrus"
	"github.com/nad2000/excelize"
)

// saveFileList provides a function to update given file content in file list
// of XLSX.
func saveFileList(file *excelize.File, name string, content []byte) {
	newContent := make([]byte, 0, len(excelize.XMLHeader)+len(content))
	newContent = append(newContent, []byte(excelize.XMLHeader)...)
	newContent = append(newContent, content...)
	file.XLSX[name] = newContent
}

func deleteAllRelationshipsToName(file *excelize.File, name string) {
	for _, n := range file.GetSheetMap() {
		rels := "xl/worksheets/_rels/" + strings.TrimPrefix(n, "xl/worksheets/") + ".rels"
		var sheetRels xlsxWorkbookRels
		content, ok := file.XLSX[rels]
		if !ok {
			continue
		}
		_ = xml.Unmarshal(content, &sheetRels)
		for k, v := range sheetRels.Relationships {
			if strings.Contains(v.Target, name) {
				sheetRels.Relationships = append(sheetRels.Relationships[:k], sheetRels.Relationships[k+1:]...)
			}
		}
		output, _ := xml.Marshal(sheetRels)
		saveFileList(file, rels, output)
	}
}

// DeleteAllComments deletes all the comments in the workbook
func DeleteAllComments(file *excelize.File) (wasRemoved bool) {
	for name := range file.XLSX {
		if strings.HasPrefix(name, "xl/comment") {
			delete(file.XLSX, name)
			id := strings.TrimPrefix(strings.TrimSuffix(name, ".xml"), "xl/comment")
			vmlName := "xl/drawings/vmlDrawing" + id + ".vml"
			deleteAllRelationshipsToName(file, vmlName)
			wasRemoved = true
		}
	}
	return
}

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

type xlsxAnyWithIntValAttribute struct {
	// XMLName xml.Name
	// XML     string `xml:",innerxml"`
	Value int `xml:"val,attr"`
}

type xlsxAnyWithStringValAttribute struct {
	// XMLName xml.Name
	// XML     string `xml:",innerxml"`
	Value string `xml:"val,attr"`
}

type xlsxAnyChart struct {
	XMLName       xml.Name
	BarDir        xlsxAnyWithStringValAttribute `xml:"barDir"`
	Data          string                        `xml:"ser>val>numRef>f"`
	XData         string                        `xml:"ser>xVal>numRef>f"`
	YData         string                        `xml:"ser>yVal>numRef>f"`
	CategoryCount xlsxAnyWithIntValAttribute    `xml:"ser>cat>strRef>strCache>ptCount"`
	XPointCount   xlsxAnyWithIntValAttribute    `xml:"ser>xVal>numRef>numCache>ptCount"`
	YPointCount   xlsxAnyWithIntValAttribute    `xml:"ser>yVal>numRef>numCache>ptCount"`
	// XML     string `xml:",innerxml"`
}

type anyHolder struct {
	// XMLName Name
	// XML string `xml:",innerxml"`
}

type xlsxTitle struct {
	Texts []string `xml:"tx>rich>p>r>t"`
}

func (t *xlsxTitle) Value() string {
	if t.Texts != nil {
		return strings.Join(t.Texts, "")
	}
	return ""
}

type xlsxValAx struct {
	Title xlsxTitle                     `xml:"title"`
	Min   xlsxAnyWithStringValAttribute `xml:"scaling>min"`
	Max   xlsxAnyWithStringValAttribute `xml:"scaling>max"`
}

type xlsxPlotArea struct {
	// XMLName xml.Name      `xml:"http://schemas.openxmlformats.org/drawingml/2006/chart plotArea"`
	ShapeProperties anyHolder    `xml:"spPr"`
	Layout          anyHolder    `xml:"layout"`
	ValAxes         []xlsxValAx  `xml:"valAx"`
	CatAxTitles     []xlsxTitle  `xml:"catAx>title"`
	DateAxis        anyHolder    `xml:"dateAx"`
	SeriesAxis      anyHolder    `xml:"serAx"`
	DataTable       anyHolder    `xml:"dTable"`
	ExtensionList   anyHolder    `xml:"extLst"`
	Chart           xlsxAnyChart `xml:",any"`
}

// XlsxBareChart - minial requiered implementation of the Chart object
type XlsxBareChart struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/chart chartSpace"`
	// XMLNSc         string          `xml:"xmlns:c,attr"`
	// XMLNSa         string          `xml:"xmlns:a,attr"`
	// XMLNSr         string          `xml:"xmlns:r,attr"`
	// XMLNSc16r2     string          `xml:"xmlns:c16r2,attr"`
	// Date1904       *attrValBool    `xml:"c:date1904"`
	// Lang           *attrValString  `xml:"lang"`
	// RoundedCorners *attrValBool    `xml:"c:roundedCorners"`
	// Chart          cChart          `xml:"c:chart"`
	// SpPr           *cSpPr          `xml:"c:spPr"`
	// TxPr           *cTxPr          `xml:"c:txPr"`
	// PrintSettings *cPrintSettings `xml:"c:printSettings"`
	// Title    string       `xml:"chart>title>tx>rich>p>r>t"`
	Title    xlsxTitle    `xml:"chart>title"`
	PlotArea xlsxPlotArea `xml:"chart>plotArea"`
}

// ItemCount - shortcut for number of itmems displayed
func (c *XlsxBareChart) ItemCount() int {
	chart := c.PlotArea.Chart
	if chart.CategoryCount.Value > chart.XPointCount.Value {
		return chart.CategoryCount.Value
	}
	return chart.XPointCount.Value
}

// Type - chart type - short-cut
func (c *XlsxBareChart) Type() string {
	fromElementName := strings.Title(strings.TrimSuffix(c.PlotArea.Chart.XMLName.Local, "Chart"))
	if fromElementName == "Bar" && c.PlotArea.Chart.BarDir.Value == "col" {
		return "Column"
	}
	return fromElementName
}

// XLabel - X-axis title
func (c *XlsxBareChart) XLabel() string {
	if c.PlotArea.ValAxes != nil && len(c.PlotArea.ValAxes) > 1 {
		return c.PlotArea.ValAxes[0].Title.Value()
	}
	if c.PlotArea.CatAxTitles != nil {
		return c.PlotArea.CatAxTitles[0].Value()
	}
	return ""
}

// YLabel - Y-axis title
func (c *XlsxBareChart) YLabel() string {
	if c.PlotArea.ValAxes != nil {
		if len(c.PlotArea.ValAxes) > 1 {
			return c.PlotArea.ValAxes[1].Title.Value()
		}
		return c.PlotArea.ValAxes[0].Title.Value()
	}
	return ""
}

// XMinValue - X-axis min value
func (c *XlsxBareChart) XMinValue() string {
	if c.PlotArea.ValAxes != nil {
		return c.PlotArea.ValAxes[0].Min.Value
	}
	return ""
}

// XMaxValue - X-axis max value
func (c *XlsxBareChart) XMaxValue() string {
	if c.PlotArea.ValAxes != nil {
		return c.PlotArea.ValAxes[0].Max.Value
	}
	return ""
}

// YMinValue - Y-axis min value
func (c *XlsxBareChart) YMinValue() string {
	if c.PlotArea.ValAxes != nil && len(c.PlotArea.ValAxes) > 1 {
		return c.PlotArea.ValAxes[1].Min.Value
	}
	return ""
}

// YMaxValue - Y-axis max value
func (c *XlsxBareChart) YMaxValue() string {
	if c.PlotArea.ValAxes != nil && len(c.PlotArea.ValAxes) > 1 {
		return c.PlotArea.ValAxes[1].Max.Value
	}
	return ""
}

type xlsxBareDrawing struct {
	XMLName xml.Name `xml:"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing wsDr"`
	// OneCellAnchor []*xdrCellAnchor `xml:"oneCellAnchor"`
	// TwoCellAnchor []*xdrCellAnchor `xml:"twoCellAnchor"`
	// A             string           `xml:"a,attr,omitempty"`
	// Xdr           string           `xml:"xdr,attr,omitempty"`
	// R             string           `xml:"r,attr,omitempty"`
	FromCol int `xml:"twoCellAnchor>from>col"`
	FromRow int `xml:"twoCellAnchor>from>row"`
	ToCol   int `xml:"twoCellAnchor>to>col"`
	ToRow   int `xml:"twoCellAnchor>to>row"`
}

// marshalRelationships provides function to read and unmarshal workbook
// relationships of XLSX file.
func unmarshalRelationships(fileContent []byte) (content xlsxRelationships) {
	xml.Unmarshal(fileContent, &content)
	return
}

// UnmarshalChart unmarshals the chart data
func UnmarshalChart(fileContent []byte) (content XlsxBareChart) {
	xml.Unmarshal(fileContent, &content)
	return
}

// UnmarshalWorksheet unmarshals a worksheets autofilter
func UnmarshalWorksheet(fileContent []byte) (content xlsx.Worksheet) {
	err := xml.Unmarshal(fileContent, &content)
	if err != nil {
		log.Errorf("ERROR: %#v", err)
		log.Info(string(fileContent))
	}
	return
}

// UnmarshalPivotCacheDefinition unmarshals a worksheets autofilter
func UnmarshalPivotCacheDefinition(fileContent []byte) (content xlsx.PivotCacheDefinition) {
	err := xml.Unmarshal(fileContent, &content)
	if err != nil {
		log.Errorf("ERROR: %#v", err)
		log.Info(string(fileContent))
	}
	return
}

// UnmarshalPivotTableDefinition  unmarshals a worksheets autofilter
func UnmarshalPivotTableDefinition(fileContent []byte) (content xlsx.PivotTableDefinition) {
	err := xml.Unmarshal(fileContent, &content)
	if err != nil {
		log.Errorf("ERROR: %#v", err)
		log.Info(string(fileContent))
	}
	return
}

func unmarshalDrawing(fileContent []byte) (content xlsxBareDrawing) {
	// log.Info("----", fileContent)
	xml.Unmarshal(fileContent, &content)
	return
}

// xmlxWorkbookRels contains xmlxWorkbookRelations which maps sheet id and sheet XML.
type xlsxWorkbookRels struct {
	XMLName       xml.Name               `xml:"http://schemas.openxmlformats.org/package/2006/relationships Relationships"`
	Relationships []xlsxWorkbookRelation `xml:"Relationship"`
}

// xmlxWorkbookRelation maps sheet id and xl/worksheets/_rels/sheet%d.xml.rels
type xlsxWorkbookRelation struct {
	ID         string `xml:"Id,attr"`
	Target     string `xml:",attr"`
	Type       string `xml:",attr"`
	TargetMode string `xml:",attr,omitempty"`
}
