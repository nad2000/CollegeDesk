package model

import (
	"encoding/xml"
	"strings"
)

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

type xlsxSapeProperties struct {
	Properies string `xml:",innerxml"`
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

// xlsxBareChart - minial requiered implementation of the Chart object
type xlsxBareChart struct {
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
func (c *xlsxBareChart) ItemCount() int {
	chart := c.PlotArea.Chart
	if chart.CategoryCount.Value > chart.XPointCount.Value {
		return chart.CategoryCount.Value
	}
	return chart.XPointCount.Value
}

// Type - chart type - short-cut
func (c *xlsxBareChart) Type() string {
	fromElementName := strings.Title(strings.TrimSuffix(c.PlotArea.Chart.XMLName.Local, "Chart"))
	if fromElementName == "Bar" && c.PlotArea.Chart.BarDir.Value == "col" {
		return "Column"
	}
	return fromElementName
}

// XLabel - X-axis title
func (c *xlsxBareChart) XLabel() string {
	if c.PlotArea.ValAxes != nil && len(c.PlotArea.ValAxes) > 1 {
		return c.PlotArea.ValAxes[0].Title.Value()
	}
	if c.PlotArea.CatAxTitles != nil {
		return c.PlotArea.CatAxTitles[0].Value()
	}
	return ""
}

// YLabel - Y-axis title
func (c *xlsxBareChart) YLabel() string {
	if c.PlotArea.ValAxes != nil {
		if len(c.PlotArea.ValAxes) > 1 {
			return c.PlotArea.ValAxes[1].Title.Value()
		}
		return c.PlotArea.ValAxes[0].Title.Value()
	}
	return ""
}

// XMinValue - X-axis min value
func (c *xlsxBareChart) XMinValue() string {
	if c.PlotArea.ValAxes != nil {
		return c.PlotArea.ValAxes[0].Min.Value
	}
	return ""
}

// XMaxValue - X-axis max value
func (c *xlsxBareChart) XMaxValue() string {
	if c.PlotArea.ValAxes != nil {
		return c.PlotArea.ValAxes[0].Max.Value
	}
	return ""
}

// YMinValue - Y-axis min value
func (c *xlsxBareChart) YMinValue() string {
	if c.PlotArea.ValAxes != nil && len(c.PlotArea.ValAxes) > 1 {
		return c.PlotArea.ValAxes[1].Min.Value
	}
	return ""
}

// YMaxValue - Y-axis max value
func (c *xlsxBareChart) YMaxValue() string {
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
func unmarshalRelationships(fileContent string) (content xlsxRelationships) {
	xml.Unmarshal([]byte(fileContent), &content)
	return
}

func UnmarshalChart(fileContent string) (content xlsxBareChart) {
	xml.Unmarshal([]byte(fileContent), &content)
	return
}

func unmarshalDrawing(fileContent string) (content xlsxBareDrawing) {
	// log.Info("----", fileContent)
	xml.Unmarshal([]byte(fileContent), &content)
	return
}
