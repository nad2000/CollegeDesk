package model

import (
	"encoding/json"
	"testing"

	_ "github.com/go-sql-driver/mysql"
	_ "github.com/jinzhu/gorm/dialects/sqlite"
)

var (
	barChart = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Students</a:t></a:r><a:r><a:rPr lang="en-US" baseline="0"/><a:t> in sections</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:invertIfNegative val="0"/><c:cat><c:strRef><c:f>bar!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>bar!$B$2:$B$5</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>12</c:v></c:pt><c:pt idx="3"><c:v>14</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls><c:gapWidth val="150"/><c:axId val="99803520"/><c:axId val="99805056"/></c:barChart><c:catAx><c:axId val="99803520"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:title><c:tx><c:rich><a:bodyPr rot="0" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Section</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="99805056"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx><c:valAx><c:axId val="99805056"/><c:scaling><c:orientation val="minMax"/><c:max val="15"/><c:min val="1"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorGridlines/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Number</a:t></a:r><a:r><a:rPr lang="en-US" baseline="0"/><a:t> of Students</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="99803520"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:layout/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`
	columnChart = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US" sz="1400"/><a:t>Students in sections</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:invertIfNegative val="0"/><c:cat><c:strRef><c:f>column!$A$2:$A$5</c:f><c:strCache><c:ptCount val="4"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt><c:pt idx="3"><c:v>D</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>column!$B$2:$B$5</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="4"/><c:pt idx="0"><c:v>10</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>12</c:v></c:pt><c:pt idx="3"><c:v>14</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls><c:gapWidth val="150"/><c:axId val="98285440"/><c:axId val="98286976"/></c:barChart><c:catAx><c:axId val="98285440"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Section</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="98286976"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx><c:valAx><c:axId val="98286976"/><c:scaling><c:orientation val="minMax"/><c:max val="14"/><c:min val="1"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:title><c:tx><c:rich><a:bodyPr rot="0" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Number of Students</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="98285440"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:layout/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`
	lineChart = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Petrol Price in Bangalore</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:marker><c:symbol val="none"/></c:marker><c:cat><c:strRef><c:f>line!$A$2:$A$7</c:f><c:strCache><c:ptCount val="6"/><c:pt idx="0"><c:v>Jan</c:v></c:pt><c:pt idx="1"><c:v>Feb</c:v></c:pt><c:pt idx="2"><c:v>Mar</c:v></c:pt><c:pt idx="3"><c:v>Apr</c:v></c:pt><c:pt idx="4"><c:v>May</c:v></c:pt><c:pt idx="5"><c:v>Jun</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>line!$B$2:$B$7</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="6"/><c:pt idx="0"><c:v>40</c:v></c:pt><c:pt idx="1"><c:v>33</c:v></c:pt><c:pt idx="2"><c:v>55</c:v></c:pt><c:pt idx="3"><c:v>49</c:v></c:pt><c:pt idx="4"><c:v>89</c:v></c:pt><c:pt idx="5"><c:v>65</c:v></c:pt></c:numCache></c:numRef></c:val><c:smooth val="0"/></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls><c:marker val="1"/><c:smooth val="0"/><c:axId val="99820672"/><c:axId val="99822208"/></c:lineChart><c:catAx><c:axId val="99820672"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Month</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="99822208"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx><c:valAx><c:axId val="99822208"/><c:scaling><c:orientation val="minMax"/><c:max val="95"/><c:min val="20"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:title><c:tx><c:rich><a:bodyPr rot="0" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Petrol Price</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="99820672"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:layout/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`

	scatterChart = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200"/></a:pPr><a:r><a:rPr lang="en-US" sz="1200"/><a:t>TCS vs Infy Returns Scatter Plot</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="1"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Returns</c:v></c:tx><c:spPr><a:ln w="28575"><a:noFill/></a:ln></c:spPr><c:xVal><c:numRef><c:f>scatter!$A$2:$A$7</c:f><c:numCache><c:formatCode>0%</c:formatCode><c:ptCount val="6"/><c:pt idx="0"><c:v>0.02</c:v></c:pt><c:pt idx="1"><c:v>0.05</c:v></c:pt><c:pt idx="2"><c:v>0.03</c:v></c:pt><c:pt idx="3"><c:v>0.09</c:v></c:pt><c:pt idx="4"><c:v>0.05</c:v></c:pt><c:pt idx="5"><c:v>-0.01</c:v></c:pt></c:numCache></c:numRef></c:xVal><c:yVal><c:numRef><c:f>scatter!$B$2:$B$7</c:f><c:numCache><c:formatCode>0%</c:formatCode><c:ptCount val="6"/><c:pt idx="0"><c:v>0.01</c:v></c:pt><c:pt idx="1"><c:v>-0.02</c:v></c:pt><c:pt idx="2"><c:v>-0.05</c:v></c:pt><c:pt idx="3"><c:v>0.1</c:v></c:pt><c:pt idx="4"><c:v>0.08</c:v></c:pt><c:pt idx="5"><c:v>0.03</c:v></c:pt></c:numCache></c:numRef></c:yVal><c:smooth val="0"/></c:ser><c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls><c:axId val="99862784"/><c:axId val="99872768"/></c:scatterChart><c:valAx><c:axId val="99862784"/><c:scaling><c:orientation val="minMax"/><c:max val="0.16000000000000003"/><c:min val="-4.0000000000000008E-2"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>TCS Returns</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title><c:numFmt formatCode="0%" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="99872768"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx><c:valAx><c:axId val="99872768"/><c:scaling><c:orientation val="minMax"/><c:max val="0.2"/><c:min val="-8.0000000000000016E-2"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:title><c:tx><c:rich><a:bodyPr rot="0" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Infy Returns</a:t></a:r></a:p></c:rich></c:tx><c:layout/><c:overlay val="0"/></c:title><c:numFmt formatCode="0%" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="99862784"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx></c:plotArea><c:legend><c:legendPos val="r"/><c:layout/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/><c:showDLblsOverMax val="0"/></c:chart><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`
)

func dumpStruct(t *testing.T, v interface{}) {
	doc, err := json.MarshalIndent(v, "", "  ")
	if err != nil {
		t.Error(err)
	} else {
		t.Log(string(doc))
	}
}

func TestChartUnmarshaling(t *testing.T) {
	chart := UnmarshalChart([]byte(barChart))
	// dumpStruct(t, chart)

	if expected := "Students in sections"; chart.Title.Value() != expected {
		t.Errorf("Wrong title: %q, expected: %q", chart.Title.Value(), expected)
	}

	// if expected := "Number of Students"; chart.XLabel() != expected {
	// 	t.Errorf("Wrong X-axix label: %q, expected: %q", chart.XLabel(), expected)
	// }

	// if expected := "Section"; chart.YLabel() != expected {
	// 	t.Errorf("Wrong Y-axix label: %q, expected: %q", chart.YLabel(), expected)
	// }

	if expected := "Section"; chart.XLabel() != expected {
		t.Errorf("Wrong X-axis label: %q, expected: %q", chart.XLabel(), expected)
	}

	if expected := "Number of Students"; chart.YLabel() != expected {
		t.Errorf("Wrong Y-axis label: %q, expected: %q", chart.YLabel(), expected)
	}

	if expected, got := "1", chart.XMinValue(); got != expected {
		t.Errorf("Wrong X-axis min value: %q, expected: %q", got, expected)
	}

	if expected, got := "15", chart.XMaxValue(); got != expected {
		t.Errorf("Wrong X-axis max value: %q, expected: %q", got, expected)
	}
	if expected, got := "", chart.YMinValue(); got != expected {
		t.Errorf("Wrong Y-axis min value: %q, expected: %q", got, expected)
	}

	if expected, got := "", chart.YMaxValue(); got != expected {
		t.Errorf("Wrong Y-axis max value: %q, expected: %q", got, expected)
	}

	if t.Failed() {
		dumpStruct(t, chart)
	}
}

func TestColumnChartUnmarshaling(t *testing.T) {
	chart := UnmarshalChart([]byte(columnChart))
	// dumpStruct(t, chart)

	if expected := "Students in sections"; chart.Title.Value() != expected {
		t.Errorf("Wrong title: %q, expected: %q", chart.Title.Value(), expected)
	}

	// if expected := "Number of Students"; chart.XLabel() != expected {
	// 	t.Errorf("Wrong X-axix label: %q, expected: %q", chart.XLabel(), expected)
	// }

	// if expected := "Section"; chart.YLabel() != expected {
	// 	t.Errorf("Wrong Y-axix label: %q, expected: %q", chart.YLabel(), expected)
	// }

	if expected := "Section"; chart.XLabel() != expected {
		t.Errorf("Wrong X-axis label: %q, expected: %q", chart.XLabel(), expected)
	}

	if expected := "Number of Students"; chart.YLabel() != expected {
		t.Errorf("Wrong Y-axis label: %q, expected: %q", chart.YLabel(), expected)
	}

	if expected, got := "1", chart.XMinValue(); got != expected {
		t.Errorf("Wrong X-axis min value: %q, expected: %q", got, expected)
	}

	if expected, got := "14", chart.XMaxValue(); got != expected {
		t.Errorf("Wrong X-axis max value: %q, expected: %q", got, expected)
	}
	if expected, got := "", chart.YMinValue(); got != expected {
		t.Errorf("Wrong Y-axis min value: %q, expected: %q", got, expected)
	}

	if expected, got := "", chart.YMaxValue(); got != expected {
		t.Errorf("Wrong Y-axis max value: %q, expected: %q", got, expected)
	}

	if t.Failed() {
		dumpStruct(t, chart)
	}
}

func TestLineChartUnmarshaling(t *testing.T) {
	chart := UnmarshalChart([]byte(lineChart))
	// dumpStruct(t, chart)

	if expected, got := "Petrol Price in Bangalore", chart.Title.Value(); got != expected {
		t.Errorf("Wrong title: %q, expected: %q", got, expected)
	}

	if expected, got := "Month", chart.XLabel(); got != expected {
		t.Errorf("Wrong X-axis label: %q, expected: %q", got, expected)
	}

	if expected, got := "Petrol Price", chart.YLabel(); got != expected {
		t.Errorf("Wrong Y-axis label: %q, expected: %q", got, expected)
	}

	if expected, got := "20", chart.XMinValue(); got != expected {
		t.Errorf("Wrong X-axis min value: %q, expected: %q", got, expected)
	}

	if expected, got := "95", chart.XMaxValue(); got != expected {
		t.Errorf("Wrong X-axis max value: %q, expected: %q", got, expected)
	}
	if expected, got := "", chart.YMinValue(); got != expected {
		t.Errorf("Wrong Y-axis min value: %q, expected: %q", got, expected)
	}

	if expected, got := "", chart.YMaxValue(); got != expected {
		t.Errorf("Wrong Y-axis max value: %q, expected: %q", got, expected)
	}

	if t.Failed() {
		dumpStruct(t, chart)
	}
}

func TestScatterChartUnmarshaling(t *testing.T) {
	chart := UnmarshalChart([]byte(scatterChart))

	if expected, got := "TCS vs Infy Returns Scatter Plot", chart.Title.Value(); got != expected {
		t.Errorf("Wrong title: %q, expected: %q", got, expected)
	}

	if expected, got := "TCS Returns", chart.XLabel(); got != expected {
		t.Errorf("Wrong X-axis label: %q, expected: %q", got, expected)
	}

	if expected, got := "Infy Returns", chart.YLabel(); got != expected {
		t.Errorf("Wrong Y-axis label: %q, expected: %q", got, expected)
	}

	if expected, got := "-4.0000000000000008E-2", chart.XMinValue(); got != expected {
		t.Errorf("Wrong X-axis min value: %q, expected: %q", got, expected)
	}

	if expected, got := "0.16000000000000003", chart.XMaxValue(); got != expected {
		t.Errorf("Wrong X-axis max value: %q, expected: %q", got, expected)
	}
	if expected, got := "-8.0000000000000016E-2", chart.YMinValue(); got != expected {
		t.Errorf("Wrong Y-axis min value: %q, expected: %q", got, expected)
	}

	if expected, got := "0.2", chart.YMaxValue(); got != expected {
		t.Errorf("Wrong Y-axis max value: %q, expected: %q", got, expected)
	}

	if t.Failed() {
		dumpStruct(t, chart)
	}
}
