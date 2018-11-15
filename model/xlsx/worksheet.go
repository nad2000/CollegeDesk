package xlsx; import "encoding/xml"
// Worksheet was generated 2018-11-15 20:05:32 by rcir178 on rcir178.
type Worksheet struct {
	XMLName   xml.Name `xml:"worksheet"`
	Text      string   `xml:",chardata"`
	Xmlns     string   `xml:"xmlns,attr"`
	R         string   `xml:"r,attr"`
	Mc        string   `xml:"mc,attr"`
	Ignorable string   `xml:"Ignorable,attr"`
	X14ac     string   `xml:"x14ac,attr"`
	SheetPr   struct {
		Text       string `xml:",chardata"`
		FilterMode string `xml:"filterMode,attr"`
	} `xml:"sheetPr"`
	Dimension struct {
		Text string `xml:",chardata"`
		Ref  string `xml:"ref,attr"`
	} `xml:"dimension"`
	SheetViews struct {
		Text      string `xml:",chardata"`
		SheetView struct {
			Text           string `xml:",chardata"`
			WorkbookViewId string `xml:"workbookViewId,attr"`
			Selection      struct {
				Text       string `xml:",chardata"`
				ActiveCell string `xml:"activeCell,attr"`
				Sqref      string `xml:"sqref,attr"`
			} `xml:"selection"`
		} `xml:"sheetView"`
	} `xml:"sheetViews"`
	SheetFormatPr struct {
		Text             string `xml:",chardata"`
		DefaultRowHeight string `xml:"defaultRowHeight,attr"`
		DyDescent        string `xml:"dyDescent,attr"`
	} `xml:"sheetFormatPr"`
	Cols struct {
		Text string `xml:",chardata"`
		Col  struct {
			Text        string `xml:",chardata"`
			Min         string `xml:"min,attr"`
			Max         string `xml:"max,attr"`
			Width       string `xml:"width,attr"`
			BestFit     string `xml:"bestFit,attr"`
			CustomWidth string `xml:"customWidth,attr"`
		} `xml:"col"`
	} `xml:"cols"`
	SheetData struct {
		Text string `xml:",chardata"`
		Row  []struct {
			Text      string `xml:",chardata"`
			R         string `xml:"r,attr"`
			Spans     string `xml:"spans,attr"`
			DyDescent string `xml:"dyDescent,attr"`
			Hidden    string `xml:"hidden,attr"`
			C         []struct {
				Text string `xml:",chardata"`
				R    string `xml:"r,attr"`
				S    string `xml:"s,attr"`
				T    string `xml:"t,attr"`
				V    struct {
					Text string `xml:",chardata"` // 0, 1, 2, 3, 4, 5, 6, 34, ...
				} `xml:"v"`
			} `xml:"c"`
		} `xml:"row"`
	} `xml:"sheetData"`
	AutoFilter []struct {
		Text         string `xml:",chardata"`
		Ref          string `xml:"ref,attr"`
		FilterColumn []struct {
			Text    string `xml:",chardata"`
			ColId   string `xml:"colId,attr"`
			Filters struct {
				Text          string `xml:",chardata"`
				DateGroupItem []struct {
					Text             string `xml:",chardata"`
					Year             string `xml:"year,attr"`
					DateTimeGrouping string `xml:"dateTimeGrouping,attr"`
					Month            string `xml:"month,attr"`
					Day              string `xml:"day,attr"`
					Hour             string `xml:"hour,attr"`
					Minute           string `xml:"minute,attr"`
					Second           string `xml:"second,attr"`
				} `xml:"dateGroupItem"`
				Filter []struct {
					Text string `xml:",chardata"`
					Val  string `xml:"val,attr"`
				} `xml:"filter"`
			} `xml:"filters"`
			CustomFilters struct {
				Text         string `xml:",chardata"`
				And          string `xml:"and,attr"`
				CustomFilter []struct {
					Text     string `xml:",chardata"`
					Operator string `xml:"operator,attr"`
					Val      string `xml:"val,attr"`
				} `xml:"customFilter"`
			} `xml:"customFilters"`
			Top10 struct {
				Text      string `xml:",chardata"`
				Val       string `xml:"val,attr"`
				FilterVal string `xml:"filterVal,attr"`
				Top       string `xml:"top,attr"`
			} `xml:"top10"`
			DynamicFilter struct {
				Text string `xml:",chardata"`
				Type string `xml:"type,attr"`
				Val  string `xml:"val,attr"`
			} `xml:"dynamicFilter"`
			ColorFilter struct {
				Text  string `xml:",chardata"`
				DxfId string `xml:"dxfId,attr"`
			} `xml:"colorFilter"`
		} `xml:"filterColumn"`
	} `xml:"autoFilter"`
	PageMargins struct {
		Text   string `xml:",chardata"`
		Left   string `xml:"left,attr"`
		Right  string `xml:"right,attr"`
		Top    string `xml:"top,attr"`
		Bottom string `xml:"bottom,attr"`
		Header string `xml:"header,attr"`
		Footer string `xml:"footer,attr"`
	} `xml:"pageMargins"`
} 

