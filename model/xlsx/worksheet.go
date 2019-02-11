package xlsx; import "encoding/xml"
// Worksheet was generated 2019-02-12 00:21:04 by rcir178 on rcir178-Latitude-E7470.
type Worksheet struct {
	XMLName   xml.Name `xml:"worksheet"`
	Text      string   `xml:",chardata"`
	Xmlns     string   `xml:"xmlns,attr"`
	R         string   `xml:"r,attr"`
	Mc        string   `xml:"mc,attr"`
	X14ac     string   `xml:"x14ac,attr"`
	Ignorable string   `xml:"Ignorable,attr"`
	SortState []struct {
		Text          string `xml:",chardata"`
		Ref           string `xml:"ref,attr"`
		ColumnSort    string `xml:"columnSort,attr"`
		SortCondition []struct {
			Text       string `xml:",chardata"`
			SortBy     string `xml:"sortBy,attr"`
			Ref        string `xml:"ref,attr"`
			DxfId      string `xml:"dxfId,attr"`
			Descending string `xml:"descending,attr"`
			CustomList string `xml:"customList,attr"`
			IconSet    string `xml:"iconSet,attr"`
			IconId     string `xml:"iconId,attr"`
		} `xml:"sortCondition"`
	} `xml:"sortState"`
	AutoFilter []struct {
		Text         string `xml:",chardata"`
		Ref          string `xml:"ref,attr"`
		FilterColumn []struct {
			Text          string `xml:",chardata"`
			ColId         string `xml:"colId,attr"`
			CustomFilters struct {
				Text         string `xml:",chardata"`
				And          string `xml:"and,attr"`
				CustomFilter []struct {
					Text     string `xml:",chardata"`
					Val      string `xml:"val,attr"`
					Operator string `xml:"operator,attr"`
				} `xml:"customFilter"`
			} `xml:"customFilters"`
			Filters struct {
				Text   string `xml:",chardata"`
				Filter []struct {
					Text string `xml:",chardata"`
					Val  string `xml:"val,attr"`
				} `xml:"filter"`
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
			} `xml:"filters"`
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
	ConditionalFormatting []struct {
		Text   string `xml:",chardata"`
		Sqref  string `xml:"sqref,attr"`
		CfRule []struct {
			Text         string `xml:",chardata"`
			Type         string `xml:"type,attr"`
			DxfId        string `xml:"dxfId,attr"`
			Priority     string `xml:"priority,attr"`
			Operator     string `xml:"operator,attr"`
			AttrText     string `xml:"text,attr"`
			TimePeriod   string `xml:"timePeriod,attr"`
			AboveAverage string `xml:"aboveAverage,attr"`
			Percent      string `xml:"percent,attr"`
			Bottom       string `xml:"bottom,attr"`
			Rank         string `xml:"rank,attr"`
			Formula      []struct {
				Text string `xml:",chardata"` // 10, 20, 10, 20, 84, 99, 3...
			} `xml:"formula"`
			DataBar struct {
				Text string `xml:",chardata"`
				Cfvo []struct {
					Text string `xml:",chardata"`
					Type string `xml:"type,attr"`
				} `xml:"cfvo"`
				Color struct {
					Text string `xml:",chardata"`
					Rgb  string `xml:"rgb,attr"`
				} `xml:"color"`
			} `xml:"dataBar"`
			ExtLst struct {
				Text string `xml:",chardata"`
				Ext  struct {
					Text string `xml:",chardata"`
					X14  string `xml:"x14,attr"`
					URI  string `xml:"uri,attr"`
					ID   struct {
						Text string `xml:",chardata"` // {4D3A9FBE-0C2F-4905-AB07-...
					} `xml:"id"`
				} `xml:"ext"`
			} `xml:"extLst"`
			ColorScale struct {
				Text string `xml:",chardata"`
				Cfvo []struct {
					Text string `xml:",chardata"`
					Type string `xml:"type,attr"`
					Val  string `xml:"val,attr"`
				} `xml:"cfvo"`
				Color []struct {
					Text string `xml:",chardata"`
					Rgb  string `xml:"rgb,attr"`
				} `xml:"color"`
			} `xml:"colorScale"`
			IconSet struct {
				Text    string `xml:",chardata"`
				IconSet string `xml:"iconSet,attr"`
				Cfvo    []struct {
					Text string `xml:",chardata"`
					Type string `xml:"type,attr"`
					Val  string `xml:"val,attr"`
				} `xml:"cfvo"`
			} `xml:"iconSet"`
		} `xml:"cfRule"`
	} `xml:"conditionalFormatting"`
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

