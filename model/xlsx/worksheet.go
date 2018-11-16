package xlsx; import "encoding/xml"
// Worksheet was generated 2018-11-16 18:31:45 by rcir178 on rcir178.
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

