package xlsx; import "encoding/xml"
// PivotTableDefinition was generated 2019-04-06 22:26:52 by rcir178 on rcir178-Latitude-E7470.
type PivotTableDefinition struct {
	XMLName                 xml.Name `xml:"pivotTableDefinition"`
	Text                    string   `xml:",chardata"`
	Xmlns                   string   `xml:"xmlns,attr"`
	Name                    string   `xml:"name,attr"`
	CacheId                 string   `xml:"cacheId,attr"`
	ApplyNumberFormats      string   `xml:"applyNumberFormats,attr"`
	ApplyBorderFormats      string   `xml:"applyBorderFormats,attr"`
	ApplyFontFormats        string   `xml:"applyFontFormats,attr"`
	ApplyPatternFormats     string   `xml:"applyPatternFormats,attr"`
	ApplyAlignmentFormats   string   `xml:"applyAlignmentFormats,attr"`
	ApplyWidthHeightFormats string   `xml:"applyWidthHeightFormats,attr"`
	DataCaption             string   `xml:"dataCaption,attr"`
	UpdatedVersion          string   `xml:"updatedVersion,attr"`
	MinRefreshableVersion   string   `xml:"minRefreshableVersion,attr"`
	UseAutoFormatting       string   `xml:"useAutoFormatting,attr"`
	ItemPrintTitles         string   `xml:"itemPrintTitles,attr"`
	CreatedVersion          string   `xml:"createdVersion,attr"`
	Indent                  string   `xml:"indent,attr"`
	Outline                 string   `xml:"outline,attr"`
	OutlineData             string   `xml:"outlineData,attr"`
	MultipleFieldFilters    string   `xml:"multipleFieldFilters,attr"`
	Location                struct {
		Text           string `xml:",chardata"`
		Ref            string `xml:"ref,attr"`
		FirstHeaderRow string `xml:"firstHeaderRow,attr"`
		FirstDataRow   string `xml:"firstDataRow,attr"`
		FirstDataCol   string `xml:"firstDataCol,attr"`
		RowPageCount   string `xml:"rowPageCount,attr"`
		ColPageCount   string `xml:"colPageCount,attr"`
	} `xml:"location"`
	PivotFields struct {
		Text       string `xml:",chardata"`
		Count      string `xml:"count,attr"`
		PivotField []struct {
			Text      string `xml:",chardata"`
			Axis      string `xml:"axis,attr"`
			NumFmtId  string `xml:"numFmtId,attr"`
			ShowAll   string `xml:"showAll,attr"`
			DataField string `xml:"dataField,attr"`
			Items     struct {
				Text  string `xml:",chardata"`
				Count string `xml:"count,attr"`
				Item  []struct {
					Text string `xml:",chardata"`
					X    string `xml:"x,attr"`
					T    string `xml:"t,attr"`
				} `xml:"item"`
			} `xml:"items"`
		} `xml:"pivotField"`
	} `xml:"pivotFields"`
	RowFields struct {
		Text  string `xml:",chardata"`
		Count string `xml:"count,attr"`
		Field []struct {
			Text string `xml:",chardata"`
			X    string `xml:"x,attr"`
		} `xml:"field"`
	} `xml:"rowFields"`
	RowItems struct {
		Text  string `xml:",chardata"`
		Count string `xml:"count,attr"`
		I     []struct {
			Text string `xml:",chardata"`
			R    string `xml:"r,attr"`
			T    string `xml:"t,attr"`
			X    struct {
				Text string `xml:",chardata"`
				V    string `xml:"v,attr"`
			} `xml:"x"`
		} `xml:"i"`
	} `xml:"rowItems"`
	ColFields struct {
		Text  string `xml:",chardata"`
		Count string `xml:"count,attr"`
		Field []struct {
			Text string `xml:",chardata"`
			X    string `xml:"x,attr"`
		} `xml:"field"`
	} `xml:"colFields"`
	ColItems struct {
		Text  string `xml:",chardata"`
		Count string `xml:"count,attr"`
		I     []struct {
			Text string `xml:",chardata"`
			R    string `xml:"r,attr"`
			I    string `xml:"i,attr"`
			T    string `xml:"t,attr"`
			X    []struct {
				Text string `xml:",chardata"`
				V    string `xml:"v,attr"`
			} `xml:"x"`
		} `xml:"i"`
	} `xml:"colItems"`
	PageFields struct {
		Text      string `xml:",chardata"`
		Count     string `xml:"count,attr"`
		PageField []struct {
			Text string `xml:",chardata"`
			Fld  string `xml:"fld,attr"`
			Hier string `xml:"hier,attr"`
		} `xml:"pageField"`
	} `xml:"pageFields"`
	DataFields struct {
		Text      string `xml:",chardata"`
		Count     string `xml:"count,attr"`
		DataField []struct {
			Text      string `xml:",chardata"`
			Name      string `xml:"name,attr"`
			Fld       string `xml:"fld,attr"`
			BaseField string `xml:"baseField,attr"`
			BaseItem  string `xml:"baseItem,attr"`
			Subtotal  string `xml:"subtotal,attr"`
		} `xml:"dataField"`
	} `xml:"dataFields"`
	PivotTableStyleInfo struct {
		Text           string `xml:",chardata"`
		Name           string `xml:"name,attr"`
		ShowRowHeaders string `xml:"showRowHeaders,attr"`
		ShowColHeaders string `xml:"showColHeaders,attr"`
		ShowRowStripes string `xml:"showRowStripes,attr"`
		ShowColStripes string `xml:"showColStripes,attr"`
		ShowLastColumn string `xml:"showLastColumn,attr"`
	} `xml:"pivotTableStyleInfo"`
	ExtLst struct {
		Text string `xml:",chardata"`
		Ext  struct {
			Text                 string `xml:",chardata"`
			URI                  string `xml:"uri,attr"`
			X14                  string `xml:"x14,attr"`
			PivotTableDefinition struct {
				Text          string `xml:",chardata"`
				HideValuesRow string `xml:"hideValuesRow,attr"`
				Xm            string `xml:"xm,attr"`
			} `xml:"pivotTableDefinition"`
		} `xml:"ext"`
	} `xml:"extLst"`
} 

