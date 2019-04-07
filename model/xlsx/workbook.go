package xlsx; import "encoding/xml"
// Workbook was generated 2019-04-06 22:26:52 by rcir178 on rcir178-Latitude-E7470.
type Workbook struct {
	XMLName     xml.Name `xml:"workbook"`
	Text        string   `xml:",chardata"`
	Xmlns       string   `xml:"xmlns,attr"`
	R           string   `xml:"r,attr"`
	FileVersion struct {
		Text         string `xml:",chardata"`
		AppName      string `xml:"appName,attr"`
		LastEdited   string `xml:"lastEdited,attr"`
		LowestEdited string `xml:"lowestEdited,attr"`
		RupBuild     string `xml:"rupBuild,attr"`
	} `xml:"fileVersion"`
	WorkbookPr struct {
		Text                string `xml:",chardata"`
		DefaultThemeVersion string `xml:"defaultThemeVersion,attr"`
	} `xml:"workbookPr"`
	BookViews struct {
		Text         string `xml:",chardata"`
		WorkbookView struct {
			Text         string `xml:",chardata"`
			XWindow      string `xml:"xWindow,attr"`
			YWindow      string `xml:"yWindow,attr"`
			WindowWidth  string `xml:"windowWidth,attr"`
			WindowHeight string `xml:"windowHeight,attr"`
		} `xml:"workbookView"`
	} `xml:"bookViews"`
	Sheets struct {
		Text  string `xml:",chardata"`
		Sheet struct {
			Text    string `xml:",chardata"`
			Name    string `xml:"name,attr"`
			SheetId string `xml:"sheetId,attr"`
			ID      string `xml:"id,attr"`
		} `xml:"sheet"`
	} `xml:"sheets"`
	DefinedNames struct {
		Text        string `xml:",chardata"`
		DefinedName []struct {
			Text         string `xml:",chardata"` // Sheet5!$C$6:$C$7, 0.0001,...
			Name         string `xml:"name,attr"`
			LocalSheetId string `xml:"localSheetId,attr"`
			Hidden       string `xml:"hidden,attr"`
		} `xml:"definedName"`
	} `xml:"definedNames"`
	CalcPr struct {
		Text   string `xml:",chardata"`
		CalcId string `xml:"calcId,attr"`
	} `xml:"calcPr"`
} 

