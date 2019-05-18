package xlsx; import "encoding/xml"
// StyleSheet was generated 2019-04-06 22:26:52 by rcir178 on rcir178-Latitude-E7470.
type StyleSheet struct {
	XMLName xml.Name `xml:"styleSheet"`
	Text    string   `xml:",chardata"`
	Xmlns   string   `xml:"xmlns,attr"`
	NumFmts struct {
		Text   string `xml:",chardata"`
		Count  string `xml:"count,attr"`
		NumFmt []struct {
			Text       string `xml:",chardata"`
			NumFmtId   string `xml:"numFmtId,attr"`
			FormatCode string `xml:"formatCode,attr"`
		} `xml:"numFmt"`
	} `xml:"numFmts"`
	Fonts struct {
		Text  string `xml:",chardata"`
		Count string `xml:"count,attr"`
		Font  []struct {
			Text string `xml:",chardata"`
			Sz   struct {
				Text string `xml:",chardata"`
				Val  string `xml:"val,attr"`
			} `xml:"sz"`
			Name struct {
				Text string `xml:",chardata"`
				Val  string `xml:"val,attr"`
			} `xml:"name"`
			Family struct {
				Text string `xml:",chardata"`
				Val  string `xml:"val,attr"`
			} `xml:"family"`
			Charset struct {
				Text string `xml:",chardata"`
				Val  string `xml:"val,attr"`
			} `xml:"charset"`
			B struct {
				Text string `xml:",chardata"`
				Val  string `xml:"val,attr"`
			} `xml:"b"`
			Color struct {
				Text string `xml:",chardata"`
				Rgb  string `xml:"rgb,attr"`
			} `xml:"color"`
		} `xml:"font"`
	} `xml:"fonts"`
	Fills struct {
		Text  string `xml:",chardata"`
		Count string `xml:"count,attr"`
		Fill  []struct {
			Text        string `xml:",chardata"`
			PatternFill struct {
				Text        string `xml:",chardata"`
				PatternType string `xml:"patternType,attr"`
			} `xml:"patternFill"`
		} `xml:"fill"`
	} `xml:"fills"`
	Borders struct {
		Text   string `xml:",chardata"`
		Count  string `xml:"count,attr"`
		Border []struct {
			Text         string `xml:",chardata"`
			DiagonalUp   string `xml:"diagonalUp,attr"`
			DiagonalDown string `xml:"diagonalDown,attr"`
			Left         struct {
				Text  string `xml:",chardata"`
				Style string `xml:"style,attr"`
			} `xml:"left"`
			Right struct {
				Text  string `xml:",chardata"`
				Style string `xml:"style,attr"`
			} `xml:"right"`
			Top struct {
				Text  string `xml:",chardata"`
				Style string `xml:"style,attr"`
			} `xml:"top"`
			Bottom struct {
				Text  string `xml:",chardata"`
				Style string `xml:"style,attr"`
			} `xml:"bottom"`
			Diagonal struct {
				Text  string `xml:",chardata"`
				Style string `xml:"style,attr"`
			} `xml:"diagonal"`
		} `xml:"border"`
	} `xml:"borders"`
	CellStyleXfs struct {
		Text  string `xml:",chardata"`
		Count string `xml:"count,attr"`
		Xf    []struct {
			Text              string `xml:",chardata"`
			NumFmtId          string `xml:"numFmtId,attr"`
			FontId            string `xml:"fontId,attr"`
			FillId            string `xml:"fillId,attr"`
			BorderId          string `xml:"borderId,attr"`
			ApplyFont         string `xml:"applyFont,attr"`
			ApplyFill         string `xml:"applyFill,attr"`
			ApplyBorder       string `xml:"applyBorder,attr"`
			ApplyAlignment    string `xml:"applyAlignment,attr"`
			ApplyProtection   string `xml:"applyProtection,attr"`
			ApplyNumberFormat string `xml:"applyNumberFormat,attr"`
			Alignment         struct {
				Text         string `xml:",chardata"`
				Horizontal   string `xml:"horizontal,attr"`
				Vertical     string `xml:"vertical,attr"`
				TextRotation string `xml:"textRotation,attr"`
				WrapText     string `xml:"wrapText,attr"`
				Indent       string `xml:"indent,attr"`
				ShrinkToFit  string `xml:"shrinkToFit,attr"`
			} `xml:"alignment"`
			Protection struct {
				Text   string `xml:",chardata"`
				Locked string `xml:"locked,attr"`
				Hidden string `xml:"hidden,attr"`
			} `xml:"protection"`
		} `xml:"xf"`
	} `xml:"cellStyleXfs"`
	CellXfs struct {
		Text  string `xml:",chardata"`
		Count string `xml:"count,attr"`
		Xf    []struct {
			Text              string `xml:",chardata"`
			NumFmtId          string `xml:"numFmtId,attr"`
			FontId            string `xml:"fontId,attr"`
			FillId            string `xml:"fillId,attr"`
			BorderId          string `xml:"borderId,attr"`
			XfId              string `xml:"xfId,attr"`
			ApplyBorder       string `xml:"applyBorder,attr"`
			ApplyFont         string `xml:"applyFont,attr"`
			ApplyFill         string `xml:"applyFill,attr"`
			ApplyAlignment    string `xml:"applyAlignment,attr"`
			ApplyNumberFormat string `xml:"applyNumberFormat,attr"`
			ApplyProtection   string `xml:"applyProtection,attr"`
			Alignment         struct {
				Text         string `xml:",chardata"`
				Horizontal   string `xml:"horizontal,attr"`
				Vertical     string `xml:"vertical,attr"`
				WrapText     string `xml:"wrapText,attr"`
				TextRotation string `xml:"textRotation,attr"`
				Indent       string `xml:"indent,attr"`
				ShrinkToFit  string `xml:"shrinkToFit,attr"`
			} `xml:"alignment"`
			Protection struct {
				Text   string `xml:",chardata"`
				Locked string `xml:"locked,attr"`
				Hidden string `xml:"hidden,attr"`
			} `xml:"protection"`
		} `xml:"xf"`
	} `xml:"cellXfs"`
	CellStyles struct {
		Text      string `xml:",chardata"`
		Count     string `xml:"count,attr"`
		CellStyle []struct {
			Text          string `xml:",chardata"`
			Name          string `xml:"name,attr"`
			XfId          string `xml:"xfId,attr"`
			BuiltinId     string `xml:"builtinId,attr"`
			CustomBuiltin string `xml:"customBuiltin,attr"`
		} `xml:"cellStyle"`
	} `xml:"cellStyles"`
	Colors struct {
		Text          string `xml:",chardata"`
		IndexedColors struct {
			Text     string `xml:",chardata"`
			RgbColor []struct {
				Text string `xml:",chardata"`
				Rgb  string `xml:"rgb,attr"`
			} `xml:"rgbColor"`
		} `xml:"indexedColors"`
	} `xml:"colors"`
} 

