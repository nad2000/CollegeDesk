package xlsx; import "encoding/xml"
// PivotCacheDefinition was generated 2019-02-12 00:21:04 by rcir178 on rcir178-Latitude-E7470.
type PivotCacheDefinition struct {
	XMLName               xml.Name `xml:"pivotCacheDefinition"`
	Text                  string   `xml:",chardata"`
	Xmlns                 string   `xml:"xmlns,attr"`
	R                     string   `xml:"r,attr"`
	ID                    string   `xml:"id,attr"`
	RefreshedBy           string   `xml:"refreshedBy,attr"`
	RefreshedDate         string   `xml:"refreshedDate,attr"`
	CreatedVersion        string   `xml:"createdVersion,attr"`
	RefreshedVersion      string   `xml:"refreshedVersion,attr"`
	MinRefreshableVersion string   `xml:"minRefreshableVersion,attr"`
	RecordCount           string   `xml:"recordCount,attr"`
	CacheSource           struct {
		Text            string `xml:",chardata"`
		Type            string `xml:"type,attr"`
		WorksheetSource struct {
			Text  string `xml:",chardata"`
			Ref   string `xml:"ref,attr"`
			Sheet string `xml:"sheet,attr"`
		} `xml:"worksheetSource"`
	} `xml:"cacheSource"`
	CacheFields struct {
		Text       string `xml:",chardata"`
		Count      string `xml:"count,attr"`
		CacheField []struct {
			Text        string `xml:",chardata"`
			Name        string `xml:"name,attr"`
			NumFmtId    string `xml:"numFmtId,attr"`
			SharedItems struct {
				Text                   string `xml:",chardata"`
				ContainsSemiMixedTypes string `xml:"containsSemiMixedTypes,attr"`
				ContainsNonDate        string `xml:"containsNonDate,attr"`
				ContainsDate           string `xml:"containsDate,attr"`
				ContainsString         string `xml:"containsString,attr"`
				MinDate                string `xml:"minDate,attr"`
				MaxDate                string `xml:"maxDate,attr"`
				Count                  string `xml:"count,attr"`
				ContainsNumber         string `xml:"containsNumber,attr"`
				ContainsInteger        string `xml:"containsInteger,attr"`
				MinValue               string `xml:"minValue,attr"`
				MaxValue               string `xml:"maxValue,attr"`
				D                      []struct {
					Text string `xml:",chardata"`
					V    string `xml:"v,attr"`
				} `xml:"d"`
				S []struct {
					Text string `xml:",chardata"`
					V    string `xml:"v,attr"`
				} `xml:"s"`
			} `xml:"sharedItems"`
		} `xml:"cacheField"`
	} `xml:"cacheFields"`
	ExtLst struct {
		Text string `xml:",chardata"`
		Ext  struct {
			Text                 string `xml:",chardata"`
			URI                  string `xml:"uri,attr"`
			X14                  string `xml:"x14,attr"`
			PivotCacheDefinition struct {
				Text string `xml:",chardata"`
			} `xml:"pivotCacheDefinition"`
		} `xml:"ext"`
	} `xml:"extLst"`
} 

