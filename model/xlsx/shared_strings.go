package xlsx; import "encoding/xml"
// Sst was generated 2019-04-06 22:26:52 by rcir178 on rcir178-Latitude-E7470.
type Sst struct {
	XMLName     xml.Name `xml:"sst"`
	Text        string   `xml:",chardata"`
	Xmlns       string   `xml:"xmlns,attr"`
	Count       string   `xml:"count,attr"`
	UniqueCount string   `xml:"uniqueCount,attr"`
	Si          []struct {
		Text string `xml:",chardata"`
		T    struct {
			Text string `xml:",chardata"` // Month, Salesman, Region, ...
		} `xml:"t"`
	} `xml:"si"`
} 

