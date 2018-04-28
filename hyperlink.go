package xls

import "fmt"

var (
	_ contentHandler = &HyperLink{}
)

// HyperLink represents a hyperlink's content
type HyperLink struct {
	CellRange
	Description      string
	TextMark         string
	TargetFrame      string
	Url              string
	ShortedFilePath  string
	ExtendedFilePath string
	IsUrl            bool
}

// Debug prints the needed file dump
func (h *HyperLink) Debug(wb *WorkBook) {
	fmt.Printf("hyper link col dump:%#+v\n", h)
}

// String gets the hyperlink string, use the public variable Url to get the original Url
func (h *HyperLink) String(wb *WorkBook) []string {
	res := make([]string, h.LastColB-h.FristColB+1)
	var str string
	if h.IsUrl {
		str = fmt.Sprintf("%s(%s)", h.Description, h.Url)
	} else {
		str = h.ExtendedFilePath
	}

	for i := uint16(0); i < h.LastColB-h.FristColB+1; i++ {
		res[i] = str
	}
	return res
}
