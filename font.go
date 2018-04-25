package xls

// FontInfo represents the font info
type FontInfo struct {
	Height     uint16
	Flag       uint16
	Color      uint16
	Bold       uint16
	Escapement uint16
	Underline  byte
	Family     byte
	Charset    byte
	Notused    byte
	NameB      byte
}

// Font represents the font
type Font struct {
	Info *FontInfo
	Name string
}
