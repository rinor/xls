package xls

import (
	"encoding/binary"
	"io"
	"unicode/utf16"
)

// bof is the information unit in xls file
type bof struct {
	Id   uint16
	Size uint16
}

// utf16String read the utf16 string from reader
func (b *bof) utf16String(buf io.ReadSeeker, count uint32) string {
	var bts = make([]uint16, count)
	binary.Read(buf, binary.LittleEndian, &bts)
	return utf16String(bts)
}

func utf16String(bts []uint16) string {
	for i, v := range bts {
		if v == '\x00' {
			bts = bts[:i]
			break
		}
	}
	runes := utf16.Decode(bts)
	return string(runes)
}

func byteString(bs []byte) string {
	var bts = make([]uint16, len(bs))
	for k, v := range bs {
		bts[k] = uint16(v)
	}
	return utf16String(bts)
}

type biffHeader struct {
	Ver    uint16
	Type   uint16
	IdMake uint16
	Year   uint16
	Flags  uint32
	MinVer uint32
}
