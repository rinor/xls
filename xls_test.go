package xls

import (
	"bytes"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestEuropeString(t *testing.T) {
	bs := []byte{66, 233, 114, 232}
	assert.EqualValues(t, "Bérè", byteString(bs))
}

func TestBof(t *testing.T) {
	b := new(bof)
	b.Id = 0x41E
	b.Size = 55
	buf := bytes.NewReader([]byte{0x07, 0x00, 0x19, 0x00, 0x01, 0x22, 0x00, 0xE5, 0xFF, 0x22, 0x00, 0x23, 0x00, 0x2C, 0x00, 0x23, 0x00, 0x23, 0x00, 0x30, 0x00, 0x2E, 0x00, 0x30, 0x00, 0x30, 0x00, 0x3B, 0x00, 0x22, 0x00, 0xE5, 0xFF, 0x22, 0x00, 0x5C, 0x00, 0x2D, 0x00, 0x23, 0x00, 0x2C, 0x20, 0x00})
	wb := new(WorkBook)
	wb.Formats = make(map[uint16]*Format)
	wb.parseBof(buf, b, b, 0)
}

func TestMaxRow(t *testing.T) {
	xlFile, err := Open("testdata/Table.xls", "utf-8")
	assert.NoError(t, err)
	assert.EqualValues(t, 1, xlFile.NumSheets())

	sheet1 := xlFile.GetSheet(0)
	assert.NotNil(t, sheet1)
	assert.EqualValues(t, 11, sheet1.MaxRow)
}
