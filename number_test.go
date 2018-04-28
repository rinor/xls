package xls

import (
	"strconv"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestNumber(t *testing.T) {
	xlFile, err := Open("testdata/number.xls", "utf-8")
	assert.NoError(t, err)
	assert.EqualValues(t, 1, xlFile.NumSheets())

	sheet1 := xlFile.GetSheet(0)
	assert.NotNil(t, sheet1)
	assert.EqualValues(t, 1, int(sheet1.MaxRow))

	var results = [][]string{
		{"123.45", "3.99E-01"},
		{"1.54E-01", "1.24E+00"},
	}

	for i := 0; i <= int(sheet1.MaxRow); i++ {
		row := sheet1.Row(i)
		assert.NotNil(t, row)
		assert.EqualValues(t, 0, row.FirstCol())
		assert.EqualValues(t, 2, row.LastCol())

		for j := row.FirstCol(); j < row.LastCol(); j++ {
			ce := row.Column(j)
			assert.True(t, ce.IsValid())
			assert.True(t, ce.IsNumber() || ce.IsRk())
			assert.False(t, ce.IsHyperLink())
			assert.False(t, ce.IsFormula())
			assert.False(t, ce.IsBlank())
			if ce.IsNumber() {
				num := ce.MustNumber()
				assert.NotNil(t, num)
				assert.EqualValues(t, results[i][j], strconv.FormatFloat(num.Float, 'E', 2, 64))
			} else if ce.IsRk() {
				rk := ce.MustRk()
				assert.NotNil(t, rk)
				assert.EqualValues(t, results[i][j], rk.String(xlFile)[0])
				assert.EqualValues(t, results[i][j], row.Col(j))
			}
		}

		/*ce = row.Column(1)
		assert.True(t, ce.IsValid())
		assert.False(t, ce.IsRk())
		assert.True(t, ce.IsNumber())
		assert.False(t, ce.IsHyperLink())
		assert.False(t, ce.IsFormula())
		assert.False(t, ce.IsBlank())
		num := ce.MustNumber()
		assert.NotNil(t, num)
		assert.EqualValues(t, results[i][1], num.String(xlFile)[0])
		assert.EqualValues(t, results[i][1], row.Col(1))*/
	}
}
