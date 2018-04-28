package xls

import (
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestHyperlink(t *testing.T) {
	f, err := Open("./testdata/hyperlink.xls", "utf-8")
	assert.NoError(t, err)
	assert.EqualValues(t, 1, f.NumSheets())

	sheet1 := f.GetSheet(0)
	assert.NotNil(t, sheet1)
	assert.EqualValues(t, 1, int(sheet1.MaxRow))
	assert.EqualValues(t, "Sheet1", sheet1.Name)

	var results = [][]string{
		{"abcdefg@123.com(mailto:abcdefg@123.com)"},
		{"gfedcba@123.com(mailto:gfedcba@123.com)"},
	}

	for i := 0; i <= int(sheet1.MaxRow); i++ {
		row := sheet1.Row(i)
		assert.NotNil(t, row)
		assert.EqualValues(t, 0, row.FirstCol())
		assert.EqualValues(t, 1, row.LastCol())
		for j := row.FirstCol(); j < row.LastCol(); j++ {
			ce := row.Column(j)
			assert.True(t, ce.IsValid())
			assert.True(t, ce.IsHyperLink())
			assert.False(t, ce.IsNumber())
			assert.False(t, ce.IsFormula())
			hl := ce.MustHpyerLink()
			assert.NotNil(t, hl)
			assert.True(t, hl.IsUrl)
			assert.EqualValues(t, strings.Split(results[i][j], "(")[0], hl.Description)
			assert.EqualValues(t, strings.TrimRight(strings.Split(results[i][j], "(")[1], ")"), hl.Url)
			assert.EqualValues(t, results[i][j], row.Col(j))
		}
	}
}
