package xls

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestTime(t *testing.T) {
	f, err := Open("./testdata/time.xls", "utf-8")
	assert.NoError(t, err)
	assert.EqualValues(t, 1, f.NumSheets())

	sheet1 := f.GetSheet(0)
	assert.NotNil(t, sheet1)
	assert.EqualValues(t, 5, int(sheet1.MaxRow))
	assert.EqualValues(t, "Sheet1", sheet1.Name)

	var results = [][]string{
		{"04-21-16", "09:35:07"},
		{"05-17-16", "14:42:12"},
		{"05-18-16", "16:28:19"},
		{"05-19-16", "19:22:50"},
		{"08-16-16", "23:11:36"},
		{"09-13-16", "10:17:54"},
	}

	for i := 0; i <= int(sheet1.MaxRow); i++ {
		row := sheet1.Row(i)
		assert.NotNil(t, row)
		assert.EqualValues(t, 0, row.FirstCol())
		assert.EqualValues(t, 2, row.LastCol())
		for j := row.FirstCol(); j < row.LastCol(); j++ {
			assert.EqualValues(t, results[i][j], row.Col(j))
		}
	}
}
