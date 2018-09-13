package tests

import (
	"testing"

	"github.com/tealeg/xlsx"
)

func TestMerged(t *testing.T) {
	file, err := xlsx.OpenFile("merged.xlsx")
	if err != nil {
		t.Error(err)
	}

	for _, sheet := range file.Sheets {
		var count int
		for _, row := range sheet.Rows {
			count += len(row.Cells)
		}
		if expected := 352; expected != count {
			t.Errorf("Expeced %d cells, got: %d", expected, count)
		}
	}
}
