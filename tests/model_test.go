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

		t.Log("***", sheet.Name)
		for i, row := range sheet.Rows {
			for j, cell := range row.Cells {
				t.Logf("[%d,%d]: %s", i, j, cell)
			}
		}
	}
}
