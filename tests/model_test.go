package tests

import (
	"extract-blocks/model"
	"testing"

	"github.com/nad2000/excelize"
	"github.com/nad2000/xlsx"
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

// TestFilters
func TestFilters(t *testing.T) {
	file, _ := excelize.OpenFile("Filter ALL TYPES.xlsx")
	t.Logf("%+v", file.WorkBook.Sheets)
	for _, sheet := range file.WorkBook.Sheets.Sheet {
		name := "xl/worksheets/sheet" + sheet.SheetID + ".xml"
		s := model.UnmarshalAutoFilter(file.XLSX[name])
		t.Log(s.AutoFilter)
	}
}
