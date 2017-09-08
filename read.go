package main

import (
	"fmt"
	"os"
	"strconv"

	"github.com/xuri/excelize"
)

func main() {
	xlsx, err := excelize.OpenFile("./demo.xlsx")
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
	// Get value from cell by given sheet index and axis.
	cell := xlsx.GetCellValue("Sheet1", "B2")
	fmt.Println(cell)
	// Get sheet index.
	index := xlsx.GetSheetIndex("Sheet1")
	// Get all the rows in a sheet.
	rows := xlsx.GetRows("sheet" + strconv.Itoa(index))
	for _, row := range rows {

		for _, colCell := range row {
			fmt.Printf("%#v\t", colCell)
		}
		fmt.Println()
	}
}
