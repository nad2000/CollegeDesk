package main

import (
	"fmt"

	log "github.com/Sirupsen/logrus"
	"github.com/tealeg/xlsx"
)

type Range struct {
	S *xlsx.Cell // Top-left cell
	E *xlsx.Cell //  Bottom-right cell
}

func main() {
	ranges := []Range{}
	excelFileName := "./demo.xlsx"
	xlFile, err := xlsx.OpenFile(excelFileName)
	if err != nil {
		log.Error(err)
	}
	for _, df := range xlFile.DefinedNames {
		log.Info(df)
	}
	for _, sheet := range xlFile.Sheets {
		for i, row := range sheet.Rows {
			fmt.Printf("\n\nROW %d\n=========\n", i)
			for j, cell := range row.Cells {
				// text := cell.String()
				if style := cell.GetStyle(); style.Fill.FgColor == "FFFFFF00" && cell.Formula() != "" {
					fmt.Printf("CELL (%d:%d) %#v: %s\n", i, j, *style.NamedStyleIndex, cell.Formula())
					ranges = append(ranges, Range{cell, cell})
				}
			}
			fmt.Println()
		}
	}
	log.Info(ranges)
}
