package model

import (
	"database/sql"
	"extract-blocks/s3"
	"extract-blocks/utils"
	"fmt"
	"os"
	"path"
	"path/filepath"

	log "github.com/Sirupsen/logrus"
	"github.com/nad2000/xlsx"
)

// Proplem - TODO: ...
type Problem struct {
	ID             int    // `gorm:"column:FileID;primary_key:true;AUTO_INCREMENT"`
	NumberOfSheets int    `gorm:"column:Number_of_sheets"`
	Name           string //`gorm:"column:S3BucketName;size:100"`
	Category       sql.NullString
	IsProcessed    bool `gorm:"column:IsProcessed;default:0"`
	Marks          sql.NullFloat64
	SourceID       int           `gorm:"column:FileID;index"`
	Source         *Source       `gorm:"foreignkey:SourceID"`
	ReferenceID    sql.NullInt64 `gorm:"index;type:int"`
}

// TableName overrides default table name for the model
func (Problem) TableName() string {
	return "Problems"
}

// ProblemSheet - TODO: ...
type ProblemSheet struct {
	ID             int
	SequenceNumber int `gorm:"column:Sequence_Number"`
	Name           string
	ProblemID      int      `gorm:"index"`
	Problem        *Problem `gorm:"foreignkey:ProblemID"`
}

// TableName overrides default table name for the model
func (ProblemSheet) TableName() string {
	return "ProblemWorkSheets"
}

// ProblemSheetData - TODO: ...
type ProblemSheetData struct {
	ID             int
	Range          string `gorm:"column:CellRange;size:10"`
	Value          sql.NullString
	Comment        sql.NullString
	Formula        sql.NullString
	IsReference    bool
	ProblemID      int           `gorm:"index"`
	Problem        *Problem      `gorm:"foreignkey:ProblemID"`
	ProblemSheetID int           `gorm:"column:ProblemWorkSheet_ID;index"`
	ProblemSheet   *ProblemSheet `gorm:"foreignkey:ProblemSheetID"`
}

// TableName overrides default table name for the model
func (ProblemSheetData) TableName() string {
	return "ProblemWorkSheetExcelData"
}

func (p Problem) String() string {
	return fmt.Sprintf("Problem{ID: %d}", p.ID)
}

// ImportFile imports form Excel file QuestionExcleData
func (p *Problem) ImportFile(fileName, color string, verbose bool, manager s3.FileManager) error {
	file, err := xlsx.OpenFile(fileName)
	if err != nil {
		return err
	}

	if VerboseLevel > 0 {
		log.Infof("Processing workbook: %s", fileName)
	}

	var sheetCount int
	for sqn, sheet := range file.Sheets {

		if sheet.Hidden {
			log.Infof("Skipping hidden worksheet %q", sheet.Name)
			continue
		}

		sheetCount++
		if VerboseLevel > 0 {
			log.Infof("Processing worksheet %q", sheet.Name)
		}

		var ps = ProblemSheet{
			ProblemID:      p.ID,
			Name:           sheet.Name,
			SequenceNumber: sqn,
		}
		if err := Db.Create(&ps).Error; err != nil {
			return err
		}

		for i, row := range sheet.Rows {
			for j, cell := range row.Cells {

				var commentText string

				cellRange := CellAddress(i, j)
				commen, ok := sheet.Comment[cellRange]
				if ok {
					commentText = commen.Text
				} else {
					commentText = ""
				}
				var psd = ProblemSheetData{
					ProblemID:      p.ID,
					ProblemSheetID: ps.ID,
					Range:          cellRange,
					Value:          NewNullString(cell.Value),
					Formula:        NewNullString(cell.Formula()),
					Comment:        NewNullString(commentText),
				}
				if err := Db.Create(&psd).Error; err != nil {
					return err
				}
			}
		}

		sheet.Cell(1, 0).SetInt(p.ID)
		for i := 2; i < 5; i++ {
			sheet.Cell(1, i).SetInt(ps.ID)
		}
	}
	p.ImportBlocks(file, color, verbose)

	// Choose the output file name
	outputName := path.Join(os.TempDir(), filepath.Base(fileName))
	file.Save(outputName)

	// Upload the file
	newKey := utils.NewS3Key() + filepath.Ext(fileName)

	location, err := manager.Upload(outputName, p.Source.S3BucketName, newKey)
	if err != nil {
		return fmt.Errorf("failed to uploade the output file %q to %q with S3 key %q: %s",
			outputName, p.Source.S3BucketName, newKey, err)
	}
	log.Infof("Output file %q uploaded to bucket %q with S3 key %q, location: %q",
		outputName, p.Source.S3BucketName, newKey, location)

	var s Source
	if err := Db.FirstOrCreate(&s, Source{
		S3BucketName: p.Source.S3BucketName,
		S3Key:        newKey,
		FileName:     filepath.Base(outputName),
		ContentType:  p.Source.ContentType,
		FileSize:     p.Source.FileSize,
	}).Error; err != nil {
		log.Error(err)
		return err
	}
	if err := Db.Model(p).UpdateColumns(map[string]interface{}{
		"Number_of_sheets": sheetCount,
		"FileID":           s.ID,
	}).Error; err != nil {
		log.Error(err)
		return err
	}
	return nil
}

// ImportBlocks extracts blocks from the given question file and stores in the DB for referencing
func (p *Problem) ImportBlocks(file *xlsx.File, color string, verbose bool) (wb Workbook) {

	// var source Source
	// Db.Model(&p).Related(&source, "Source")
	fileName := p.Source.FileName
	if !DryRun {
		wb = Workbook{FileName: fileName, IsReference: true}

		if err := Db.Create(&wb).Error; err != nil {
			log.WithError(err).Errorf("failed to create workbook entry %#v", wb)
			return
		}
		if DebugLevel > 1 {
			log.Debugf("Created workbook entry %#v", wb)
		}
	}

	if verbose {
		log.Infof("*** Processing workbook: %s", fileName)
	}

	for orderNum, sheet := range file.Sheets {

		if sheet.Hidden {
			log.Infof("Skipping hidden worksheet %q", sheet.Name)
			continue
		}

		if verbose {
			log.Infof("Processing worksheet %q", sheet.Name)
		}

		var ws Worksheet
		if !DryRun {
			Db.FirstOrCreate(&ws, Worksheet{
				Name:             sheet.Name,
				WorkbookID:       wb.ID,
				WorkbookFileName: fileName,
				IsReference:      true,
				OrderNum:         orderNum,
			})
		}
		if Db.Error != nil {
			log.Fatalf("*** Failed to create worksheet entry: %s", Db.Error.Error())
		}

		blocks := blockList{}
		sheetFillColors := []string{}

		for i, row := range sheet.Rows {
			for j, cell := range row.Cells {

				if blocks.includes(i, j) {
					continue
				}
				style := cell.GetStyle()
				fgColor := style.Fill.FgColor
				if fgColor != "" {
					for _, c := range sheetFillColors {
						if c == fgColor {
							goto MATCH
						}
					}
					sheetFillColors = append(sheetFillColors, fgColor)
				}
			MATCH:

				if fgColor == color {

					b := Block{
						WorksheetID:     ws.ID,
						Color:           color,
						Formula:         cell.Formula(),
						RelativeFormula: RelativeFormula(i, j, cell.Formula()),
						IsReference:     true,
						TRow:            i,
						LCol:            j,
					}

					if !DryRun {
						Db.Create(&b)
					}

					if DebugLevel > 1 {
						log.Debugf("Created %#v", b)
					}

					b.findWhole(sheet, color)
					b.save()
					blocks = append(blocks, b)
					if verbose && b.Range != "" {
						log.Infof("Found: %s", b)
					}

				}
			}
		}
		if len(blocks) == 0 {
			log.Warningf("No block found in the worksheet %q of the workbook %q with color %q", sheet.Name, fileName, color)
			if len(sheetFillColors) > 0 {
				log.Infof("Following colors were found in the worksheet you could use: %v", sheetFillColors)
			}
		}
	}

	if !DryRun {
		p.ReferenceID = NewNullInt64(wb.ID)
		Db.Save(&p)
	}

	return
}
