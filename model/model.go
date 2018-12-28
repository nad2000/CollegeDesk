package model

import (
	"database/sql"
	"database/sql/driver"
	"encoding/xml"
	"errors"
	"extract-blocks/s3"
	"fmt"
	"path"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	log "github.com/Sirupsen/logrus"
	"github.com/jinzhu/gorm"

	x "extract-blocks/model/xlsx"
	//"github.com/tealeg/xlsx"
	"github.com/nad2000/excelize"
	"github.com/nad2000/xlsx"
)

// Db - shared DB connection
var Db *gorm.DB

// VerboseLevel - the level of verbosity
var VerboseLevel int

// DebugLevel - the level of verbosity of the debug information
var DebugLevel int

// DryRun - perform processing without actually updating or changing files
var DryRun bool

var cellIDRe = regexp.MustCompile("\\$?[A-Z]+\\$?[0-9]+")

// CellAddress maps a cell coordiantes (row, column) to its address
func CellAddress(rowIndex, colIndex int) string {
	return xlsx.GetCellIDStringFromCoords(colIndex, rowIndex)
}

// RelCellAddress - relative cell R1R1 representation against the given cell
func RelCellAddress(address string, rowIncrement, colIncrement int) (string, error) {
	colIndex, rowIndex, err := xlsx.GetCoordsFromCellIDString(address)
	if err != nil {
		log.WithError(err).Error("Failed to map address ", address)
		return "", err
	}
	return xlsx.GetCellIDStringFromCoords(colIndex+colIncrement, rowIndex+rowIncrement), nil
}

// RelativeCellAddress converts cell ID into a relative R1C1 representation
func RelativeCellAddress(rowIndex, colIndex int, cellID string) string {
	x, y, err := xlsx.GetCoordsFromCellIDString(cellID)
	if err != nil {
		log.WithError(err).Errorln("Failed to find coordinates for ", cellID)
	}
	var r1c1 string

	if strings.Contains(cellID[1:], "$") {
		r1c1 = fmt.Sprintf("R[%d]", y)
	} else {
		r1c1 = fmt.Sprintf("R[%+d]", y-rowIndex)
	}

	if cellID[0] == '$' {
		r1c1 += fmt.Sprintf("C[%d]", x)
	} else {
		r1c1 += fmt.Sprintf("C[%+d]", x-colIndex)
	}
	//return fmt.Sprintf("R[%d]C[%d]", y-rowIndex, x-colIndex)
	return r1c1
}

// RelativeFormula transforms the cell formula into the relative in R1C1 notation
func RelativeFormula(rowIndex, colIndex int, formula string) string {
	cellIDs := cellIDRe.FindAllString(formula, -1)
	for _, cellID := range cellIDs {
		relCellID := RelativeCellAddress(rowIndex, colIndex, cellID)
		if DebugLevel > 1 {
			log.Debugf("Replacing %q with %q at (%d, %d)", cellID, relCellID, rowIndex, colIndex)
		}
		formula = strings.Replace(formula, cellID, relCellID, -1)
	}
	return formula
}

// QuestionType - workaround for MySQL EMUM(...)
type QuestionType string

// Scan - workaround for MySQL EMUM(...)
func (qt *QuestionType) Scan(value interface{}) error { *qt = QuestionType(value.([]byte)); return nil }

// Value - workaround for MySQL EMUM(...)
func (qt QuestionType) Value() (driver.Value, error) { return string(qt), nil }

// NewNullInt64 - a helper function that makes nullable from a plain int or a string
func NewNullInt64(value interface{}) sql.NullInt64 {
	switch value.(type) {
	case int:
		return sql.NullInt64{Valid: true, Int64: int64(value.(int))}
	case string:
		if value.(string) == "" {
			return sql.NullInt64{}
		}
		v, _ := strconv.Atoi(value.(string))
		return sql.NullInt64{Valid: true, Int64: int64(v)}
	}
	return sql.NullInt64{}
}

// Question - questions
type Question struct {
	ID                 int            `gorm:"column:QuestionID;primary_key:true;AUTO_INCREMENT"`
	QuestionType       QuestionType   `gorm:"column:QuestionType"`
	QuestionSequence   int            `gorm:"column:QuestionSequence;not null"`
	QuestionText       string         `gorm:"column:QuestionText;type:text;not null"`
	AnswerExplanation  sql.NullString `gorm:"column:AnswerExplanation;type:text"`
	MaxScore           float32        `gorm:"column:MaxScore;type:float;not null"`
	AuthorUserID       int            `gorm:"column:AuthorUserID;not null"`
	WasCompared        bool
	IsProcessed        bool `gorm:"column:IsProcessed;default:0"`
	Source             Source
	SourceID           sql.NullInt64       `gorm:"column:FileID;type:int"`
	Answers            []Answer            `gorm:"ForeignKey:QuestionID"`
	QuestionExcelDatas []QuestionExcelData `gorm:"ForeignKey:QuestionID"`
	ReferenceID        sql.NullInt64       `gorm:"index;type:int"`
}

// TableName overrides default table name for the model
func (Question) TableName() string {
	return "Questions"
}

func (q Question) String() string {
	return fmt.Sprintf("Question{ID: %d, Type: %s, Text: %q}",
		q.ID, q.QuestionType, q.QuestionText)
}

// ImportFile imports form Excel file QuestionExcleData
func (q *Question) ImportFile(fileName, color string, verbose bool) error {
	file, err := xlsx.OpenFile(fileName)
	if err != nil {
		return err
	}

	if VerboseLevel > 0 {
		log.Infof("Processing workbook: %s", fileName)
	}

	for _, sheet := range file.Sheets {

		if sheet.Hidden {
			log.Infof("Skipping hidden worksheet %q", sheet.Name)
			continue
		}

		if VerboseLevel > 0 {
			log.Infof("Processing worksheet %q", sheet.Name)
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
				var qed QuestionExcelData

				Db.FirstOrCreate(&qed, QuestionExcelData{
					QuestionID: q.ID,
					SheetName:  sheet.Name,
					CellRange:  cellRange,
					Value:      cell.Value,
					Formula:    cell.Formula(),
					Comment:    commentText,
				})
				if Db.Error != nil {
					return Db.Error
				}
			}
		}

	}
	q.ImportBlocks(file, color, verbose)

	return nil
}

// ImportBlocks extracts blocks from the given question file and stores in the DB for referencing
func (q *Question) ImportBlocks(file *xlsx.File, color string, verbose bool) (wb Workbook) {

	var source Source
	Db.Model(&q).Related(&source, "Source")
	fileName := source.FileName
	if !DryRun {
		wb = Workbook{FileName: fileName, IsReference: true}

		if err := Db.Create(&wb).Error; err != nil {
			log.WithError(err).Errorf("Failed to create workbook entry %#v", wb)
			return
		}
		if DebugLevel > 1 {
			log.Debugf("Ceated workbook entry %#v", wb)
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

				if blocks.wasFound(i, j) {
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
			log.Warningf("No block found ot the worksheet %q of the workbook %q with color %q", sheet.Name, fileName, color)
			if len(sheetFillColors) > 0 {
				log.Infof("Following colors were found in the worksheet you could use: %v", sheetFillColors)
			}
		}
	}

	if !DryRun {
		q.ReferenceID = NewNullInt64(wb.ID)
		Db.Save(&q)
	}

	return
}

// QuestionExcelData - extracted celles from question Workbooks
type QuestionExcelData struct {
	ID         int    `gorm:"column:Id;primary_key:true;AUTO_INCREMENT"`
	SheetName  string `gorm:"column:SheetName"`
	CellRange  string `gorm:"column:CellRange;size:10"`
	Value      string `gorm:"column:Value;size:2000"`
	Comment    string `gorm:"column:Comment"`
	Formula    string `gorm:"column:Formula"`
	Question   Question
	QuestionID int `gorm:"column:QuestionID"`
}

// TableName overrides default table name for the model
func (QuestionExcelData) TableName() string {
	return "QuestionExcelData"
}

// Source - student answer file sources
type Source struct {
	ID           int        `gorm:"column:FileID;primary_key:true;AUTO_INCREMENT"`
	S3BucketName string     `gorm:"column:S3BucketName;size:100"`
	S3Key        string     `gorm:"column:S3Key;size:100"`
	FileName     string     `gorm:"column:FileName;size:100"`
	ContentType  string     `gorm:"column:ContentType;size:100"`
	FileSize     int64      `gorm:"column:FileSize"`
	Answers      []Answer   `gorm:"ForeignKey:FileID"`
	Questions    []Question `gorm:"ForeignKey:FileID"`
}

// TableName overrides default table name for the model
func (Source) TableName() string {
	return "FileSources"
}

// DownloadTo - download and store source file to a specified directory
func (s Source) DownloadTo(manager s3.FileManager, dest string) (fileName string, err error) {
	destinationName := path.Join(dest, s.FileName)
	log.Infof(
		"Downloading %q (%q) form %q into %q",
		s.S3Key, s.FileName, s.S3BucketName, destinationName)
	fileName, err = manager.Download(
		s.FileName, s.S3BucketName, s.S3Key, destinationName)
	if err != nil {
		err = fmt.Errorf(
			"Failed to retrieve file %q from %q into %q: %s",
			s.S3Key, s.S3BucketName, destinationName, err.Error())
	}
	return
}

// Answer - student submitted answers
type Answer struct {
	ID                  int `gorm:"column:StudentAnswerID;primary_key:true;AUTO_INCREMENT"`
	Assignment          Assignment
	AssignmentID        int           `gorm:"column:StudentAssignmentID"`
	MCQOptionID         sql.NullInt64 `gorm:"column:MCQOptionID;type:int"`
	ShortAnswer         string        `gorm:"column:ShortAnswerText;type:text"`
	Marks               float64       `gorm:"column:Marks;type:float"`
	SubmissionTime      time.Time     `gorm:"column:SubmissionTime;default:NULL"`
	Worksheets          []Worksheet   `gorm:"ForeignKey:AnswerID"`
	Source              Source        `gorm:"Association_ForeignKey:FileID"`
	SourceID            sql.NullInt64 `gorm:"column:FileID;type:int"`
	Question            Question
	QuestionID          sql.NullInt64 `gorm:"column:QuestionID;type:int"`
	WasCommentProcessed uint8         `gorm:"type:tinyint;default:0"`
	WasXLProcessed      uint8         `gorm:"type:tinyint;default:0"`
	WasAutocommented    bool
	AnswerComments      []AnswerComment `gorm:"ForeignKey:AnswerID"`
}

// TableName overrides default table name for the model
func (Answer) TableName() string {
	return "StudentAnswers"
}

// Assignment - assigment
type Assignment struct {
	ID                 int    `gorm:"column:AssignmentID;primary_key:true;AUTO_INCREMENT"`
	Title              string `gorm:"column:Title;type:varchar(80)"`
	AssignmentSequence int    `gorm:"column:AssignmentSequence"`
	// StartDateAndTime   time.Time `gorm:"column:StartDateAndTime"`
	// DueDateAndTime     time.Time `gorm:"column:DueDateAndTime"`
	// UpdateTime         time.Time `gorm:"column:UpdateTime"`
	// IsHidden           int8      `gorm:"type:tinyint(4)"`
	// TotalMarks         float64   `gorm:"column:TotalMarks;type:float"`
	// TotalQuestion      int       `gorm:"column:TotalQuestion"`
	// CourseID           int      `gorm:"column:CourseID"`
	State        string `gorm:"column:State"` // `gorm:"column:State;type:enum('UNDER_CREATION','CREATED','READY_FOR_GRADING','GRADED')"`
	WasProcessed int8   `gorm:"type:tinyint(1)"`
}

// TableName overrides default table name for the model
func (Assignment) TableName() string {
	return "CourseAssignments"
}

// Workbook - Excel file / workbook
type Workbook struct {
	ID          int `gorm:"primary_key:true"`
	FileName    string
	CreatedAt   time.Time
	AnswerID    sql.NullInt64 `gorm:"column:StudentAnswerID;index;type:int"`
	Answer      Answer        `gorm:"foreignkey:AnswerID"`
	Worksheets  []Worksheet   // `gorm:"foreignkey:WorkbookID"`
	IsReference bool          // the workbook is used for referencing the expected bloks
}

// TableName overrides default table name for the model
func (Workbook) TableName() string {
	return "WorkBooks"
}

// Reset deletes all underlying objects: worksheets, blocks, and cells
func (wb *Workbook) Reset() {

	var worksheets []Worksheet

	if err := Db.Where("workbook_id = ?", wb.ID).Find(&worksheets).Error; err != nil {
		log.WithError(err).Errorln("Couldn't find the record of the Workbook, ID:", wb.ID)
	}
	log.Debugf("Deleting worksheets: %#v", worksheets)
	for _, ws := range worksheets {
		Db.Delete(Chart{}, "worksheet_id = ?", ws.ID)
		var blocks []Block
		Db.Model(&ws).Related(&blocks)
		if err := Db.Where("worksheet_id = ?", ws.ID).Find(&blocks).Error; err != nil {
			log.WithError(err).Error("Failed to find any blocks of the worksheet", ws)
		}
		for _, b := range blocks {
			log.Debugf("Deleting blocks: %#v", blocks)
			Db.Where("block_id = ?", b.ID).Delete(Cell{})
			Db.Delete(b)
		}
	}
	Db.Where("workbook_id = ?", wb.ID).Delete(Worksheet{})
}

// ImportComments - import comments from workbook file
func (wb *Workbook) ImportComments(fileName string) (err error) {

	xlsx, err := excelize.OpenFile(fileName)
	if err != nil {
		return fmt.Errorf("Failed to open file %q: %s", fileName, err.Error())
	}
	for name, comments := range xlsx.GetComments() {
		var ws Worksheet
		Db.First(&ws, Worksheet{
			Name:       name,
			AnswerID:   wb.AnswerID,
			WorkbookID: wb.ID,
		})
		authors := comments.Authors
		for _, c := range comments.CommentList.Comment {
			var a string
			if len(authors) > 0 && c.AuthorID >= 0 && c.AuthorID < len(authors) {
				a = strings.Split(authors[c.AuthorID].Author, ":")[0]
			}
			for _, t := range c.Text.R {
				if a != "" && !strings.HasPrefix(t.T, a) {
					log.Debugf("*** [%s] %s: %s", c.Ref, a, t.T)
					var cell Cell
					result := Db.First(&cell, "worksheet_id = ? AND range = ?", ws.ID, c.Ref)
					if !result.RecordNotFound() {
						Db.Model(&cell).UpdateColumn("comment", t.T)
					}
				}
			}

		}
	}
	return
}

// SharedStrings - workbook shared strings
type SharedStrings []string

// GetSharedStrings - loads and stores the shared string into a string slice
func GetSharedStrings(file *excelize.File) SharedStrings {
	var sharedStrings SharedStrings
	content, ok := file.XLSX["xl/sharedStrings.xml"]
	if ok {
		var sst x.Sst
		_ = xml.Unmarshal(content, &sst)
		if sst.Si != nil && len(sst.Si) > 0 {
			sharedStrings = make([]string, len(sst.Si))
			for i, si := range sst.Si {
				sharedStrings[i] = si.T.Text
			}
		}

	}
	return sharedStrings
}

// Get returns a shared string
func (sharedStrings SharedStrings) Get(idx interface{}) (ss string) {
	var id int
	switch idx.(type) {
	case string:
		id, _ = strconv.Atoi(idx.(string))
	case int:
		id = idx.(int)
	}
	if sharedStrings != nil && id < len(sharedStrings) && id >= 0 {
		ss = sharedStrings[id]
	}
	return
}

// ImportWorksheets - import charts, filters, ...  form workbook file
// also read and match plagiarism key
func (wb *Workbook) ImportWorksheets(fileName string) {
	file, err := excelize.OpenFile(fileName)
	if err != nil {
		log.WithError(err).Errorf("Failed to open file %q", fileName)
		return
	}

	for sheetIdx, sheetName := range file.GetSheetMap() {
		var ws Worksheet
		result := Db.First(&ws, Worksheet{
			Name:       sheetName,
			AnswerID:   wb.AnswerID,
			WorkbookID: wb.ID,
		})
		if result.RecordNotFound() && !DryRun {
			ws = Worksheet{
				Name:             sheetName,
				WorkbookFileName: filepath.Base(fileName),
				AnswerID:         wb.AnswerID,
				WorkbookID:       wb.ID,
			}
			if err := Db.Create(&ws).Error; err != nil {
				log.Fatalf("Failed to create worksheet entry %#v: %s", ws, err.Error())
			}
			if DebugLevel > 1 {
				log.Debugf("Ceated workbook entry %#v", wb)
			}
		}
		ws.Idx = sheetIdx
		Db.Save(ws)
		sharedStrings := GetSharedStrings(file)
		ws.ImportCharts(file)
		ws.ImportWorksheetData(file, sharedStrings)
		wb.MatchPlagiarismKeys(file)
	}
}

// ImportCharts - import charts for the wroksheet
func (ws *Worksheet) ImportCharts(file *excelize.File) {

	name := "xl/worksheets/_rels/sheet" + strconv.Itoa(ws.Idx) + ".xml.rels"
	sheetRels := unmarshalRelationships(file.XLSX[name])
	for _, r := range sheetRels.Relationships {
		if strings.Contains(r.Target, "drawings/drawing") {
			name := "xl/drawings/_rels/" + filepath.Base(r.Target) + ".rels"
			drawing := unmarshalDrawing(file.XLSX["xl/drawings/"+filepath.Base(r.Target)])
			drawingRels := unmarshalRelationships(file.XLSX[name])
			for _, dr := range drawingRels.Relationships {
				if strings.Contains(dr.Target, "charts/chart") {
					chartName := "xl/charts/" + filepath.Base(dr.Target)
					chart := UnmarshalChart(file.XLSX[chartName])
					chartTitle := chart.Title.Value()
					log.Infof("Found %q chart (titled: %q) on the sheet %q", chart.Type(), chartTitle, ws.Name)
					itemCount := chart.ItemCount()
					chartEntry := Chart{
						WorksheetID: ws.ID,
						Title:       chartTitle,
						XLabel:      chart.XLabel(),
						YLabel:      chart.YLabel(),
						FromCol:     drawing.FromCol,
						FromRow:     drawing.FromRow,
						ToCol:       drawing.ToCol,
						ToRow:       drawing.ToRow,
						Type:        chart.Type(),
						Data:        chart.PlotArea.Chart.Data,
						XData:       chart.PlotArea.Chart.XData,
						YData:       chart.PlotArea.Chart.YData,
						ItemCount:   itemCount,
						XMinValue:   chart.XMinValue(),
						XMaxValue:   chart.XMaxValue(),
						YMaxValue:   chart.YMaxValue(),
						YMinValue:   chart.YMinValue(),
					}
					Db.Create(&chartEntry)
					chartID := NewNullInt64(chartEntry.ID)

					Db.Create(&Block{
						WorksheetID: ws.ID,
						Range:       "ChartRange",
						Formula: fmt.Sprintf("((%d,%d),(%d,%d))",
							drawing.FromRow,
							drawing.FromCol,
							drawing.ToRow,
							drawing.ToCol+2),
						ChartID: chartID,
					})
					c := chart.PlotArea.Chart
					var propCount int
					properties := []struct{ Name, Value string }{
						{"ChartType", chart.Type()},
						{"ChartTitle", chartTitle},
						{"X-Axis Title", chart.XLabel()},
						{"Y-Axis Title", chart.YLabel()},
						{"SourceData", c.Data},
						{"X-Axis Data", c.XData},
						{"Y-Axis Data", c.YData},
						{"ItemCount", strconv.Itoa(itemCount)},
						{"X-Axis MinValue", normalizeFloatRepr(chart.XMinValue())},
						{"Y-Axis MinValue", normalizeFloatRepr(chart.YMinValue())},
						{"X-Axis MaxValue", normalizeFloatRepr(chart.XMaxValue())},
						{"Y-Axis MaxValue", normalizeFloatRepr(chart.YMaxValue())},
					}
					for _, p := range properties {
						block := Block{
							WorksheetID: ws.ID,
							Range:       p.Name,
							Formula:     p.Value,
							RelativeFormula: CellAddress(
								drawing.FromRow+propCount, drawing.ToCol+2),
							ChartID: chartID,
						}
						Db.Create(&block)
						Db.Create(&Cell{
							Block:       block,
							WorksheetID: ws.ID,
							Range:       block.Range,
							Formula:     block.Formula,
						})
						propCount++
					}
				}
			}
		}
	}
}

func joinStr(del string, strs ...string) (js string) {
	for _, s := range strs {
		if s == "" {
			continue
		}
		if js != "" {
			js += del
		}
		js += s
	}
	return
}

// ImportWorksheetData imports all filters
func (ws *Worksheet) ImportWorksheetData(file *excelize.File, sharedStrings SharedStrings) {

	name := "xl/worksheets/sheet" + strconv.Itoa(ws.Idx) + ".xml"
	sheet := UnmarshalWorksheet(file.XLSX[name])

	// Sorting:
	for _, ss := range sheet.SortState {
		ds := DataSource{
			WorksheetID: ws.ID,
			Range:       ss.Ref,
		}
		Db.Create(&ds)
		Db.Create(&Block{
			WorksheetID: ws.ID,
			Range:       "SortSource",
			Formula:     ss.Ref,
		})

		var method string
		switch ss.ColumnSort {
		case "1":
			method = "Horizontal"
		default:
			method = "Vertical"
		}
		for _, sc := range ss.SortCondition {
			var st string
			if sc.Descending == "1" {
				st = "descending"
			} else {
				st = "assending"
			}
			sorting := Sorting{
				DataSourceID: ds.ID,
				Method:       method,
				Reference:    sc.Ref,
				Type:         st,
				CustomList:   sc.CustomList,
				IconSet:      sc.IconSet,
				IconID:       sc.IconId,
			}
			Db.Create(&sorting)
			Db.Create(&Block{
				WorksheetID: ws.ID,
				Range:       sc.Ref,
				Formula: joinStr(",",
					sorting.Method,
					sorting.Type,
					sorting.SortBy,
					sorting.CustomList,
					sorting.IconSet,
					sorting.IconID),
				SortingID: NewNullInt64(sorting.ID),
			})
		}
	}

	// Filters:
	for _, af := range sheet.AutoFilter {
		ds := DataSource{
			WorksheetID: ws.ID,
			Range:       af.Ref,
		}
		Db.Create(&ds)
		Db.Create(&Block{
			WorksheetID: ws.ID,
			Range:       "FilterSource",
			Formula:     af.Ref,
		})

		for _, fc := range af.FilterColumn {
			colID, _ := strconv.Atoi(fc.ColId)
			colName := sharedStrings.Get(colID)
			filter := Filter{
				WorksheetID:  ws.ID,
				DataSourceID: ds.ID,
				ColID:        colID,
				ColName:      colName,
			}
			if fc.Filters.Filter != nil {
				for i, ff := range fc.Filters.Filter {
					var f *Filter
					if i == 0 {
						f = &filter
					} else {
						f = &Filter{
							WorksheetID:  ws.ID,
							DataSourceID: ds.ID,
							ColID:        colID,
							ColName:      colName,
						}
					}
					f.Operator = "="
					f.Value = ff.Val
					if i > 0 {
						Db.Create(f)
						Db.Create(&Block{
							WorksheetID:     ws.ID,
							Range:           colName,
							Formula:         f.Operator,
							RelativeFormula: f.Value,
							FilterID:        NewNullInt64(f.ID),
						})
					}
				}

			} else if fc.Top10.Top != "" || fc.Top10.Val != "" || fc.Top10.FilterVal != "" {
				filter.Operator = "top10"
				filter.Value = fc.Top10.Val
			} else if fc.CustomFilters.CustomFilter != nil {
				for i, cf := range fc.CustomFilters.CustomFilter {
					var f *Filter
					if i == 0 {
						f = &filter
					} else {
						f = &Filter{
							WorksheetID:  ws.ID,
							DataSourceID: ds.ID,
							ColID:        colID,
							ColName:      colName,
						}
					}
					switch cf.Operator {
					case "greaterThan":
						f.Operator = ">"
					case "lessThan":
						f.Operator = "<"
					case "notEqual":
						f.Operator = "#"
					case "greaterThanOrEqual":
						f.Operator = ">="
					case "lessThanOrEqual":
						f.Operator = "<="
					default:
						f.Operator = "="
					}
					f.Value = cf.Val
					if i > 0 {
						Db.Create(f)
						Db.Create(&Block{
							WorksheetID:     ws.ID,
							Range:           colName,
							Formula:         f.Operator,
							RelativeFormula: f.Value,
							FilterID:        NewNullInt64(f.ID),
						})
					}
				}
			} else if fc.DynamicFilter.Type != "" || fc.DynamicFilter.Val != "" {
				filter.Operator = fc.DynamicFilter.Type
				filter.Value = fc.DynamicFilter.Val
			}
			Db.Create(&filter)
			Db.Create(&Block{
				WorksheetID:     ws.ID,
				Range:           colName,
				Formula:         filter.Operator,
				RelativeFormula: filter.Value,
				FilterID:        NewNullInt64(filter.ID),
			})
			for _, dgi := range fc.Filters.DateGroupItem {
				item := DateGroupItem{
					FilterID: filter.ID,
					Grouping: dgi.DateTimeGrouping,
					Year:     NewNullInt64(dgi.Year),
					Month:    NewNullInt64(dgi.Month),
					Day:      NewNullInt64(dgi.Day),
					Hour:     NewNullInt64(dgi.Hour),
					Minute:   NewNullInt64(dgi.Minute),
					Second:   NewNullInt64(dgi.Second),
				}
				Db.Create(&item)
				var date string
				switch item.Grouping {
				case "month":
					date = dgi.Year + "/" + dgi.Month
				case "day":
					date = dgi.Year + "/" + dgi.Month + "/" + dgi.Day
				}
				Db.Create(&Block{
					WorksheetID:     ws.ID,
					Range:           colName,
					Formula:         item.Grouping,
					RelativeFormula: date,
					FilterID:        NewNullInt64(filter.ID),
				})
			}
		}
	}

	// Pivot Tables:
	if content, ok := file.XLSX["xl/pivotCache/pivotCacheDefinition"+strconv.Itoa(ws.Idx)+".xml"]; ok {
		pcd := UnmarshalPivotCacheDefinition(content)
		ptd := UnmarshalPivotTableDefinition(
			file.XLSX["xl/pivotTables/pivotTable"+strconv.Itoa(ws.Idx)+".xml"])
		ds := DataSource{
			WorksheetID: ws.ID,
			Range:       pcd.CacheSource.WorksheetSource.Ref,
		}
		Db.Create(&ds)
		Db.Create(&Block{
			WorksheetID: ws.ID,
			Range:       "PivotSource",
			Formula:     ds.Range,
		})
		log.Info(ptd.XMLName)
		if ptd.PivotFields.Count >= "0" {
			var pfIdx, rfIdx, cfIdx, dfIdx = 0, 0, 0, 0
			for _, pf := range ptd.PivotFields.PivotField {
				var label, fieldType, blockCellRange string
				var rec PivotTable

				if pf.DataField == "" {
					switch pf.Axis {
					case "axisPage":
						fieldType, blockCellRange = "Filter", "PageField"
						label = sharedStrings.Get(ptd.PageFields.PageField[pfIdx].Fld)
						pfIdx++
					case "axisRow":
						fieldType, blockCellRange = "Row", "RowField"
						label = sharedStrings.Get(ptd.RowFields.Field[rfIdx].X)
						rfIdx++
					case "axisCol":
						fieldType, blockCellRange = "Column", "ColField"
						label = sharedStrings.Get(ptd.ColFields.Field[cfIdx].X)
						cfIdx++
					default:
						log.Warnf("Unhandled pivot field: %#v", pf)
						continue
					}

					rec = PivotTable{
						DataSourceID: ds.ID,
						Type:         fieldType,
						Label:        label,
					}

				} else {
					// DataField
					var function string
					df := ptd.DataFields.DataField[dfIdx]
					fieldType, blockCellRange = "Value", "DataField"
					label = sharedStrings.Get(df.Fld)
					if df.Subtotal != "" {
						function = df.Subtotal
					} else {
						function = "sum"
					}
					dfIdx++

					rec = PivotTable{
						DataSourceID: ds.ID,
						Type:         "Value",
						Label:        label,
						Function:     function,
						DisplayName:  df.Name,
					}
				}

				Db.Create(&rec)
				Db.Create(&Block{
					WorksheetID: ws.ID,
					Range:       blockCellRange,
					Formula:     label,
					PivotID:     NewNullInt64(rec.ID),
				})
			}
		}
	}

	// Conditional Formatting
	for _, cf := range sheet.ConditionalFormatting {
		ds := DataSource{
			WorksheetID: ws.ID,
			Range:       cf.Sqref,
		}
		Db.Create(&ds)
		Db.Create(&Block{
			WorksheetID:  ws.ID,
			Range:        "CFSource",
			Formula:      ds.Range,
			DataSourceID: NewNullInt64(ds.ID),
		})
		for _, cfr := range cf.CfRule {
			var operator, formula1, formula2, formula3 string
			switch cfr.Type {
			case "iconSet":
				operator = cfr.IconSet.IconSet
			case "aboveAverage":
				if cfr.AboveAverage == "0" {
					operator = "bellow average"
				} else {
					operator = "above average"
				}
			case "top10":
				if cfr.Bottom == "1" {
					operator = "bottom"
				} else {
					operator = "top"
				}
			default:
				operator = cfr.Operator
			}
			switch cfr.Type {
			case "containsText":
				formula1 = cfr.AttrText
			case "timePeriod":
				formula1 = cfr.TimePeriod
			case "top10":
				formula1 = cfr.Rank
			default:
				if len(cfr.Formula) > 0 {
					formula1 = cfr.Formula[0].Text
				}
			}
			if cfr.Type == "top10" {
				if cfr.Percent != "" {
					formula2 = "percent"
				}
			} else if len(cfr.Formula) > 1 {
				formula2 = cfr.Formula[1].Text
			}
			if len(cfr.Formula) > 2 {
				formula3 = cfr.Formula[2].Text
			}

			rec := ConditionalFormatting{
				DataSourceID: ds.ID,
				Type:         cfr.Type,
				Operator:     operator,
				Formula1:     formula1,
				Formula2:     formula2,
				Formula3:     formula3,
			}
			Db.Create(&rec)
			Db.Create(&Block{
				WorksheetID:  ws.ID,
				Range:        rec.Type,
				Formula:      joinStr(",", operator, formula1, formula2, formula3),
				DataSourceID: NewNullInt64(ds.ID),
			})

		}

	}
}

// normalizeFloatRepr - if val is float representation round it to 3 digits after the '.'
func normalizeFloatRepr(val string) string {
	if strings.Contains(val, ".") {
		floatVal, err := strconv.ParseFloat(val, 64)
		if err == nil {
			return strings.TrimRight(fmt.Sprintf("%.3f", floatVal), "0")
		}
	}
	return val
}

// Worksheet - Excel workbook worksheet
type Worksheet struct {
	ID               int
	Name             string
	WorkbookFileName string
	Blocks           []Block       `gorm:"ForeignKey:WorksheetID"`
	Answer           Answer        `gorm:"ForeignKey:AnswerID"`
	AnswerID         sql.NullInt64 `gorm:"column:StudentAnswerID;index;type:int"`
	Workbook         Workbook      `gorm:"ForeignKey:WorkbookId"`
	WorkbookID       int           `gorm:"index"`
	IsReference      bool
	OrderNum         int
	Idx              int
	IsPlagiarised    bool   // sql.NullBool
	Cells            []Cell `gorm:"ForeignKey:WorksheetID"`
}

// TableName overrides default table name for the model
func (Worksheet) TableName() string {
	return "WorkSheets"
}

// BlockCommentRow - a block comment row
type BlockCommentRow struct {
	Range                  string
	CommentText            string
	TRow, LCol, BRow, RCol int
}

// GetBlockComments retrieves all block comments in a form of a map
func (ws *Worksheet) GetBlockComments() (res map[int][]BlockCommentRow, err error) {
	// var (
	// 	blockRange             string
	// 	commentText            string
	// 	tRow, lCol, bRow, rCol int
	// )

	rows, err := Db.Raw(`SELECT 
	  b.BlockCellRange,
	  c.CommentText,
	  b.t_row, b.l_col, b.b_row, b.r_col
    FROM ExcelBlocks AS b
      LEFT JOIN BlockCommentMapping AS bc ON bc.ExcelBlockID = b.ExcelBlockID
      LEFT JOIN Comments AS c ON c.CommentID = bc.ExcelCommentID
    WHERE b.worksheet_id = ?
	ORDER BY b.l_col, b.t_row`, ws.ID).Rows()
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	res = make(map[int][]BlockCommentRow)
	col := -1
	for rows.Next() {
		r := BlockCommentRow{}
		rows.Scan(&r.Range, &r.CommentText, &r.TRow, &r.LCol, &r.BRow, &r.RCol)
		if r.LCol != col {
			col = r.LCol
			res[col] = make([]BlockCommentRow, 1)
			res[col][0] = r
		} else {
			res[col] = append(res[col], r)
		}
	}

	return
}

// CellCommentRow - a cell comment row
type CellCommentRow struct {
	Range       string
	CommentText string
	Row, Col    int
}

// GetCellComments retrieves all block comments in a form of a map
func (ws *Worksheet) GetCellComments() (res []CellCommentRow, err error) {
	rows, err := Db.Raw(`SELECT cell.cell_range, c.CommentText, cell.row, cell.col
    FROM Cells AS cell JOIN Comments AS c ON c.CommentID = cell.CommentID
    WHERE  cell.worksheet_id = ?
	ORDER BY cell.col, cell."row"`, ws.ID).Rows()
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	res = make([]CellCommentRow, 0)
	for rows.Next() {
		r := CellCommentRow{}
		rows.Scan(&r.Range, &r.CommentText, &r.Row, &r.Col)
		res = append(res, r)
	}
	return

}

// Block - Excel block
type Block struct {
	ID              int `gorm:"column:ExcelBlockID;primary_key:true;AUTO_INCREMENT"`
	Color           string
	Range           string                       `gorm:"column:BlockCellRange"`
	Formula         string                       `gorm:"column:BlockFormula"` // first block cell formula
	RelativeFormula string                       // first block cell relative formula formula
	Cells           []Cell                       `gorm:"ForeignKey:BlockID"`
	Worksheet       Worksheet                    `gorm:"ForeignKey:WorksheetID"`
	WorksheetID     int                          `gorm:"index"`
	CommentMappings []BlockCommentMapping        `gorm:"ForeignKey:ExcelBlockID"`
	Chart           Chart                        `gorm:"ForeignKey:ChartId"`
	ChartID         sql.NullInt64                `grom:"type:int;index"`
	IsReference     bool                         // the block is used for referencing the expected bloks
	TRow            int                          `gorm:"index"` // Top row
	LCol            int                          `gorm:"index"` // Left column
	BRow            int                          `gorm:"index"` // Bottom row
	RCol            int                          `gorm:"index"` // Right column
	DataSourceID    sql.NullInt64                `gorm:"column:source_id;type:int"`
	FilterID        sql.NullInt64                `gorm:"type:int"`
	SortingID       sql.NullInt64                `gorm:"column:sort_id;type:int"`
	PivotID         sql.NullInt64                `gorm:"type:int"`
	i               struct{ sr, sc, er, ec int } `gorm:"-"` // "Inner" block - the block containing values
	isEmpty         bool                         `gorm:"-"` // All block cells are empty
}

// TableName overrides default table name for the model
func (b Block) TableName() string {
	return "ExcelBlocks"
}

func (b Block) String() string {
	var output string
	if b.IsReference {
		output = "Reference"
	} else {
		output = "Block"
	}
	return fmt.Sprintf(
		"%s {ID: %d, Range: %q [%d, %d, %d, %d], Color: %q, Formula: %q, Relative Formula: %q, WorksheetID: %d}",
		output, b.ID, b.Range, b.TRow, b.LCol, b.BRow, b.RCol, b.Color, b.Formula, b.RelativeFormula, b.WorksheetID)
}

func (b *Block) save() {
	if !DryRun {
		if !b.IsReference {
			for i := b.LCol; i <= b.RCol; i++ {
				for j := b.TRow; j <= b.BRow; j++ {
					if b.isEmpty || i < b.i.sc || i > b.i.ec || j < b.i.sr || j > b.i.er {
						address := CellAddress(j, i)
						address += ":" + address
						empty := Block{
							WorksheetID: b.WorksheetID,
							Range:       address,
							Color:       b.Color,
							TRow:        j,
							LCol:        i,
							BRow:        j,
							RCol:        i,
						}
						if VerboseLevel > 0 {
							log.Infof("*** Created an empty cell/block: %#v", empty)
						}
						r := Db.Where("worksheet_id = ? AND BlockCellRange = ?", b.WorksheetID, address).
							First(&empty)
						if r.RecordNotFound() {
							Db.Create(&empty)
						}
					}
				}
			}
		}
		if b.isEmpty && !b.IsReference {
			Db.Delete(&Cell{}, "block_id = ?", b.ID)
			Db.Delete(b)
		} else {
			if b.IsReference {
				b.Range = b.Address()
			} else {
				b.Range = b.InnerAddress()
				if b.LCol <= b.i.sc && b.i.ec <= b.RCol && b.TRow <= b.i.sr && b.i.er <= b.BRow {
					b.LCol, b.TRow, b.RCol, b.BRow = b.i.sc, b.i.sr, b.i.ec, b.i.er
				}
			}
			Db.Save(b)
		}
	}
}

// Address - the block range
func (b *Block) Address() string {
	return CellAddress(b.TRow, b.LCol) + ":" + CellAddress(b.BRow, b.RCol)
}

// InnerAddress - the block "inner" range excluding empty cells
func (b *Block) InnerAddress() string {
	return CellAddress(b.i.sr, b.i.sc) + ":" + CellAddress(b.i.er, b.i.ec)
}

//  getCellComment returns cell comment text value
func getCellComment(file *xlsx.File, cellID string) string {
	if file.Comments != nil {
		for _, c := range file.Comments {
			if cellID == c.Ref {
				return c.Text
			}
		}
	}
	return ""
}

// cellValue returns cell value
func cellValue(cell *xlsx.Cell) (value string) {
	var err error
	if cell.Type() == 2 {
		if value, err = cell.FormattedValue(); err != nil {
			log.WithError(err).Error("Failed to read cell value: ", *cell)
			value = cell.Value
		}
	} else {
		value = cell.Value
	}
	return
}

// fildWhole finds whole range of the specified color
// and the same "relative" formula starting with the set top-left cell.
func (b *Block) findWhole(sheet *xlsx.Sheet, color string) {

	b.BRow, b.RCol = b.TRow, b.LCol
	for i, row := range sheet.Rows {

		// skip all rows until the first block row
		if i < b.TRow {
			continue
		}

		log.Debugf("Total cells: %d at %d", len(row.Cells), i)
		// Range is discontinued or of a differnt color
		if len(row.Cells) <= b.RCol ||
			row.Cells[b.RCol].GetStyle().Fill.FgColor != color ||
			RelativeFormula(i, b.RCol, row.Cells[b.RCol].Formula()) != b.RelativeFormula {
			log.Debugf("Reached the edge row of the block at row %d", i)
			b.BRow = i - 1
			break
		} else {
			b.BRow = i
		}

		for j, cell := range row.Cells {
			// skip columns until the start:
			if j < b.LCol {
				continue
			}

			// Reached the top-right corner:
			if fgColor := cell.GetStyle().Fill.FgColor; fgColor == color {
				if !b.IsReference {
					relFormula := RelativeFormula(i, j, cell.Formula())
					if relFormula == b.RelativeFormula {
						cellID := CellAddress(i, j)
						if value := cellValue(cell); value != "" {
							c := Cell{
								BlockID:     b.ID,
								WorksheetID: b.WorksheetID,
								Formula:     cell.Formula(),
								Value:       value,
								Range:       cellID,
							}
							if DebugLevel > 1 {
								log.Debugf("Inserting %#v", c)
							}

							if err := Db.Create(&c).Error; err != nil {
								log.WithError(err).Error("Failed to create a cell: ", c)
							}
						}
					}
				}
				b.RCol = j
			} else {
				log.Debugf("Reached the edge column  of the block at column %d", j)
				if j > b.RCol {
					b.RCol = j - 1
				}
				break
			}
		}
	}

	if b.IsReference {
		return
	}
	// Find the part containing values
	b.findInner(sheet)
}

// fildWholeWithin finds whole range with the same "relative" formula
// withing the specific reference block ignoring the filling color.
func (b *Block) findWholeWithin(sheet *xlsx.Sheet, rb Block) {
	b.BRow, b.RCol = b.TRow, b.LCol
	for r := b.TRow; r <= rb.BRow; r++ {

		row := sheet.Row(r)
		// Range is discontinued or of a differnt relative formula
		if len(row.Cells) <= b.RCol ||
			RelativeFormula(r, b.RCol, row.Cells[b.RCol].Formula()) != b.RelativeFormula {
			log.Debugf("Reached the edge row of the block at row %d", r)
			b.BRow = r - 1
			break
		} else {
			b.BRow = r
		}

		for c := b.RCol + 1; c <= rb.RCol; c++ {
			// skip columns until the start:
			cell := sheet.Cell(r, c)

			// Reached the bottom-right corner:
			if relFormula := RelativeFormula(r, c, cell.Formula()); relFormula == b.RelativeFormula {
				b.RCol = c
			} else {
				log.Debugf("Reached the edge column  of the block at column %d", c)
				if c > b.RCol {
					b.RCol = c - 1
				}
				break
			}
		}
	}
	b.Range = b.Address()
	for r := b.TRow; r <= b.BRow; r++ {
		for c := b.LCol; c <= b.RCol; c++ {
			cell := sheet.Cell(r, c)
			if value := cellValue(cell); value != "" {
				err := Db.Create(&Cell{
					BlockID:     b.ID,
					WorksheetID: b.WorksheetID,
					Formula:     cell.Formula(),
					Value:       value,
					Range:       CellAddress(r, c),
				}).Error
				if err != nil {
					log.WithError(err).Error("Failed to create a cell.")
				}
			}
		}
	}
}

// findInner finds the part containing values
func (b *Block) findInner(sheet *xlsx.Sheet) {
	sr, sc, er, ec := b.TRow, b.LCol, b.BRow, b.RCol
	for sc <= ec {
		for r := sr; r <= er; r++ {
			if value := cellValue(sheet.Cell(r, sc)); value != "" {
				goto RIGHT_COL
			}
		}
		sc++
	}
	if sc > ec {
		b.isEmpty = true
		return
	}

RIGHT_COL:
	for ec >= sc {
		for r := sr; r <= er; r++ {
			if value := cellValue(sheet.Cell(r, ec)); value != "" {
				goto TOP_ROW
			}
		}
		ec--
	}
TOP_ROW:
	for sr <= er {
		for c := sc; c <= ec; c++ {
			if value := cellValue(sheet.Cell(sr, c)); value != "" {
				goto BOTTOM_ROW
			}
		}
		sr++
	}
BOTTOM_ROW:
	for er >= sr {
		for c := sc; c <= ec; c++ {
			if value := cellValue(sheet.Cell(er, c)); value != "" {
				goto FOUND
			}
		}
		er--
	}
FOUND:
	b.i.sr, b.i.sc, b.i.er, b.i.ec = sr, sc, er, ec
}

// IsInside tests if the cell with given coordinates is inside the coordinates
func (b *Block) IsInside(r, c int) bool {
	return (b.TRow <= r &&
		r <= b.BRow &&
		b.LCol <= c &&
		c <= b.RCol)
}

// Cell - a sigle cell of the block
type Cell struct {
	ID          int
	Block       Block `gorm:"ForeignKey:BlockID"`
	BlockID     int   `gorm:"index"`
	Worksheet   Worksheet
	WorksheetID int    `gorm:"index"`
	Range       string `gorm:"column:cell_range"`
	Formula     string
	Value       string `gorm:"size:2000"`
	Comment     Comment
	CommentID   sql.NullInt64 `gorm:"column:CommentID;type:int"`
	Row         int           `gorm:"index"`
	Col         int           `gorm:"index"`
}

// TableName overrides default table name for the model
func (Cell) TableName() string {
	return "Cells"
}

// OpenDb opens DB connection based on given URL
func OpenDb(url string) (db *gorm.DB, err error) {

	parts := strings.Split(url, "://")
	if len(parts) < 2 {
		log.Warnf("Driver name not given in %q, assuming 'mysql'.", url)
		parts = []string{"mysql", parts[0]}
	}

	switch parts[0] {
	case "sqlite", "sqlite3":
		log.Debugf("Connecting to Sqlite3 DB: %q.", parts[1])
		parts[0] = "sqlite3"
	case "mysql":
		log.Debugf("Connecting to MySQL DB: %q.", parts[1])
	default:
		log.Fatalf("Unsupported driver: %q. It should be either 'mysql' or 'sqlite'.", parts[0])
	}
	db, err = gorm.Open(parts[0], parts[1])
	if parts[0] == "mysql" {
		db.Set("gorm:table_options", "collation_connection=utf8_bin")
		if err := db.Exec("SET @@sql_mode='ANSI'").Error; err != nil {
			log.Error(err)
		}
	}
	if err != nil {
		log.Error(err)
		log.Fatalf("failed to connect database %q", url)
	}
	Db = db
	if DebugLevel > 1 {
		db.LogMode(true)
	}
	SetDb()
	return
}

// MySQLQuestion - questions
type MySQLQuestion struct {
	ID                 int                 `gorm:"column:QuestionID;primary_key:true;AUTO_INCREMENT"`
	QuestionType       QuestionType        `gorm:"column:QuestionType;type:ENUM('ShortAnswer','MCQ','FileUpload')"`
	QuestionSequence   int                 `gorm:"column:QuestionSequence;not null"`
	QuestionText       string              `gorm:"column:QuestionText;type:text;not null"`
	AnswerExplanation  sql.NullString      `gorm:"column:AnswerExplanation;type:text"`
	MaxScore           float32             `gorm:"column:MaxScore;type:float;not null"`
	SourceID           sql.NullInt64       `gorm:"column:FileID;type:int"`
	AuthorUserID       int                 `gorm:"column:AuthorUserID;not null"`
	WasCompared        bool                `gorm:"default:0"`
	IsProcessed        bool                `gorm:"column:IsProcessed;default:0"`
	Source             Source              `gorm:"ForeignKey:FileID"`
	Answers            []Answer            `gorm:"ForeignKey:QuestionID"`
	QuestionExcelDatas []QuestionExcelData `gorm:"ForeignKey:QuestionID"`
	ReferenceID        sql.NullInt64       `gorm:"index;type:int"`
}

// TableName overrides default table name for the model
func (MySQLQuestion) TableName() string {
	return "Questions"
}

// Comment - added comments  with marks
type Comment struct {
	ID              int                   `gorm:"column:CommentID;primary_key:true;AUTO_INCREMENT"`
	Text            string                `gorm:"column:CommentText"`
	Marks           sql.NullFloat64       `gorm:"column:Marks;type:float"`
	CommentMappings []BlockCommentMapping `gorm:"ForeignKey:ExcelCommentID"`
	AnswerComments  []AnswerComment       `gorm:"ForeignKey:CommentID"`
}

// TableName overrides default table name for the model
func (Comment) TableName() string {
	return "Comments"
}

// BlockCommentMapping - block-comment mapping
type BlockCommentMapping struct {
	Block     Block
	BlockID   int `gorm:"column:ExcelBlockID"`
	Comment   Comment
	CommentID int `gorm:"column:ExcelCommentID"`
}

// TableName overrides default table name for the model
func (BlockCommentMapping) TableName() string {
	return "BlockCommentMapping"
}

// AnswerComment - answer-comment mapping:
type AnswerComment struct {
	Answer    Answer
	AnswerID  int `gorm:"column:StudentAnswerID;index"`
	Comment   Comment
	CommentID int `gorm:"column:CommentID"`
}

// TableName overrides default table name for the model
func (AnswerComment) TableName() string {
	return "StudentAnswerCommentMapping"
}

// QuestionAssignment - question-assignment mapping
type QuestionAssignment struct {
	Assignment   Assignment
	AssignmentID int `gorm:"column:AssignmentID"`
	Question     Question
	QuestionID   int `gorm:"column:QuestionID"`
}

// TableName overrides default table name for the model
func (QuestionAssignment) TableName() string {
	return "QuestionAssignmentMapping"
}

// SetDb initializes DB
func SetDb() {
	// Migrate the schema
	isMySQL := strings.HasPrefix(Db.Dialect().GetName(), "mysql")
	log.Debug("Add to automigrate...")

	Db.AutoMigrate(&Source{})
	if isMySQL {
		// Modify struct tag for MySQL
		Db.AutoMigrate(&MySQLQuestion{})
	} else {
		Db.AutoMigrate(&Question{})
	}
	Db.AutoMigrate(&User{})
	Db.AutoMigrate(&QuestionExcelData{})
	Db.AutoMigrate(&Answer{})
	Db.AutoMigrate(&Workbook{})
	Db.AutoMigrate(&Worksheet{})
	Db.AutoMigrate(&DataSource{})
	Db.AutoMigrate(&Filter{})
	Db.AutoMigrate(&DateGroupItem{})
	Db.AutoMigrate(&Sorting{})
	Db.AutoMigrate(&PivotTable{})
	Db.AutoMigrate(&ConditionalFormatting{})
	Db.AutoMigrate(&Chart{})
	Db.AutoMigrate(&Block{})
	Db.AutoMigrate(&Cell{})
	Db.AutoMigrate(&Comment{})
	Db.AutoMigrate(&BlockCommentMapping{})
	Db.AutoMigrate(&Assignment{})
	Db.AutoMigrate(&QuestionAssignment{})
	Db.AutoMigrate(&AnswerComment{})
	Db.AutoMigrate(&XLQTransformation{})
	if isMySQL {
		// Add some foreing key constraints to MySQL DB:
		log.Debug("Adding a constraint to Wroksheets -> Answers...")
		Db.Model(&Worksheet{}).AddForeignKey("StudentAnswerID", "StudentAnswers(StudentAnswerID)", "CASCADE", "CASCADE")
		log.Debug("Adding a constraint to Cells...")
		Db.Model(&Cell{}).AddForeignKey("block_id", "ExcelBlocks(ExcelBlockID)", "CASCADE", "CASCADE")
		log.Debug("Adding a constraint to Blocks...")
		Db.Model(&Block{}).AddForeignKey("worksheet_id", "WorkSheets(id)", "CASCADE", "CASCADE")
		log.Debug("Adding a constraint to Worksheets -> Workbooks...")
		Db.Model(&Block{}).AddForeignKey("filter_id", "Filters(id)", "CASCADE", "CASCADE")
		Db.Model(&Block{}).AddForeignKey("sort_id", "Sortings(id)", "CASCADE", "CASCADE")
		Db.Model(&Block{}).AddForeignKey("pivot_id", "PivotTables(id)", "CASCADE", "CASCADE")
		Db.Model(&Block{}).AddForeignKey("source_id", "DataSources(id)", "CASCADE", "CASCADE")
		Db.Model(&Worksheet{}).AddForeignKey("workbook_id", "WorkBooks(id)", "CASCADE", "CASCADE")
		log.Debug("Adding a constraint to Questions...")
		Db.Model(&Question{}).AddForeignKey("FileID", "FileSources(FileID)", "CASCADE", "CASCADE")
		log.Debug("Adding a constraint to QuestionExcelData...")
		Db.Model(&QuestionExcelData{}).AddForeignKey("QuestionID", "Questions(QuestionID)", "CASCADE", "CASCADE")
		Db.Model(&DataSource{}).AddForeignKey("worksheet_id", "WorkSheets(id)", "CASCADE", "CASCADE")
		Db.Model(&Filter{}).AddForeignKey("worksheet_id", "WorkSheets(id)", "CASCADE", "CASCADE")
		Db.Model(&Filter{}).AddForeignKey("DataSourceID", "DataSources(id)", "CASCADE", "CASCADE")
		Db.Model(&DateGroupItem{}).AddForeignKey("filter_id", "Filters(id)", "CASCADE", "CASCADE")
		Db.Model(&PivotTable{}).AddForeignKey("DataSourceID", "DataSources(id)", "CASCADE", "CASCADE")
		Db.Model(&XLQTransformation{}).AddForeignKey("UserID", "Users(UserID)", "CASCADE", "CASCADE")
		Db.Model(&XLQTransformation{}).AddForeignKey("QuestionID", "Questions(QuestionID)", "CASCADE", "CASCADE")
		Db.Model(&XLQTransformation{}).AddForeignKey("FileID", "FileSources(FileID)", "CASCADE", "CASCADE")
	}
}

// QuestionsToProcess returns list of questions that need to be processed
func QuestionsToProcess() ([]Question, error) {

	var questions []Question
	result := (Db.
		Joins("JOIN FileSources ON FileSources.FileID = Questions.FileID").
		Where("IsProcessed = ?", 0).
		Where("FileSources.FileName LIKE ?", "%.xlsx").
		Find(&questions))
	return questions, result.Error
}

// RowsToProcessResult stores query resut
type RowsToProcessResult struct {
	ID              int           `gorm:"column:FileID"`
	S3BucketName    string        `gorm:"column:S3BucketName"`
	S3Key           string        `gorm:"column:S3Key"`
	FileName        string        `gorm:"column:FileName"`
	StudentAnswerID int           `gorm:"column:StudentAnswerID"`
	QuestionID      sql.NullInt64 `gorm:"column:QuestionID;type:int"`
}

// RowsToProcess returns answer file sources
func RowsToProcess() ([]RowsToProcessResult, error) {

	// TODO: select file links from StudentAnswers and download them form S3 buckets..."
	rows, err := Db.Table("FileSources").
		Select("FileSources.FileID, S3BucketName, S3Key, FileName, StudentAnswerID, QuestionID").
		Joins("JOIN StudentAnswers ON StudentAnswers.FileID = FileSources.FileID").
		Where("FileName IS NOT NULL").
		Where("FileName != ?", "").
		Where("FileName LIKE ?", "%.xlsx").
		Where("StudentAnswers.was_xl_processed = ?", 0).Rows()
	defer rows.Close()

	if err != nil {
		return nil, err
	}

	var results []RowsToProcessResult
	for rows.Next() {
		var r RowsToProcessResult
		Db.ScanRows(rows, &r)
		results = append(results, r)
	}

	return results, nil
}

// RowsToComment returns slice with all recored of source files
// and AswerIDs that need to be commeted
func RowsToComment(assignmentID int) ([]RowsToProcessResult, error) {
	q := Db.Table("FileSources").
		Select("DISTINCT FileSources.FileID, S3BucketName, S3Key, FileName, StudentAnswerID").
		Joins("JOIN StudentAnswers ON StudentAnswers.FileID = FileSources.FileID").
		Joins("JOIN Questions ON Questions.QuestionID = StudentAnswers.QuestionID").
		Joins("JOIN QuestionAssignmentMapping ON QuestionAssignmentMapping.QuestionID = Questions.QuestionID")
	if assignmentID > 0 {
		q = q.Joins("JOIN CourseAssignments ON CourseAssignments.AssignmentID = QuestionAssignmentMapping.AssignmentID")
	}
	q = q.Where("was_comment_processed = ?", 0).
		Where("FileName IS NOT NULL").
		Where("FileName != ?", "").
		Where("FileName LIKE ?", "%.xlsx").
		Where(`
			EXISTS(
				SELECT NULL
				FROM WorkSheets AS ws JOIN ExcelBlocks AS b ON b.worksheet_id = ws.id
				JOIN BlockCommentMapping AS bcm ON bcm.ExcelBlockID = b.ExcelBlockID
				WHERE ws.StudentAnswerID = StudentAnswers.StudentAnswerID
			)
		`)
		// Where("CourseAssignments.State = ?", "GRADED").
	if assignmentID > 0 {
		q = q.Where("QuestionAssignmentMapping.AssignmentID = ?", assignmentID)
	}
	rows, err := q.Rows()
	defer rows.Close()

	if err != nil {
		return nil, err
	}

	var results []RowsToProcessResult
	for rows.Next() {
		var r RowsToProcessResult
		Db.ScanRows(rows, &r)
		results = append(results, r)
	}

	return results, nil
}

type blockList []Block

// wasFound tests if the range containing the cell
// coordinates hhas been already found.
func (bl *blockList) wasFound(r, c int) bool {
	for _, b := range *bl {
		if b.IsInside(r, c) {
			return true
		}
	}
	return false
}

// createEmptyCellBlock - create a block consisting of a single cell
func (ws *Worksheet) createEmptyCellBlock(r, c int) (err error) {
	address := CellAddress(r, c)
	address += ":" + address
	return Db.Create(&Block{
		WorksheetID: ws.ID,
		Range:       address,
		TRow:        r,
		LCol:        c,
		BRow:        r,
		RCol:        c,
	}).Error
}

// FindBlocksInside - find answer blocks within the reference block (rb) and store them
func (ws *Worksheet) FindBlocksInside(sheet *xlsx.Sheet, rb Block) (err error) {
	var (
		b      Block
		cell   *xlsx.Cell
		blocks = blockList{}
	)

	for r := rb.TRow; r <= rb.BRow; r++ {
		for c := rb.LCol; c <= rb.RCol; c++ {

			if blocks.wasFound(r, c) {
				continue
			}

			cell = sheet.Cell(r, c)
			if formula, value := cell.Formula(), cellValue(cell); value != "" && formula != "" {
				b = Block{
					WorksheetID:     ws.ID,
					TRow:            r,
					LCol:            c,
					Formula:         formula,
					RelativeFormula: RelativeFormula(r, c, formula),
				}
				if !DryRun {
					Db.Create(&b)
				}

				if DebugLevel > 1 {
					log.Debugf("Created %#v", b)
				}

				b.findWholeWithin(sheet, rb)
				b.Range = b.Address()
				Db.Save(b)
				blocks = append(blocks, b)
			}
		}
	}
	for r := rb.TRow; r <= rb.BRow; r++ {
		for c := rb.LCol; c <= rb.RCol; c++ {
			if !blocks.wasFound(r, c) {
				ws.createEmptyCellBlock(r, c)
			}
		}
	}
	return
}

// ExtractBlocksFromFile extracts blocks from the given file and stores in the DB
func ExtractBlocksFromFile(fileName, color string, force, verbose bool, answerIDs ...int) (wb Workbook, err error) {
	var (
		answerID int
		answer   Answer
	)
	if len(answerIDs) > 0 {
		answerID = answerIDs[0]
	} else {
		err = errors.New("Missing AnswerID")
		return
	}
	res := Db.First(&answer, answerID)
	if res.RecordNotFound() {
		err = fmt.Errorf("Answer (ID: %d) not found", answerID)
		return
	}

	file, err := xlsx.OpenFile(fileName)
	if err != nil {
		log.WithError(err).Errorf("Failed to open the file %q (AnswerID: %d), file might be corrupt.",
			fileName, answerID)

		if err := Db.Model(&answer).Updates(map[string]interface{}{
			"was_xl_processed": 0,
			"FileID":           gorm.Expr("NULL"),
		}).Error; err != nil {
			log.WithError(err).Errorln("Failed to update the answer entry.")
		}
		return
	}

	result := Db.First(&wb, Workbook{FileName: fileName, AnswerID: NewNullInt64(answerID)})
	if !result.RecordNotFound() {
		if !force {
			log.Errorf("File %q was already processed.", fileName)
			return
		}
		log.Warnf("File %q was already processed.", fileName)
		if !DryRun {
			wb.Reset()
		}
	} else if result.RecordNotFound() {
		if !DryRun {

			wb = Workbook{FileName: fileName, AnswerID: NewNullInt64(answerID)}

			if err = Db.Create(&wb).Error; err != nil {
				log.WithError(err).Errorf("Failed to create workbook entry %#v", wb)
				return
			}
			if DebugLevel > 1 {
				log.Debugf("Ceated workbook entry %#v", wb)
			}
		}
	} else if err = result.Error; err != nil {
		return
	}

	if verbose {
		log.Infof("*** Processing workbook: %s", fileName)
	}
	var q Question
	err = Db.
		Joins("JOIN StudentAnswers ON StudentAnswers.QuestionID = Questions.QuestionID").
		Where("StudentAnswers.StudentAnswerID = ?", answerID).
		Order("Questions.reference_id DESC").
		First(&q).Error
	if err != nil {
		log.WithError(err).Errorln("Failed to retrieve the question entry for the answer ID: ", answerID)
		return
	}
	if verbose {
		log.Infof("*** Processing the answer ID: %d for the queestion %s", answerID, q)
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
		var references []Block
		if !DryRun {
			err = Db.FirstOrCreate(&ws, Worksheet{
				Name:             sheet.Name,
				WorkbookID:       wb.ID,
				WorkbookFileName: wb.FileName,
				AnswerID:         NewNullInt64(answerID),
				OrderNum:         orderNum,
			}).Error
			if err != nil {
				log.WithError(err).Errorln("*** Failed to create worksheet entry: ", sheet.Name)
			}
		}

		// Attempt to use reference blocks for the answer if it's given:
		if q.ReferenceID.Valid {
			err = Db.
				Joins("JOIN WorkSheets ON WorkSheets.id = ExcelBlocks.worksheet_id").
				Where("ExcelBlocks.is_reference").
				Where("WorkSheets.workbook_id = ?", q.ReferenceID).
				Where("WorkSheets.order_num = ?", orderNum).
				Find(&references).Error
			if err != nil {
				log.WithError(err).Errorln("Failed to fetch reference blocks for the question:", q)
				continue
			}
			// Attempt to use reference blocks for the answer:
			for _, rb := range references {
				if verbose {
					log.Info("Attempt to use a reference block: ", rb)
				}
				err = ws.FindBlocksInside(sheet, rb)
				if err != nil {
					log.WithError(err).Errorln("Failed to find the blocks using the reference block: ", rb)
				}
			}
		} else {
			// Keep looking for color-coded blocks
			blocks := blockList{}
			sheetFillColors := []string{}

			for i, row := range sheet.Rows {
				for j, cell := range row.Cells {

					if blocks.wasFound(i, j) {
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
				log.Warningf("No block found ot the worksheet %q of the workbook %q with color %q", sheet.Name, fileName, color)
				if len(sheetFillColors) > 0 {
					log.Infof("Following colors were found in the worksheet you could use: %v", sheetFillColors)
				}
			}
		}
	}

	wb.ImportWorksheets(fileName)
	if _, err := Db.DB().
		Exec(`
			UPDATE StudentAnswers
			SET was_xl_processed = 1
			WHERE StudentAnswerID = ?`, answerID); err != nil {
		log.WithError(err).Errorf("Failed to update the answer entry.")
	}

	// Comments that should be linked with the file:
	const commentsToMapSQL = `
	SELECT
		nsa.StudentAnswerID, nb.ExcelBlockID, MAX(bc.ExcelCommentID) AS ExcelCommentID
	-- Newly added/processed answers
	FROM StudentAnswers AS nsa
	JOIN WorkBooks AS nwb ON nwb.StudentAnswerID = nsa.StudentAnswerID
	JOIN WorkSheets AS nws ON nws.workbook_id = nwb.id
	JOIN ExcelBlocks AS nb
	ON nb.worksheet_id = nws.id
	-- Existing asnwers with mapped comments
	JOIN StudentAnswers AS sa ON sa.QuestionID = nsa.QuestionID
		AND sa.was_xl_processed != 0
		AND sa.StudentAnswerID != nsa.StudentAnswerID
	JOIN WorkBooks AS wb ON wb.StudentAnswerID = sa.StudentAnswerID
	JOIN WorkSheets AS ws ON ws.workbook_id = wb.id
	JOIN ExcelBlocks AS b ON b.worksheet_id = ws.id
		AND b.BlockCellRange = nb.BlockCellRange
		AND b.BlockFormula = nb.BlockFormula
	JOIN BlockCommentMapping AS bc
	ON bc.ExcelBlockID = b.ExcelBlockID
	-- Make sure the newly added block isn't mapped already
	LEFT OUTER JOIN BlockCommentMapping AS nbc
	ON nbc.ExcelBlockID = nb.ExcelBlockID
	WHERE nwb.StudentAnswerID = ?
	AND nbc.ExcelCommentID IS NULL
	GROUP BY nsa.StudentAnswerID, nb.ExcelBlockID`

	// Insert block -> comment mapping:
	sql := `
		INSERT INTO BlockCommentMapping(ExcelBlockID, ExcelCommentID)
		SELECT ExcelBlockID, ExcelCommentID FROM (` +
		commentsToMapSQL + ") AS c"
	_, err = Db.DB().Exec(sql, answerID)
	if err != nil {
		log.Info("SQL: ", sql)
		log.WithError(err).Errorln("Failed to insert block -> comment mapping.")
	}

	// Insert block -> comment mapping:
	sql = `
		INSERT INTO StudentAnswerCommentMapping(StudentAnswerID, CommentID)
		SELECT DISTINCT StudentAnswerID, ExcelCommentID FROM (` +
		commentsToMapSQL + ") AS c"
	_, err = Db.DB().Exec(sql, answerID)
	if err != nil {
		log.Info("SQL: ", sql)
		log.WithError(err).Errorln("Failed to insert block -> comment mapping.")
	}

	err = Db.Model(&answer).UpdateColumn("was_xl_processed", 1).Error
	if err != nil {
		log.WithError(err).Errorln("Failed to update the answer entry.")
	}
	return
}

// Chart - Excel chart
type Chart struct {
	ID                                         int
	Worksheet                                  Worksheet
	WorksheetID                                int `gorm:"index"`
	Title, XLabel, YLabel                      string
	FromCol, FromRow, ToCol, ToRow, ItemCount  int
	Data, XData, YData, Type                   string
	XMinValue, XMaxValue, YMaxValue, YMinValue string
}

// DataSource - autofilter
type DataSource struct {
	ID          int
	WorksheetID int
	Worksheet   Workbook
	Range       string `gorm:"column:Sourcerange;type:varchar(255)"`
}

// TableName overrides default table name for the model
func (DataSource) TableName() string {
	return "DataSources"
}

// Filter - filters
type Filter struct {
	ID           int
	WorksheetID  int
	Worksheet    *Workbook
	DataSourceID int `gorm:"column:DataSourceID"`
	DataSource   *DataSource
	ColID        int    `gorm:"column:ColID"`
	ColName      string `gorm:"column:ColName;type:varchar(255)"`
	Operator     string `gorm:"column:Operator;type:varchar(50)"`
	Value        string `gorm:"column:Value;type:varchar(255)"`
}

// TableName overrides default table name for the model
func (Filter) TableName() string {
	return "Filters"
}

// DateGroupItem - data group
type DateGroupItem struct {
	ID                   int
	Grouping             string        `gorm:"column:datetTimeGroupingType;type:varchar(10)"`
	Year, Month, Day     sql.NullInt64 `gorm:"type:int"`
	Hour, Minute, Second sql.NullInt64 `gorm:"type:int"`
	FilterID             int
	Filter               *Filter
}

// TableName overrides default table name for the model
func (DateGroupItem) TableName() string {
	return "DateGroupItems"
}

// Sorting - column sorting
type Sorting struct {
	ID           int
	DataSourceID int    `gorm:"column:DataSourceID"`
	Method       string `gorm:"column:SortMethod;type:varchar(10);not null"`
	Reference    string `gorm:"column:SortingReference;type:varchar(255);not null"`
	Type         string `gorm:"column:SortType;type:varchar(50);not null"`
	SortBy       string `gorm:"column:sortBy;type:varchar(50);not null"`
	CustomList   string `gorm:"column:customList;type:varchar(255)"`
	IconSet      string `gorm:"column:iconSet;type:varchar(255)"`
	IconID       string `gorm:"column:iconId;type:varchar(255)"`
}

// TableName overrides default table name for the model
func (Sorting) TableName() string {
	return "Sortings"
}

// PivotTable - pivot table
type PivotTable struct {
	ID           int
	DataSourceID int    `gorm:"column:DataSourceId"`
	Type         string `gorm:"column:Type;type:varchar(50)"`
	Label        string `gorm:"column:Label;type:varchar(255)"`
	DisplayName  string `gorm:"column:DisplayName;type:varchar(255)"`
	Function     string `gorm:"column:Function;type:varchar(255)"`
}

// TableName overrides default table name for the model
func (PivotTable) TableName() string {
	return "PivotTables"
}

// ConditionalFormatting - conditional formatting entries
type ConditionalFormatting struct {
	ID           int
	DataSourceID int    `gorm:"column:DataSourceId"`
	Type         string `gorm:"column:Type;type:varchar(50)"`
	Operator     string `gorm:"column:Operator;type:varchar(50)"`
	Formula1     string `gorm:"column:Formula1;type:varchar(255)"`
	Formula2     string `gorm:"column:Formula2;type:varchar(255)"`
	Formula3     string `gorm:"column:Formula3;type:varchar(255)"`
}

// TableName overrides default table name for the model
func (ConditionalFormatting) TableName() string {
	return "ConditionalFormattings"
}

// XLQTransformation - XLQ Transformations
type XLQTransformation struct {
	ID            int
	CellReference string    `gorm:"type:varchar(10);not null"`
	TimeStamp     time.Time `gorm:"not null"`
	UserID        int       `gorm:"column:UserID;not null;index"`
	Question      Question
	QuestionID    int `gorm:"column:QuestionID;not null;index"`
	Source        Source
	SourceID      int `gorm:"column:FileID;not null;index"`
}

// TableName overrides default table name for the model
func (XLQTransformation) TableName() string {
	return "XLQTransformation"
}

// User - users
type User struct {
	ID int `gorm:"column:UserID;primary_key;auto_increment"`
}

// TableName overrides default table name for the model
func (User) TableName() string {
	return "Users"
}

// AutoCommentAnswerCells adds automatic comment to the student answer cells
func AutoCommentAnswerCells() {
	var answers []Answer
	Db.Preload("Worksheets").Preload("Worksheets.Cells").
		Where("was_autocommented = ?", 0).
		Or("was_autocommented IS NULL").Find(&answers)
	for _, a := range answers {
		for _, w := range a.Worksheets {
			for _, c := range w.Cells {
				log.Info(c)
			}
		}
	}
}

// MatchPlagiarismKeys reads plagiarism key and match with the one stored
// in SpreadsheetTransformationTable (NB! the worksheets should be already imported)
func (wb *Workbook) MatchPlagiarismKeys(file *excelize.File) {
	var transformations []XLQTransformation
	Db.Joins(`JOIN StudentAnswers AS a 
		ON a.QuestionID = XLQTransformation.QuestionID
			AND a.FileID = XLQTransformation.FileID`).
		Joins("JOIN WorkSheets AS s ON s.StudentAnswerID = a.StudentAnswerID").
		Preload("Question").
		Preload("Question.Answers").
		Preload("Question.Answers.Worksheets").
		// Where("s.is_plagiarised IS NULL").
		Where("a.StudentAnswerID = ?", wb.AnswerID).
		Find(&transformations)

	for _, t := range transformations {
		keyValue := t.TimeStamp.UTC().Format(time.UnixDate) + " | " + strconv.Itoa(t.UserID)
		for _, a := range t.Question.Answers {
			for _, ws := range a.Worksheets {
				value := file.GetCellValue(ws.Name, t.CellReference)
				ws.IsPlagiarised = (value != keyValue) // sql.NullBool{Valid: true, Bool: (value == keyValue)}
				Db.Save(&ws)
			}
		}
	}
}
