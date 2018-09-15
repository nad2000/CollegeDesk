package model

import (
	"database/sql"
	"database/sql/driver"
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

func cellAddress(rowIndex, colIndex int) string {
	return xlsx.GetCellIDStringFromCoords(colIndex, rowIndex)
}

// RelCellAddress - relative cell R1R1 representation against the given cell
func RelCellAddress(address string, rowIncrement, colIncrement int) (string, error) {
	colIndex, rowIndex, err := xlsx.GetCoordsFromCellIDString(address)
	if err != nil {
		log.Errorf("Failed to map address %q: %s", address, err.Error())
		return "", err
	}
	return xlsx.GetCellIDStringFromCoords(colIndex+colIncrement, rowIndex+rowIncrement), nil
}

// RelativeCellAddress converts cell ID into a relative R1C1 representation
func RelativeCellAddress(rowIndex, colIndex int, cellID string) string {
	x, y, err := xlsx.GetCoordsFromCellIDString(cellID)
	if err != nil {
		log.Fatalf("Failed to find coordinates for %q: %s", cellID, err.Error())
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

// NewNullInt64 - a helper function that makes nullable from a plain int
func NewNullInt64(value int) sql.NullInt64 {
	return sql.NullInt64{Valid: true, Int64: int64(value)}
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
	WasCompared        bool           `gorm:"type:tinyint(1)"`
	IsProcessed        bool           `gorm:"column:IsProcessed;type:tinyint(1);default:0"`
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
	xlFile, err := xlsx.OpenFile(fileName)

	if err != nil {
		return err
	}

	if VerboseLevel > 0 {
		log.Infof("Processing workbook: %s", fileName)
	}

	for _, sheet := range xlFile.Sheets {

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

				cellRange := cellAddress(i, j)
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
	wb := ExtractBlocksFromFile(fileName, color, true, verbose, true)
	q.ReferenceID = NewNullInt64(wb.ID)
	Db.Save(&q)
	return nil
}

// QuestionExcelData - extracted celles from question Workbooks
type QuestionExcelData struct {
	ID         int    `gorm:"column:Id;primary_key:true;AUTO_INCREMENT"`
	SheetName  string `gorm:"column:SheetName"`
	CellRange  string `gorm:"column:CellRange"`
	Value      string `gorm:"column:Value"`
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
	ID                  int           `gorm:"column:StudentAnswerID;primary_key:true;AUTO_INCREMENT"`
	AssignmentID        int           `gorm:"column:StudentAssignmentID"`
	MCQOptionID         sql.NullInt64 `gorm:"column:MCQOptionID;type:int"`
	ShortAnswer         string        `gorm:"column:ShortAnswerText;type:text"`
	Marks               float64       `gorm:"column:Marks"`
	SubmissionTime      time.Time     `gorm:"column:SubmissionTime"`
	Worksheets          []Worksheet   `gorm:"ForeignKey:AnswerID"`
	Source              Source        `gorm:"Association_ForeignKey:FileID"`
	SourceID            int           `gorm:"column:FileID"`
	Question            Question
	QuestionID          sql.NullInt64   `gorm:"column:QuestionID;type:int"`
	WasCommentProcessed uint8           `gorm:"type:tinyint unsigned;default:0"`
	WasXLProcessed      uint8           `gorm:"type:tinyint unsigned;default:0"`
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
	// CourseID           uint      `gorm:"column:CourseID;type:int(10) unsigned"`
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
	Answer      Answer        `gorm:"ForeignKey:AnswerID"`
	Worksheets  []Worksheet   `gorm:"ForeignKey:WorkbookID"`
	IsReference bool          // the workbook is used for referencing the expected bloks
}

// TableName overrides default table name for the model
func (Workbook) TableName() string {
	return "WorkBooks"
}

// Reset deletes all underlying objects: worksheets, blocks, and cells
func (wb *Workbook) Reset() {

	var worksheets []Worksheet
	result := Db.Where("workbook_id = ?", wb.ID).Find(&worksheets)
	if result.Error != nil {
		log.Error(result.Error)
	}
	log.Debugf("Deleting worksheets: %#v", worksheets)
	for _, ws := range worksheets {
		Db.Delete(Chart{}, "worksheet_id = ?", ws.ID)
		var blocks []Block
		Db.Model(&ws).Related(&blocks)
		result := Db.Where("worksheet_id = ?", ws.ID).Find(&blocks)
		if result.Error != nil {
			log.Error(result.Error)
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
						Db.LogMode(true)
						Db.Model(&cell).UpdateColumn("comment", t.T)
						Db.LogMode(false)
					}
				}
			}

		}
	}
	return
}

// ImportCharts - import charts form workbook file
func (wb *Workbook) ImportCharts(fileName string) {

	xlFile, err := excelize.OpenFile(fileName)
	if err != nil {
		log.Errorf("Failed to open file %q", fileName)
		log.Errorln(err)
		return
	}

	for _, sheet := range xlFile.WorkBook.Sheets.Sheet {
		var ws Worksheet
		result := Db.First(&ws, Worksheet{
			Name:       sheet.Name,
			AnswerID:   wb.AnswerID,
			WorkbookID: wb.ID,
		})
		if result.RecordNotFound() && !DryRun {
			ws = Worksheet{
				Name:             sheet.Name,
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

		name := "xl/worksheets/_rels/sheet" + sheet.SheetID + ".xml.rels"
		sheetRels := unmarshalRelationships(xlFile.XLSX[name])
		for _, r := range sheetRels.Relationships {
			if strings.Contains(r.Target, "drawings/drawing") {
				name := "xl/drawings/_rels/" + filepath.Base(r.Target) + ".rels"
				drawing := unmarshalDrawing(xlFile.XLSX["xl/drawings/"+filepath.Base(r.Target)])
				drawingRels := unmarshalRelationships(xlFile.XLSX[name])
				for _, dr := range drawingRels.Relationships {
					if strings.Contains(dr.Target, "charts/chart") {
						chartName := "xl/charts/" + filepath.Base(dr.Target)
						chart := unmarshalChart(xlFile.XLSX[chartName])
						chartTitle := chart.Title.Value()
						log.Debugf("*** %s: %#v", chartName, chart)
						log.Infof("Found %q chart (titled: %q) on the sheet %q", chart.Type(), chartTitle, sheet.Name)
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
								RelativeFormula: cellAddress(
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
}

// TableName overrides default table name for the model
func (Worksheet) TableName() string {
	return "WorkSheets"
}

// Block - the uniformly filled with specific color block
type Block struct {
	ID              int `gorm:"column:ExcelBlockID;primary_key:true;AUTO_INCREMENT"`
	Color           string
	Range           string                `gorm:"column:BlockCellRange"`
	Formula         string                `gorm:"column:BlockFormula"` // first block cell formula
	RelativeFormula string                // first block cell relative formula formula
	Cells           []Cell                `gorm:"ForeignKey:BlockID"`
	Worksheet       Worksheet             `gorm:"ForeignKey:WorksheetID"`
	WorksheetID     int                   `gorm:"index"`
	CommentMappings []BlockCommentMapping `gorm:"ForeignKey:ExcelBlockID"`
	Chart           Chart                 `gorm:"ForeignKey:ChartId"`
	ChartID         sql.NullInt64
	IsReference     bool // the block is used for referencing the expected bloks
	Row, Col        int  // the block top-left cell coordinates

	s       struct{ r, c int }           `gorm:"-"` // Top-left cell
	e       struct{ r, c int }           `gorm:"-"` // Bottom-right cell
	i       struct{ sr, sc, er, ec int } `gorm:"-"` // "Inner" block - the block containing values
	isEmpty bool                         `gorm:"-"` // All block cells are empty
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
		"%s {ID: %d, Range: %q, Color: %q, Formula: %q, Relative Formula: %q, WorksheetID: %d}",
		output, b.ID, b.Range, b.Color, b.Formula, b.RelativeFormula, b.WorksheetID)
}

func (b *Block) save() {
	if !DryRun {
		if !b.IsReference && b.Color != "" {
			for i := b.Col; i <= b.e.c; i++ {
				for j := b.Row; j <= b.e.r; j++ {
					address := cellAddress(j, i)
					address += ":" + address
					if b.isEmpty || i < b.i.sc || i > b.i.ec || j < b.i.sr || j > b.i.er {
						empty := Block{
							WorksheetID: b.WorksheetID,
							Range:       address,
							Color:       b.Color,
							Row:         j,
							Col:         i,
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
				b.Range = b.address()
			} else {
				b.Range = b.innerAddress()
			}
			Db.Save(b)
		}
	}
}

func (b *Block) address() string {
	return cellAddress(b.Row, b.Col) + ":" + cellAddress(b.e.r, b.e.c)
}

func (b *Block) innerAddress() string {
	return cellAddress(b.i.sr, b.i.sc) + ":" + cellAddress(b.i.er, b.i.ec)
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
			log.Error(err.Error())
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

	b.e.r, b.e.c = b.Row, b.Col
	for i, row := range sheet.Rows {

		// skip all rows until the first block row
		if i < b.Row {
			continue
		}

		log.Debugf("Total cells: %d at %d", len(row.Cells), i)
		// Range is discontinued or of a differnt color
		if len(row.Cells) <= b.e.c ||
			row.Cells[b.e.c].GetStyle().Fill.FgColor != color ||
			RelativeFormula(i, b.e.c, row.Cells[b.e.c].Formula()) != b.RelativeFormula {
			log.Debugf("Reached the edge row of the block at row %d", i)
			b.e.r = i - 1
			break
		} else {
			b.e.r = i
		}

		for j, cell := range row.Cells {
			// skip columns until the start:
			if j < b.Col {
				continue
			}

			// Reached the top-right corner:
			if fgColor := cell.GetStyle().Fill.FgColor; fgColor == color {
				if !b.IsReference {
					relFormula := RelativeFormula(i, j, cell.Formula())
					if relFormula == b.RelativeFormula {
						cellID := cellAddress(i, j)
						commentText := ""
						comment, ok := sheet.Comment[cellID]
						if ok {
							commentText = comment.Text
						}
						if value := cellValue(cell); value != "" {
							c := Cell{
								BlockID:     b.ID,
								WorksheetID: b.WorksheetID,
								Formula:     cell.Formula(),
								Value:       value,
								Range:       cellID,
								Comment:     commentText,
							}
							if DebugLevel > 1 {
								log.Debugf("Inserting %#v", c)
							}
							Db.Create(&c)
							if Db.Error != nil {
								log.Error("Error occured: ", Db.Error.Error())
							}
						}
					}
				}
				b.e.c = j
			} else {
				log.Debugf("Reached the edge column  of the block at column %d", j)
				if j > b.e.c {
					b.e.c = j - 1
				}
				break
			}
		}
	}

	if b.IsReference {
		return
	}
	// Find the part containing values
	sr, sc, er, ec := b.Row, b.Col, b.e.r, b.e.c
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
	return (b.Row <= r &&
		r <= b.e.r &&
		b.Col <= c &&
		c <= b.e.c)
}

// Cell - a sigle cell of the block
type Cell struct {
	ID          int
	Block       Block `gorm:"ForeignKey:BlockID"`
	BlockID     int   `gorm:"index"`
	Worksheet   Worksheet
	WorksheetID int `gorm:"index"`
	Range       string
	Formula     string
	Value       string
	Comment     string
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
	FileID             sql.NullInt64       `gorm:"column:FileID;type:int"`
	AuthorUserID       int                 `gorm:"column:AuthorUserID;not null"`
	WasCompared        bool                `gorm:"column:was_compared;type:tinyint(1)"`
	IsProcessed        bool                `gorm:"column:IsProcessed;type:tinyint(1);default:0"`
	Source             Source              `gorm:"ForeignKey:FileID"`
	Answers            []Answer            `gorm:"ForeignKey:QuestionID"`
	QuestionExcelDatas []QuestionExcelData `gorm:"ForeignKey:QuestionID"`
}

// TableName overrides default table name for the model
func (MySQLQuestion) TableName() string {
	return "Questions"
}

// Comment - added comments  with marks
type Comment struct {
	ID              int                   `gorm:"column:CommentID;primary_key:true;AUTO_INCREMENT"`
	Text            string                `gorm:"column:CommentText"`
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
	BlockID   uint `gorm:"column:ExcelBlockID"`
	Comment   Comment
	CommentID uint `gorm:"column:ExcelCommentID"`
}

// TableName overrides default table name for the model
func (BlockCommentMapping) TableName() string {
	return "BlockCommentMapping"
}

// AnswerComment - answer-comment mapping:
type AnswerComment struct {
	Answer    Answer
	AnswerID  int `gorm:"column:StudentAnswerID;index;type:int unsigned"`
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
	AssignmentID int `gorm:"column:AssignmentID;type:uint"`
	Question     Question
	QuestionID   int `gorm:"column:QuestionID;type:uint"`
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
	worksheetsExists := Db.HasTable("WorkSheets")
	Db.AutoMigrate(&Source{})
	if isMySQL {
		// Modify struct tag for MySQL
		Db.AutoMigrate(&MySQLQuestion{})
	} else {
		Db.AutoMigrate(&Question{})
	}
	Db.AutoMigrate(&QuestionExcelData{})
	Db.AutoMigrate(&Answer{})
	Db.AutoMigrate(&Workbook{})
	Db.AutoMigrate(&Worksheet{})
	Db.AutoMigrate(&Chart{})
	Db.AutoMigrate(&Block{})
	Db.AutoMigrate(&Cell{})
	Db.AutoMigrate(&Comment{})
	Db.AutoMigrate(&BlockCommentMapping{})
	Db.AutoMigrate(&Assignment{})
	Db.AutoMigrate(&QuestionAssignment{})
	Db.AutoMigrate(&AnswerComment{})
	if isMySQL && !worksheetsExists {
		// Add some foreing key constraints to MySQL DB:
		log.Debug("Adding a constraint to Wroksheets -> Answers...")
		Db.Model(&Worksheet{}).AddForeignKey("StudentAnswerID", "StudentAnswers(StudentAnswerID)", "CASCADE", "CASCADE")
		log.Debug("Adding a constraint to Cells...")
		Db.Model(&Cell{}).AddForeignKey("block_id", "ExcelBlocks(ExcelBlockID)", "CASCADE", "CASCADE")
		log.Debug("Adding a constraint to Blocks...")
		Db.Model(&Block{}).AddForeignKey("worksheet_id", "worksheets(id)", "CASCADE", "CASCADE")
		log.Debug("Adding a constraint to Worksheets -> Workbooks...")
		Db.Model(&Worksheet{}).AddForeignKey("workbook_id", "workbooks(id)", "CASCADE", "CASCADE")
		log.Debug("Adding a constraint to Questions...")
		Db.Model(&Question{}).AddForeignKey("FileID", "FileSources(FileID)", "CASCADE", "CASCADE")
		log.Debug("Adding a constraint to QuestionExcelData...")
		Db.Model(&QuestionExcelData{}).AddForeignKey("QuestionID", "Questions(QuestionID)", "CASCADE", "CASCADE")
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
func RowsToComment() ([]RowsToProcessResult, error) {
	rows, err := Db.Table("FileSources").
		Select("DISTINCT FileSources.FileID, S3BucketName, S3Key, FileName, StudentAnswerID").
		Joins("JOIN StudentAnswers ON StudentAnswers.FileID = FileSources.FileID").
		Joins("JOIN Questions ON Questions.QuestionID = StudentAnswers.QuestionID").
		Joins("JOIN QuestionAssignmentMapping ON QuestionAssignmentMapping.QuestionID = Questions.QuestionID").
		// Joins("JOIN CourseAssignments ON CourseAssignments.AssignmentID = QuestionAssignmentMapping.AssignmentID").
		Where("was_comment_processed = ?", 0).
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
		`).
		// Where("CourseAssignments.State = ?", "GRADED").
		Rows()
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

// alreadyFound tests if the range containing the cell
// coordinates hhas been already found.
func (bl *blockList) alreadyFound(r, c int) bool {
	for _, b := range *bl {
		if b.IsInside(r, c) {
			return true
		}
	}
	return false
}

// ExtractBlocksFromFile extracts blocks from the given file and stores in the DB
func ExtractBlocksFromFile(fileName, color string, force, verbose, isReference bool, answerIDs ...int) (wb Workbook) {
	xlFile, err := xlsx.OpenFile(fileName)
	if err != nil {
		log.Fatal(err)
	}

	var answerID sql.NullInt64
	if !isReference && len(answerIDs) > 0 {
		answerID = NewNullInt64(answerIDs[0])
	}

	var result *gorm.DB
	if isReference {
		result = Db.First(&wb, Workbook{FileName: fileName, IsReference: isReference})
	} else {
		result = Db.First(&wb, Workbook{FileName: fileName, AnswerID: answerID})
	}
	if !result.RecordNotFound() {
		if !force {
			log.Errorf("File %q was already processed.", fileName)
			return
		}
		log.Warnf("File %q was already processed.", fileName)
		if !DryRun {
			wb.Reset()
		}
	} else if !DryRun {

		wb = Workbook{FileName: fileName, AnswerID: answerID, IsReference: isReference}
		result = Db.Create(&wb)
		if result.Error != nil {
			log.Fatalf("Failed to create workbook entry %#v: %s", wb, result.Error.Error())
		}
		if DebugLevel > 1 {
			log.Debugf("Ceated workbook entry %#v", wb)
		}
	}

	if verbose {
		log.Infof("*** Processing workbook: %s", fileName)
	}
	var q Question
	if answerID.Valid && !isReference {
		// Attempt to use reference blocks for the answer:
		result := Db.
			Joins("JOIN StudentAnswers ON StudentAnswers.QuestionID = Questions.QuestionID").
			Where("StudentAnswers.StudentAnswerID = ?", answerID).
			Where("Questions.reference_id IS NOT NULL").
			First(&q)
		if result.Error != nil {
			log.Error(result.Error)
		}
		if verbose {
			log.Infof("*** Processing the answer ID:%d for the queestion %s", answerID.Int64, q)
		}
	}

	for orderNum, sheet := range xlFile.Sheets {

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
			Db.FirstOrCreate(&ws, Worksheet{
				Name:             sheet.Name,
				WorkbookID:       wb.ID,
				WorkbookFileName: wb.FileName,
				AnswerID:         answerID,
				IsReference:      isReference,
				OrderNum:         orderNum,
			})
		}
		if Db.Error != nil {
			log.Fatalf("*** Failed to create worksheet entry: %s", Db.Error.Error())
		}

		if answerID.Valid && !isReference && q.ReferenceID.Valid {
			// Attempt to use reference blocks for the answer:
			result = Db.
				Joins("JOIN WorkSheets ON WorkSheets.id = ExcelBlocks.worksheet_id").
				Where("ExcelBlocks.is_reference").
				Where("WorkSheets.workbook_id = ?", q.ReferenceID).
				Where("Worksheets.order_num = ?", orderNum).
				Find(&references)
			if result.Error != nil {
				log.Error(result.Error)
			}
		}

		blocks := blockList{}
		sheetFillColors := []string{}

		for i, row := range sheet.Rows {
			for j, cell := range row.Cells {

				if blocks.alreadyFound(i, j) {
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
						IsReference:     isReference,
						Row:             i,
						Col:             j,
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
		if len(blocks) == 0 && references == nil {
			log.Warningf("No block found ot the worksheet %q of the workbook %q with color %q", sheet.Name, fileName, color)
			if len(sheetFillColors) > 0 {
				log.Infof("Following colors were found in the worksheet you could use: %v", sheetFillColors)
			}
		}

		if answerID.Valid && references != nil {
			// Attempt to use reference blocks for the answer:
			for _, rb := range references {
				var (
					b              Block
					cell           *xlsx.Cell
					sc, sr, ec, er int
					err            error
				)
				rangeAddr := strings.Split(rb.Range, ":")
				if len(rangeAddr) < 2 {
					log.Errorln("Incorrect block range: ", rb)
					continue
				}
				sc, sr, err = xlsx.GetCoordsFromCellIDString(rangeAddr[0])
				if err != nil {
					log.Error(err)
					continue
				}
				for _, b := range blocks {
					if b.Row == rb.Row && b.Col == rb.Col {
						goto NEXT
					}
				}
				ec, er, err = xlsx.GetCoordsFromCellIDString(rangeAddr[1])
				if err != nil {
					log.Error(err)
					continue
				}
				cell = sheet.Cell(sr, sc)
				b = Block{
					WorksheetID:     ws.ID,
					Formula:         cell.Formula(),
					RelativeFormula: RelativeFormula(sr, sc, cell.Formula()),
					Range:           rb.Range,
					Row:             sr,
					Col:             sc,
				}
				Db.Create(&b)
				for r := sr; r <= er; r++ {
					for c := sc; c <= ec; c++ {
						var b Block
						cell := sheet.Cell(sc, sr)
						if value := cellValue(cell); value != "" {
							c := Cell{
								BlockID:     b.ID,
								WorksheetID: ws.ID,
								Formula:     cell.Formula(),
								Value:       value,
								Range:       cellAddress(r, c),
							}
							if DebugLevel > 1 {
								log.Debugf("Inserting %#v", c)
							}
							Db.Create(&c)
							if Db.Error != nil {
								log.Error("Error occured: ", Db.Error.Error())
							}
						}
					}
				}
			NEXT:
			}
		}
	}

	wb.ImportCharts(fileName)

	if _, err := Db.DB().
		Exec(`
			UPDATE StudentAnswers
			SET was_xl_processed = 1
			WHERE StudentAnswerID = ?`, answerID); err != nil {
		log.Errorf("Failed to update the answer entry: %s", err.Error())
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
		log.Errorf("Failed to insert block -> comment mapping: %s", err.Error())
	}

	// Insert block -> comment mapping:
	sql = `
		INSERT INTO StudentAnswerCommentMapping(StudentAnswerID, CommentID)
		SELECT DISTINCT StudentAnswerID, ExcelCommentID FROM (` +
		commentsToMapSQL + ") AS c"
	_, err = Db.DB().Exec(sql, answerID)
	if err != nil {
		log.Info("SQL: ", sql)
		log.Errorf("Failed to insert block -> comment mapping: %s", err.Error())
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
