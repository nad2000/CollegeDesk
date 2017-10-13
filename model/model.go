package models

import (
	"database/sql"
	"database/sql/driver"
	"fmt"
	"regexp"
	"strings"
	"time"

	log "github.com/Sirupsen/logrus"
	"github.com/jinzhu/gorm"
	//"github.com/tealeg/xlsx"
	"github.com/nad2000/xlsx"
)

// Db - shared DB connection
var Db *gorm.DB

// VerboseLevel - the level of verbosity
var VerboseLevel int

// DebugLevel - the level of verbosity of the debug information
var DebugLevel int

var cellIDRe = regexp.MustCompile("\\$?[A-Z]+\\$?[0-9]+")

func cellAddress(rowIndex, colIndex int) string {
	return xlsx.GetCellIDStringFromCoords(colIndex, rowIndex)

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
	ID                 int                 `gorm:"column:QuestionID;primary_key;AUTO_INCREMENT"`
	QuestionType       QuestionType        `gorm:"column:QuestionType"`
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
func (Question) TableName() string {
	return "Questions"
}

// ImportFile imports form Excel file QuestionExcleData
func (q *Question) ImportFile(fileName string) error {
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
	return nil
}

// QuestionExcelData - extracted celles from question Workbooks
type QuestionExcelData struct {
	ID         int      `gorm:"column:Id;primary_key;AUTO_INCREMENT"`
	QuestionID int      `gorm:"column:QuestionID"`
	SheetName  string   `gorm:"column:SheetName"`
	CellRange  string   `gorm:"column:CellRange"`
	Value      string   `gorm:"column:Value"`
	Comment    string   `gorm:"column:Comment"`
	Formula    string   `gorm:"column:Formula"`
	Question   Question `gorm:"ForeignKey:QuestionID"`
}

// TableName overrides default table name for the model
func (QuestionExcelData) TableName() string {
	return "QuestionExcelData"
}

// Source - student answer file sources
type Source struct {
	ID           int        `gorm:"column:FileID;primary_key;AUTO_INCREMENT"`
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

// Answer - student submitted answers
type Answer struct {
	ID             int           `gorm:"column:StudentAnswerID;primary_key;AUTO_INCREMENT"`
	AssignmentID   int           `gorm:"column:StudentAssignmentID"`
	QuestionID     sql.NullInt64 `gorm:"column:QuestionID;type:int"`
	MCQOptionID    int           `gorm:"column:MCQOptionID"`
	ShortAnswer    string        `gorm:"column:ShortAnswerText;type:text"`
	Marks          string        `gorm:"column:Marks"`
	SubmissionTime time.Time     `gorm:"column:SubmissionTime"`
	FileID         int           `gorm:"column:FileID"`
	Worksheets     []Worksheet   `gorm:"ForeignKey:AnswerID"`
	Source         Source        `gorm:"ForeignKey:FileID"`
	Question       Question      `gorm:"ForeignKey:QuestionID"`
}

// TableName overrides default table name for the model
func (Answer) TableName() string {
	return "StudentAnswers"
}

// Workbook - Excel file / workbook
type Workbook struct {
	ID         int `gorm:"primary_key"`
	FileName   string
	CreatedAt  time.Time
	AnswerID   sql.NullInt64 `gorm:"index;type:int"`
	Answer     Answer        `gorm:"ForeignKey:AnswerID"`
	Worksheets []Worksheet   `gorm:"ForeignKey:WorkbookID"`
}

// Reset deletes all underlying objects: worksheets, blocks, and cells
func (wb *Workbook) Reset() {

	var worksheets []Worksheet
	result := Db.Model(&wb).Related(&worksheets)
	if result.Error != nil {
		log.Error(result.Error)
	}
	log.Debugf("Deleting worksheets: %#v", worksheets)
	for ws := range worksheets {
		var blocks []Block
		Db.Model(&ws).Related(&blocks)
		for _, b := range blocks {
			log.Debugf("Deleting blocks: %#v", blocks)
			Db.Where("block_id = ?", b.ID).Delete(Cell{})
			Db.Delete(b)
		}
	}
	Db.Where("workbook_id = ?", wb.ID).Delete(Worksheet{})
}

// Worksheet - Excel workbook worksheet
type Worksheet struct {
	ID               int
	WorkbookID       int `gorm:"index"`
	Name             string
	WorkbookFileName string
	AnswerID         sql.NullInt64 `gorm:"index;type:int"`
	Blocks           []Block       `gorm:"ForeignKey:WorksheetID"`
	Answer           Answer        `gorm:"ForeignKey:AnswerID"`
	Workbook         Workbook      `gorm:"ForeignKey:WorkbookId"`
}

// Block - the univormly filled with specific color block
type Block struct {
	ID              int `gorm:"column:ExcelBlockID;primary_key;AUTO_INCREMENT"`
	WorksheetID     int `gorm:"index"`
	Color           string
	Range           string    `gorm:"column:BlockCellRange"`
	Formula         string    `gorm:"column:BlockFormula"` // first block cell formula
	RelativeFormula string    // first block cell relative formula formula
	Cells           []Cell    `gorm:"ForeignKey:BlockID"`
	Worksheet       Worksheet `gorm:"ForeignKey:WorksheetID"`

	s struct{ r, c int } `gorm:"-"` // Top-left cell
	e struct{ r, c int } `gorm:"-"` //  Bottom-right cell
}

func (b Block) String() string {
	return fmt.Sprintf("Block {Range: %q, Color: %q, Formula: %q, Relative Formula: %q}",
		b.Range, b.Color, b.Formula, b.RelativeFormula)
}

// TableName overrides default table name for the model
func (b Block) TableName() string {
	return "ExcelBlocks"
}

func (b *Block) save() {
	b.Range = b.address()
	Db.Save(b)
}

func (b *Block) address() string {
	return cellAddress(b.s.r, b.s.c) + ":" + cellAddress(b.e.r, b.e.c)
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

// fildWhole finds whole range of the specified color
// and the same "relative" formula starting with the set top-left cell.
func (b *Block) findWhole(sheet *xlsx.Sheet, color string) {

	b.e = b.s
	for i, row := range sheet.Rows {

		// skip all rows until the first block row
		if i < b.s.r {
			continue
		}

		log.Debugf("Total cells: %d at %d", len(row.Cells), i)
		// Range is discontinued or of a differnt color
		//log.Infof("*** b.e.c: %d, len: %d, %#v", b.e.c, len(row.Cells), row.Cells)
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
			if j < b.s.c {
				continue
			}

			fgColor := cell.GetStyle().Fill.FgColor
			relFormula := RelativeFormula(i, j, cell.Formula())
			// Reached the top-right corner:
			if fgColor == color && relFormula == b.RelativeFormula {
				cellID := cellAddress(i, j)
				commentText := ""
				comment, ok := sheet.Comment[cellID]
				if ok {
					commentText = comment.Text
				}
				c := Cell{
					BlockID: b.ID,
					Formula: cell.Formula(),
					Value:   cell.Value,
					Range:   cellID,
					Comment: commentText,
				}
				if DebugLevel > 1 {
					log.Debugf("Inserting %#v", c)
				}
				Db.Create(&c)
				if Db.Error != nil {
					log.Error("Error occured: ", Db.Error.Error())
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
}

// IsInside tests if the cell with given coordinates is inside the coordinates
func (b *Block) IsInside(r, c int) bool {
	return (b.s.r <= r &&
		r <= b.e.r &&
		b.s.c <= c &&
		c <= b.e.c)
}

// Cell - a sigle cell of the block
type Cell struct {
	ID      int
	BlockID int `gorm:"index"`
	Range   string
	Formula string
	Value   string
	Comment string
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

// SetDb initializes DB
func SetDb() {
	// Migrate the schema
	log.Debug("Add to automigrate...")
	worksheetsExists := Db.HasTable("worksheets")
	Db.AutoMigrate(&Source{})
	Db.AutoMigrate(&Question{})
	Db.AutoMigrate(&QuestionExcelData{})
	Db.AutoMigrate(&Answer{})
	Db.AutoMigrate(&Workbook{})
	Db.AutoMigrate(&Worksheet{})
	Db.AutoMigrate(&Block{})
	Db.AutoMigrate(&Cell{})
	if strings.HasPrefix(Db.Dialect().GetName(), "mysql") && !worksheetsExists {
		log.Debug("Adding a constraint to Wroksheets -> Answers...")
		Db.Model(&Worksheet{}).AddForeignKey("answer_id", "StudentAnswers(StudentAnswerID)", "CASCADE", "CASCADE")
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
	(Db.Joins(
		"JOIN FileSources ON FileSources.FileID = Questions.FileID").Where(
		"IsProcessed = ?", 0).Where(
		"FileSources.FileName LIKE ?", "%.xlsx").Find(
		&questions))
	return questions, Db.Error
}

// RowsToProcessResult stores query resut
type RowsToProcessResult struct {
	ID              int    `gorm:"column:FileID"`
	S3BucketName    string `gorm:"column:S3BucketName"`
	S3Key           string `gorm:"column:S3Key"`
	FileName        string `gorm:"column:FileName"`
	StudentAnswerID int    `gorm:"column:StudentAnswerID"`
}

// RowsToProcess returns answer file sources
func RowsToProcess() ([]RowsToProcessResult, error) {

	currentTime := time.Now()
	midnight := time.Date(
		currentTime.Year(),
		currentTime.Month(),
		currentTime.Day(),
		0, 0, 0, 0, time.UTC)

	// TODO: select file links from StudentAnswers and download them form S3 buckets..."
	rows, err := Db.Table("FileSources").Select(
		"FileSources.FileID, S3BucketName, S3Key, FileName, StudentAnswerID").Joins(
		"JOIN StudentAnswers ON StudentAnswers.FileID = FileSources.FileID").Where(
		"FileName IS NOT NULL").Where(
		"FileName != ''").Where(
		"FileName LIKE '%.xlsx'").Where(
		"SubmissionTime <= ?", midnight).Rows()
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
func ExtractBlocksFromFile(fileName, color string, force, verbose bool, answerIDs ...int) (wb Workbook) {
	xlFile, err := xlsx.OpenFile(fileName)
	if err != nil {
		log.Fatal(err)
	}

	var answerID sql.NullInt64
	if len(answerIDs) > 0 {
		answerID = NewNullInt64(answerIDs[0])
	}

	result := Db.First(&wb, Workbook{FileName: fileName, AnswerID: answerID})
	if !result.RecordNotFound() {
		if !force {
			log.Errorf("File %q was already processed.", fileName)
			return
		}
		log.Warnf("File %q was already processed.", fileName)
		wb.Reset()
	} else {
		wb = Workbook{FileName: fileName, AnswerID: answerID}
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

	for _, sheet := range xlFile.Sheets {

		if sheet.Hidden {
			log.Infof("Skipping hidden worksheet %q", sheet.Name)
			continue
		}

		if verbose {
			log.Infof("Processing worksheet %q", sheet.Name)
		}

		var ws Worksheet
		Db.FirstOrCreate(&ws, Worksheet{
			Name:             sheet.Name,
			WorkbookID:       wb.ID,
			WorkbookFileName: wb.FileName,
			AnswerID:         answerID,
		})
		if Db.Error != nil {
			log.Fatalf("Failed to create worksheet entry: %s", Db.Error.Error())
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
					}
					b.s.r, b.s.c = i, j

					Db.Create(&b)
					if DebugLevel > 1 {
						log.Debugf("Created %#v", b)
					}

					b.findWhole(sheet, color)
					b.save()
					blocks = append(blocks, b)
					if verbose {
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
	return
}
