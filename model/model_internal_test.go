package model

import (
	"fmt"
	"os"
	"path"
	"testing"

	log "github.com/Sirupsen/logrus"
	_ "github.com/jinzhu/gorm/dialects/sqlite"
	"github.com/nad2000/excelize"
	"github.com/nad2000/xlsx"
)

func TestModel(t *testing.T) {
	url, ok := os.LookupEnv("DATABASE_URL")
	if !ok {
		testDbFileName := "/tmp/test_model.db"
		if _, err := os.Stat(testDbFileName); !os.IsNotExist(err) {
			os.RemoveAll(testDbFileName)
		}
		url = "sqlite://" + testDbFileName
	}
	t.Log("TEST DB URL: ", url)
	_, err := OpenDb(url)
	if err != nil {
		log.Error(err)
		log.Fatalf("failed to connect database %q", url)
	}
	defer Db.Close()
	t.Log("DIALECT: ", Db.Dialect().GetName(), Db.Dialect().CurrentDatabase())
	if Db.Dialect().GetName() == "mysql" {
		Db.Exec("DELETE FROM ProblemWorkSheetExcelData")
		Db.Exec("DELETE FROM ProblemWorkSheets")
		Db.Exec("DELETE FROM QuestionFiles")
		Db.Exec("DELETE FROM Problems")
		Db.Exec("TRUNCATE TABLE DefinedNames")
		Db.Exec("TRUNCATE TABLE AutoEvaluation")
		Db.Exec("DELETE FROM Cells")
		Db.Exec("TRUNCATE TABLE BlockCommentMapping")
		Db.Exec("TRUNCATE TABLE StudentAnswerCommentMapping")
		Db.Exec("TRUNCATE TABLE QuestionAssignmentMapping")
		Db.Exec("TRUNCATE TABLE QuestionExcelData")
		Db.Exec("TRUNCATE TABLE Rubrics")
		Db.Exec("DELETE FROM Questions")
		Db.Exec("DELETE FROM ExcelBlocks")
		Db.Exec("DELETE FROM StudentAssignments")
		Db.Exec("DELETE FROM CourseAssignments")
		Db.Exec("DELETE FROM FileSources")
		Db.Exec("DELETE FROM WorkSheets")
		Db.Exec("DELETE FROM WorkBooks")
		Db.Exec("TRUNCATE TABLE Comments")
	}
	// Db.LogMode(true)
	source := Source{S3Key: "KEY", FileName: "test.xlsx"}
	Db.Create(&source)
	answer := Answer{
		Assignment: Assignment{Title: "TEST ASSIGNMENT", AssignmentSequence: 888},
		Marks:      98.7654,
		Source:     source,
		Question: Question{
			QuestionType: "FileUpload",
			Source:       source,
			MaxScore:     98.76453,
		},
	}
	Db.Create(&answer)
	block := Block{
		Color: "FF00AA0000",
		Range: "A1:C3",
		Worksheet: Worksheet{
			Name:   "Sheet1",
			Answer: answer,
			Workbook: Workbook{
				FileName: "test.xlsx",
				Answer:   answer,
			},
		},
	}

	Db.Create(&block)
	var wb Workbook
	Db.Preload("Worksheets").First(&wb, "file_name = ?", "test.xlsx")
	if wb.Worksheets == nil || len(wb.Worksheets) < 1 {
		t.Error("Failed to get worksheets:", wb)
	}
	for _, r := range []string{"A1", "B2", "C3"} {
		cell := Cell{
			Range:     r,
			Block:     block,
			Worksheet: block.Worksheet,
			Value:     "1234.567",
			Comment:   Comment{Text: fmt.Sprintf("JUST A COMMENT FOR %q", r)},
		}
		cell.Row, cell.Col, _ = xlsx.GetCoordsFromCellIDString(r)
		Db.Create(&cell)
		if r == "C3" {
			Db.Create(&AutoEvaluation{
				CellID:         cell.ID,
				IsValueCorrect: true,
			})
		}
	}
	Db.Create(&BlockCommentMapping{
		Block:   block,
		Comment: Comment{Text: fmt.Sprintf("A COMMENT FOR BLOCK %q", block.Range)},
	})
	var count int
	if err := Db.DB().QueryRow(`
			SELECT COUNT(DISTINCT worksheet_id) AS WorksheetCount
			FROM Cells`).Scan(&count); err != nil {
		t.Error(err)
	}
	if expected := 1; count != expected {
		t.Errorf("Expected to select %d worksheets linked to all cells, got: %d", expected, count)
	}
	// Db.LogMode(false)
	var answers []Answer
	Db.Preload("Worksheets").
		Preload("Worksheets.Cells").
		Preload("Worksheets.Cells.AutoEvaluation").
		Where("was_autocommented = ?", 0).Find(&answers)

	var (
		blocks    blockList
		questions []Question
	)
	Db.Find(&blocks)
	Db.Find(&questions)
	for qi, q := range questions {
		for bi, b := range blocks {
			Db.Create(&Rubric{
				NumCell:    qi*bi + 1,
				BlockID:    b.ID,
				QuestionID: q.ID,
			})
		}
	}
}

func TestNormalizeFloatRepr(t *testing.T) {
	if expected, got := "0.16", normalizeFloatRepr("0.16000000000000003"); got != expected {
		t.Errorf("Failed not normalizeFloatRepr, expected: %q, got: %q", expected, got)
	}
	if expected, got := "-0.08", normalizeFloatRepr("-8.0000000000000016E-2"); got != expected {
		t.Errorf("Failed not normalizeFloatRepr, expected: %q, got: %q", expected, got)
	}
}

func TestRemoveComments(t *testing.T) {
	fileName := "with_comments.xlsx"
	outputName := path.Join(os.TempDir(), "output_with_comments0.xlsx")
	file, err := excelize.OpenFile(fileName)
	if err != nil {
		log.Errorf("Failed to open file %q", fileName)
		log.Errorln(err)
		return
	}
	DeleteAllComments(file)
	file.SaveAs(outputName)
	log.Infoln("Output: ", outputName)
}
