package model

import (
	"fmt"
	"os"
	"path"
	"testing"

	log "github.com/Sirupsen/logrus"
	_ "github.com/jinzhu/gorm/dialects/sqlite"
	"github.com/nad2000/excelize"
)

func TestModel(t *testing.T) {
	testDbFileName := "/tmp/test_model.db"
	if _, err := os.Stat(testDbFileName); !os.IsNotExist(err) {
		os.RemoveAll(testDbFileName)
	}
	url := "sqlite://" + testDbFileName
	t.Log("TEST DB URL: ", url)
	_, err := OpenDb(url)
	if err != nil {
		log.Error(err)
		log.Fatalf("failed to connect database %q", url)
	}
	defer Db.Close()

	source := Source{S3Key: "KEY", FileName: "test.xlsx"}
	answer := Answer{
		Assignment: Assignment{Title: "TEST ASSIGNMENT", AssignmentSequence: 888},
		Marks:      98.7654,
		Source:     source,
		Question:   Question{QuestionType: "TYPE", Source: source, MaxScore: 98.76453},
	}
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
	for _, r := range []string{"A1", "B2", "C3"} {
		Db.Create(&Cell{
			Range:     r,
			Block:     block,
			Worksheet: block.Worksheet,
			Value:     "1234.567",
			Comment:   Comment{Text: fmt.Sprintf("JUST A COMMENT FOR %q", r)},
		})
	}
	var count int
	if err := Db.DB().QueryRow(`
			SELECT COUNT(DISTINCT worksheet_id) AS WorksheetCount
			FROM Cells`).Scan(&count); err != nil {
		t.Error(err)
	}
	if expected := 1; count != expected {
		t.Errorf("Expected to select %d worksheets linked to all cells, got: %d", expected, count)
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
