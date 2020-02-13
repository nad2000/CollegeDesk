package tests

import (
	"extract-blocks/cmd"
	model "extract-blocks/model"
	"path"
	"testing"
)

func testHandleProblems(t *testing.T) {

	var files = []string{
		"Problem File Sample Before Utility is Run.xlsx",
		// "ProblemB4Utility.xlsx",
	}
	for _, fn := range files {
		var f model.Source
		if err := db.FirstOrCreate(&f, model.Source{
			FileName:     fn,
			S3BucketName: "studentanswers",
			S3Key:        fn,
		}).Error; err != nil {
			t.Error(err)
		}
		var p model.Problem
		if err := db.FirstOrCreate(&p, model.Problem{SourceID: f.ID}).Error; err != nil {
			t.Error(err)
		}
	}

	tm := testManager{SourceDirectory: path.Join("data", "problems")}
	cmd.HandleProblems(&tm)

	for _, fn := range files {
		var count int
		if err := db.Where("file_name = ?", fn).Model(&model.Workbook{}).Count(&count).Error; err != nil {
			t.Error(err)
		}
		if expected := 1; count != expected {
			t.Errorf("Expected %d question reference workbooks, got: %d", expected, count)
		}
	}

	// if err := db.DB().QueryRow(`
	// 	SELECT COUNT(DISTINCT reference_id) AS ReferenceCount
	// 	FROM Questions
	// 	WHERE QuestionText LIKE 'Question with merged cells #%'`).Scan(&count); err != nil {
	// 	t.Error(err)
	// }
	// if expected := 3; count != expected {
	// 	t.Errorf("Expected %d questions, got: %d", expected, count)
	// }

	// var blockCount int
	// if err := db.DB().QueryRow(`
	// 	SELECT count(*) AS RowCount, count(distinct b.BlockCellRange) AS BlockCount
	// 	FROM Questions AS q JOIN WorkSheets AS s ON s.workbook_id = q.reference_id
	// 	JOIN ExcelBlocks AS b ON b.worksheet_id = s.id
	// 	WHERE QuestionText LIKE 'Question with merged cells #%'`).Scan(&count, &blockCount); err != nil {
	// 	t.Error(err)
	// }
	// if expected := 24; count != expected {
	// 	t.Errorf("Expected %d rows, got: %d", expected, count)
	// }
	// if expected := 8; blockCount != expected {
	// 	t.Errorf("Expected %d different blocks, got: %d", expected, blockCount)
	// }
}
