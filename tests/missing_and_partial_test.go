package tests

import (
	model "extract-blocks/model"
	"testing"
)

// TestMissingOrPartialMultipleTypes tests partial and unanswered questions.
func TestMissingOrPartialMultipleTypes(t *testing.T) {

	db = createTestDB()
	defer db.Close()

	assignment := model.Assignment{
		Title: "Full Cycel Testing...",
		State: "READY_FOR_GRADING",
	}
	db.Create(&assignment)
	for _, r := range []struct {
		base, questionFileName, modelAnswerFileName, anserFileName, cr, ts string
		uid                                                                int
		isPlagiarised                                                      bool
	}{
		{"data/missing-or-partial/Conditional Formatting/", "CFQuestion.xlsx", "CFModelAnswer.xlsx", "CFStudentAnswer.xlsx", "KE4423", "2018-12-27 19:18:05", 4951, false},
		{"data/missing-or-partial/Filter/", "FilterQuestion.xlsx", "FilterModelAnswer.xlsx", "FilterStudentAnswer.xlsx", "KE4423", "2018-12-27 19:18:05", 4951, false},
		{"data/missing-or-partial/Solver/", "Solver Question.xlsx", "Solver ModelAnswer.xlsx", "Solver StudentlAnswer.xlsx", "KE4423", "2018-12-27 19:18:05", 4951, false},
	} {

		qf := model.Source{
			FileName:     r.base + r.questionFileName,
			S3BucketName: "studentanswers",
			S3Key:        r.base + r.questionFileName,
		}
		db.Create(&qf)
		q := model.Question{
			SourceID:     model.NewNullInt64(qf.ID),
			QuestionType: model.QuestionType("FileUpload"),
			QuestionText: r.base + r.questionFileName,
			MaxScore:     1010.88,
			AuthorUserID: 123456789,
			WasCompared:  false,
			IsFormatting: true,
		}
		if err := db.Create(&q).Error; err != nil {
			t.Error(err)
		}
		q.ImportFile(r.base+r.questionFileName, "FFFFFF00", true, true)

		// Model answer:
		msa := model.StudentAssignment{
			UserID:       10000,
			AssignmentID: assignment.ID,
		}
		db.Create(&msa)
		maf := model.Source{FileName: r.base + r.modelAnswerFileName, S3BucketName: "studentanswers"}
		db.Create(&maf)
		ma := model.Answer{
			SourceID:            model.NewNullInt64(maf.ID),
			QuestionID:          model.NewNullInt64(q.ID),
			SubmissionTime:      *parseTime("2018-09-30 12:42"),
			StudentAssignmentID: msa.ID,
		}
		db.Create(&ma)
		model.ExtractBlocksFromFile(r.base+r.modelAnswerFileName, "FFFFFF00", true, true, true, ma.ID)

		// Answer
		af := model.Source{FileName: r.base + r.anserFileName, S3BucketName: "studentanswers"}
		db.Create(&af)

		sa := model.StudentAssignment{
			UserID:       r.uid,
			AssignmentID: assignment.ID,
		}
		db.Create(&sa)

		a := model.Answer{
			SourceID:            model.NewNullInt64(af.ID),
			QuestionID:          model.NewNullInt64(q.ID),
			SubmissionTime:      *parseTime("2018-09-30 12:42"),
			StudentAssignmentID: sa.ID,
		}
		db.Create(&a)
		// if r.uid > 0 {
		// 	db.Create(&model.XLQTransformation{
		// 		CellReference: r.cr,
		// 		UserID:        r.uid,
		// 		TimeStamp:     *parseTime(r.ts),
		// 		QuestionID:    q.ID,
		// 	})
		// 	// Create extar entries for  non-pagiarised examples
		// 	if !r.isPlagiarised {
		// 		for i := 1; i < 10; i++ {
		// 			db.Create(&model.XLQTransformation{
		// 				CellReference: "A" + strconv.Itoa(i),
		// 				UserID:        r.uid,
		// 				TimeStamp:     *parseTime(r.ts),
		// 				QuestionID:    q.ID,
		// 			})
		// 		}
		// 	}
		// }
		model.ExtractBlocksFromFile(r.base+r.anserFileName, "FFFFFF00", true, true, true, a.ID)
		// Test if is marked plagiarised:
		// {
		// 	var ws model.Worksheet
		// 	db.Where("workbook_file_name = ?", r.base+r.anserFileName).First(&ws)
		// 	if ws.IsPlagiarised != r.isPlagiarised {
		// 		t.Errorf("Exected that %#v will get marked as plagiarised.", ws)
		// 	}
		// }
	}

	// // Auto-commenting
	// if err := db.Exec(`
	// 	INSERT INTO AutoEvaluation (cell_id, IsValueCorrect, IsFormulaCorrect, is_hardcoded)
	// 	SELECT c.id, c.id%3 = 1, c.id%3 = 0, c.id%4 =0
	// 	FROM Cells AS c LEFT OUTER JOIN AutoEvaluation AS ae ON ae.cell_id = c.id
	// 	WHERE ae.cell_id IS NULL AND c.id % 3 != 2`).Error; err != nil {
	// 	t.Error(err)
	// }
	// var countBefore int
	// if err := db.Model(&model.AutoEvaluation{}).Count(&countBefore).Error; err != nil {
	// 	t.Error(err)
	// }
	// model.AutoCommentAnswerCells(12345, 10000)

	// var countAfter int
	// if err := db.Model(&model.AutoEvaluation{}).Count(&countAfter).Error; err != nil {
	// 	t.Error(err)
	// }
	// if countAfter != countBefore {
	// 	t.Errorf(
	// 		"Exected unchanged rowcount of AutoEvaluation table. Expected: %d, got: %d",
	// 		countBefore, countAfter)
	// }
}

// TestMissingOrPartialFilter tests partial and unanswered questions with 2 answers submitted.
func TestMissingOrPartialFilter(t *testing.T) {

	db = createTestDB()
	defer db.Close()

	assignment := model.Assignment{
		Title: "Test Missing Or Partial / Filter...",
		State: "READY_FOR_GRADING",
	}
	db.Create(&assignment)
	type answerRec struct {
		fileName string
		uid      int
	}
	for _, r := range []struct {
		base, questionFileName, modelAnswerFileName string
		answers                                     []answerRec
	}{
		{"data/missing-or-partial/Filter_2_answers/", "FilterQuestion.xlsx", "FilterModelAnswer.xlsx",
			[]answerRec{
				{"FilterStudentAnswer1.xlsx", 4951},
				{"FilterStudentAnswer2.xlsx", 4952},
			},
		}} {

		qf := model.Source{
			FileName:     r.base + r.questionFileName,
			S3BucketName: "studentanswers",
			S3Key:        r.base + r.questionFileName,
		}
		db.Create(&qf)
		q := model.Question{
			SourceID:     model.NewNullInt64(qf.ID),
			QuestionType: model.QuestionType("FileUpload"),
			QuestionText: r.base + r.questionFileName,
			MaxScore:     1010.88,
			AuthorUserID: 123456789,
			WasCompared:  false,
			IsFormatting: true,
		}
		if err := db.Create(&q).Error; err != nil {
			t.Error(err)
		}
		q.ImportFile(r.base+r.questionFileName, "FFFFFF00", true, true)

		// Model answer:
		msa := model.StudentAssignment{
			UserID:       10000,
			AssignmentID: assignment.ID,
		}
		db.Create(&msa)
		maf := model.Source{FileName: r.base + r.modelAnswerFileName, S3BucketName: "studentanswers"}
		db.Create(&maf)
		ma := model.Answer{
			SourceID:            model.NewNullInt64(maf.ID),
			QuestionID:          model.NewNullInt64(q.ID),
			SubmissionTime:      *parseTime("2018-09-30 12:42"),
			StudentAssignmentID: msa.ID,
		}
		db.Create(&ma)
		model.ExtractBlocksFromFile(r.base+r.modelAnswerFileName, "FFFFFF00", true, true, true, ma.ID)

		// Answer
		for _, ar := range r.answers {
			af := model.Source{FileName: r.base + ar.fileName, S3BucketName: "studentanswers"}
			db.Create(&af)

			sa := model.StudentAssignment{
				UserID:       ar.uid,
				AssignmentID: assignment.ID,
			}
			db.Create(&sa)

			a := model.Answer{
				SourceID:            model.NewNullInt64(af.ID),
				QuestionID:          model.NewNullInt64(q.ID),
				SubmissionTime:      *parseTime("2018-09-30 12:42"),
				StudentAssignmentID: sa.ID,
			}
			db.Create(&a)
			model.ExtractBlocksFromFile(r.base+ar.fileName, "FFFFFF00", true, true, true, a.ID)
		}
	}
}
