package tests

import (
	"database/sql"
	"flag"
	"fmt"
	"os"
	"path"
	"strconv"
	"strings"
	"sync"
	"testing"
	"time"

	"github.com/jinzhu/now"
	"github.com/nad2000/xlsx"

	log "github.com/Sirupsen/logrus"

	"extract-blocks/cmd"
	model "extract-blocks/model"
	"extract-blocks/s3"
	"extract-blocks/utils"

	_ "github.com/go-sql-driver/mysql"
	"github.com/jinzhu/gorm"
	_ "github.com/jinzhu/gorm/dialects/sqlite"
)

var testFileNames = []string{
	"demo.xlsx",
	"partial.xlsx",
	"Sample3_A2E1.xlsx",
	"Sample4_A2E1.xlsx",
	"test2.xlsx",
	"test.xlsx",
	"test_floats.xlsx",
	"Sample-poi-file.xlsx",
}
var testDb = path.Join(os.TempDir(), "extract-block-test.db")
var defaultURL = "sqlite3://" + testDb
var url string

func parseTime(str string) *time.Time {
	t := now.New(time.Now().UTC()).MustParse(str)
	return &t
}

func init() {
	wd, _ := os.Getwd()
	log.Info("Running tests in: ", wd)
	os.Setenv("TZ", "UTC")
	if _, err := os.Stat(testDb); !os.IsNotExist(err) {
		os.RemoveAll(testDb)
	}
	flag.StringVar(&url, "url", defaultURL, "Test database URL")
	flag.Parse()
	if url == defaultURL {
		databaseURL, ok := os.LookupEnv("DATABASE_URL")
		if ok {
			url = databaseURL
		}
	}
	log.Info("DATABASE URL: ", url)
	if strings.HasPrefix(url, "sqlite") {
		parts := strings.Split(url, "://")
		dbFileName := parts[len(parts)-1]
		if _, err := os.Stat(dbFileName); !os.IsNotExist(err) {
			os.RemoveAll(dbFileName)
		}
	}
}

func TestR1C1(t *testing.T) {
	for _, r := range []struct {
		x, y       int
		ID, expect string
	}{
		{0, 0, "A1", "R[+0]C[+0]"},
		{1, 1, "A1", "R[-1]C[-1]"},
		{3, 11, "A1", "R[-3]C[-11]"},
		{111, 11, "$A1", "R[-111]C[0]"}} {
		relID := model.RelativeCellAddress(r.x, r.y, r.ID)
		if relID != r.expect {
			t.Errorf("Expecte %q for %#v; got %q", r.expect, r, relID)
		}
	}
}

func TestRelativeFormulas(t *testing.T) {
	for _, r := range []struct {
		x, y       int
		ID, expect string
	}{
		{
			0, 0,
			"A1 / B11 - 67 * $ZZ$123 % ZA$233",
			"R[+0]C[+0] / R[+10]C[+1] - 67 * R[122]C[701] % R[232]C[+676]",
		},
	} {
		relID := model.RelativeFormula(r.x, r.y, r.ID)
		if relID != r.expect {
			t.Errorf("Expecte %q for %#v; got %q", r.expect, r, relID)
		}
	}
}

func TestDemoFile(t *testing.T) {
	deleteData()
	var wb model.Workbook
	var fileName = "demo.xlsx"

	db, _ := model.OpenDb(url)
	q := model.Question{
		QuestionType:      "ShortAnswer",
		QuestionSequence:  0,
		QuestionText:      "DUMMY",
		AnswerExplanation: sql.NullString{String: "DUMMY", Valid: true},
		MaxScore:          999.99,
	}
	db.FirstOrCreate(&q, &q)
	a := model.Answer{
		ShortAnswer:    fileName,
		SubmissionTime: *parseTime("2017-01-01 14:42"),
		QuestionID:     model.NewNullInt64(q.ID),
	}
	db.FirstOrCreate(&a, &a)

	t.Log("+++ Start extration")
	model.ExtractBlocksFromFile(fileName, "FFFFFF00", true, true, a.ID)

	result := db.First(&wb, "file_name = ?", fileName)
	if result.Error != nil {
		t.Error(result.Error)
		t.Fail()
	}

	if wb.FileName != "demo.xlsx" {
		t.Logf("Missing workbook: expected %q, got: %q", fileName, wb.FileName)
		t.Fail()
	}
	var count int
	db.Model(&model.Block{}).Count(&count)
	if expected := 16; count != expected {
		t.Errorf("Expected %d blocks, got: %d", expected, count)
	}
	db.Model(&model.Cell{}).Count(&count)
	if expected := 42; count != expected {
		t.Errorf("Expected %d cells, got: %d", expected, count)
	}

	db.DB().Exec("UPDATE StudentAnswers SET ShortAnswerText = 'FIRST-DEMO.xlsx'")
	db.DB().Exec("INSERT INTO Comments (CommentText) VALUES('DUMMY')")
	db.DB().Exec(`INSERT INTO BlockCommentMapping(ExcelBlockID, ExcelCommentID)
		SELECT ExcelBlockID, (SELECT CommentID FROM Comments LIMIT 1)
		FROM ExcelBlocks
		WHERE color = 'FFFFFF00'`)
	db.Close()

	cmd.RootCmd.SetArgs([]string{
		"run", "-U", url, "-t", "-f", "demo.xlsx"})
	cmd.Execute()

	db, _ = model.OpenDb(url)
	defer db.Close()

	db.Model(&model.BlockCommentMapping{}).Count(&count)
	if expected := 6; count != expected {
		t.Errorf("Expected %d block -> comment mapping entries, got: %d", expected, count)
	}
	db.Model(&model.Block{}).Count(&count)
	if expected := 32; count != expected {
		t.Errorf("Expected %d blocks, got: %d", expected, count)
	}
	db.Model(&model.Cell{}).Count(&count)
	if expected := 84; count != expected {
		t.Errorf("Expected %d cells, got: %d", expected, count)
	}

	// With references
	refWB := model.Workbook{IsReference: true}
	db.FirstOrCreate(&refWB, refWB)
	refWS := model.Worksheet{IsReference: true, WorkbookID: refWB.ID}
	db.FirstOrCreate(&refWS, refWS)
	refBlock := model.Block{
		WorksheetID: refWS.ID,
		IsReference: true,
		TRow:        14,
		BRow:        32,
		LCol:        11,
		RCol:        16,
		Range:       "L15:Q33",
	}
	_ = refBlock
	db.FirstOrCreate(&refBlock, refBlock)
	q = model.Question{
		QuestionType:      "ShortAnswer",
		QuestionSequence:  0,
		QuestionText:      "DUMMY",
		AnswerExplanation: sql.NullString{String: "DUMMY", Valid: true},
		MaxScore:          999.99,
		IsFormatting:      true,
		ReferenceID:       model.NewNullInt64(refWB.ID),
	}
	db.FirstOrCreate(&q, q)
	a = model.Answer{
		ShortAnswer:    fileName,
		SubmissionTime: *parseTime("2017-01-01 14:42"),
		QuestionID:     model.NewNullInt64(q.ID),
	}
	db.FirstOrCreate(&a, a)

	model.ExtractBlocksFromFile(fileName, "", true, true, a.ID)
}

func TestFormattingImport(t *testing.T) {
	db, _ = model.OpenDb(url)
	defer db.Close()
	deleteData()
	var fileName = "Formatting Test File.xlsx"

	refWB := model.Workbook{IsReference: true}
	db.FirstOrCreate(&refWB, refWB)
	refWS := model.Worksheet{IsReference: true, WorkbookID: refWB.ID, Idx: 1}
	db.FirstOrCreate(&refWS, refWS)
	refBlock := model.Block{
		WorksheetID: refWS.ID,
		IsReference: true,
		TRow:        1,
		BRow:        13,
		LCol:        1,
		RCol:        10,
		Range:       "B2:K14",
	}
	db.FirstOrCreate(&refBlock, refBlock)
	q := model.Question{
		QuestionType:      "ShortAnswer",
		QuestionSequence:  0,
		QuestionText:      "DUMMY",
		AnswerExplanation: sql.NullString{String: "DUMMY", Valid: true},
		MaxScore:          999.99,
		IsFormatting:      true,
		ReferenceID:       model.NewNullInt64(refWB.ID),
	}
	db.FirstOrCreate(&q, q)
	a := model.Answer{
		ShortAnswer:    fileName,
		SubmissionTime: *parseTime("2017-01-01 14:42"),
		QuestionID:     model.NewNullInt64(q.ID),
	}
	db.FirstOrCreate(&a, a)

	model.ExtractBlocksFromFile(fileName, "FFFFFF00", true, true, a.ID)
}

func deleteData() {

	if db == nil || db.DB() == nil {
		db, _ = model.OpenDb(url)
		defer db.Close()
	}
	for _, m := range []interface{}{
		&model.XLQTransformation{},
		&model.StudentAssignment{},
		&model.Cell{},
		&model.Block{},
		&model.Chart{},
		&model.BlockCommentMapping{},
		&model.AnswerComment{},
		&model.QuestionExcelData{},
		&model.QuestionAssignment{},
		&model.Assignment{},
		&model.Worksheet{},
		&model.Workbook{},
		&model.Question{},
		&model.Answer{},
		&model.Source{},
		&model.Comment{},
		&model.ConditionalFormatting{},
		&model.Filter{},
		&model.Sorting{},
		&model.PivotTable{},
		&model.DataSource{},
		&model.User{},
		&model.Alignment{},
		&model.Border{},
	} {
		err := db.Delete(m).Error
		if err != nil {
			fmt.Println("ERROR: ", err)
		}
	}
}

func createTestDB() *gorm.DB {
	// this is a makeshift solution to deal with 'unclodes' DBs...
	if db != nil {
		db.Close()
	}
	if cmd.Db != nil {
		cmd.Db.Close()
	}
	if strings.HasPrefix(url, "sqlite") {
		if _, err := os.Stat(testDb); !os.IsNotExist(err) {
			os.RemoveAll(testDb)
		}
	}
	db, _ = model.OpenDb(url)
	cmd.Db = db

	deleteData()
	for _, uid := range []int{4951, 4952, 4953} {
		db.Create(&model.User{ID: uid})
	}
	assignment := model.Assignment{
		Title: "Testing...",
		State: "READY_FOR_GRADING",
	}
	db.Create(&assignment)
	sa := model.StudentAssignment{
		UserID:       4951,
		AssignmentID: assignment.ID,
	}
	db.Create(&sa)

	for _, fn := range testFileNames {
		f := model.Source{
			FileName:     fn,
			S3BucketName: "studentanswers",
			S3Key:        fn,
		}

		db.Create(&f)
		q := model.Question{
			SourceID:         model.NewNullInt64(f.ID),
			QuestionType:     model.QuestionType("FileUpload"),
			QuestionSequence: 123,
			QuestionText:     "QuestionText...",
			MaxScore:         9999.99,
			AuthorUserID:     123456789,
			WasCompared:      true,
			IsFormatting:     true,
		}
		db.Create(&q)
		db.Create(&model.QuestionAssignment{
			QuestionID:   q.ID,
			AssignmentID: assignment.ID,
		})
		db.Create(&model.Answer{
			SourceID:            model.NewNullInt64(f.ID),
			QuestionID:          model.NewNullInt64(q.ID),
			StudentAssignmentID: sa.ID,
			SubmissionTime:      *parseTime("2017-01-01 14:42"),
		})
		if fn == "Sample-poi-file.xlsx" {
			rwb := model.Workbook{IsReference: true}
			db.Create(&rwb)
			q.ReferenceID = model.NewNullInt64(rwb.ID)
			db.Save(&q)
			rws := model.Worksheet{IsReference: true, WorkbookID: rwb.ID}
			db.Create(&rws)
			db.Create(&model.Block{
				WorksheetID: rws.ID,
				Range:       "A1:M98",
				TRow:        0,
				LCol:        0,
				BRow:        97,
				RCol:        12,
				IsReference: true,
			})
			t, _ := time.Parse(time.UnixDate, "Thu Dec 20 12:06:10 UTC 2042")
			db.Create(&model.XLQTransformation{
				CellReference: "AT9013",
				UserID:        4951,
				TimeStamp:     t,
				QuestionID:    q.ID,
				// SourceID:      f.ID,
			})
		}
	}

	ignore := model.Source{FileName: "ignore.abc"}
	db.Create(&ignore)
	db.Create(&model.Answer{SourceID: model.NewNullInt64(ignore.ID), SubmissionTime: *parseTime("2017-01-01 14:42")})

	var fileID int
	db.DB().QueryRow("SELECT MAX(FileID)+1 AS FileID FROM FileSources").Scan(&fileID)
	for i := fileID + 1; i < fileID+10; i++ {
		fn := "question" + strconv.Itoa(i)
		if i < fileID+4 {
			fn += ".xlsx"
		} else {
			fn += ".ignore"
		}
		f := model.Source{
			ID:           i,
			FileName:     fn,
			S3BucketName: "studentanswers",
			S3Key:        fn,
		}
		db.Create(&f)
		db.Create(&model.Question{
			SourceID:         model.NewNullInt64(i),
			QuestionType:     model.QuestionType("FileUpload"),
			QuestionSequence: 123,
			QuestionText:     "QuestionText...",
			MaxScore:         9999.99,
			AuthorUserID:     123456789,
			WasCompared:      true,
			IsFormatting:     true,
		})

	}

	return db
}

var db *gorm.DB

func testQuestionsToProcess(t *testing.T) {

	var rows []model.Question
	db.Find(&rows)
	for _, r := range rows {
		var s model.Source
		db.Model(&r).Related(&s, "FileID")
	}
	if expected, count := 17, len(rows); expected != count {
		t.Errorf("Expected %d question rows, got %d", expected, count)
	}

	questions, err := model.QuestionsToProcess()
	if err != nil {
		t.Error(err)
		t.Fail()
	}
	for _, r := range questions {
		var s model.Source
		db.Model(&r).Related(&s, "FileID")
		if !strings.HasSuffix(s.FileName, ".xlsx") {
			t.Errorf("Wrong extension: %q, expected: '.xlsx'", s.FileName)
		}
	}
	if expected, count := 11, len(questions); expected != count {
		t.Errorf("Expected %d rows, got %d", expected, count)
	}

}

func testImportFile(t *testing.T) {
	q := model.Question{
		QuestionType:     model.QuestionType("FileUpload"),
		QuestionSequence: 123,
		QuestionText:     "Test Import Question...",
		MaxScore:         9999.99,
		AuthorUserID:     123456789,
		WasCompared:      true,
		Source:           model.Source{FileName: "question.xlsx"},
		IsFormatting:     true,
	}
	db.Create(&q)
	q.ImportFile("question.xlsx", "FFFFFF00", true)

	var (
		count int
		ws    model.Worksheet
	)
	db.First(&ws, "workbook_id = ?", q.ReferenceID)
	db.Model(&model.Block{}).Where("is_reference = ? AND worksheet_id = ?", true, ws.ID).Count(&count)
	if expected := 8; count != expected {
		t.Errorf("Expected %d blocks, got: %d", expected, count)
	}

}

func testHandleNotcolored(t *testing.T) {

	var (
		q          model.Question
		assignment model.Assignment
		f          model.Source
		a          model.Answer
	)
	q = model.Question{
		QuestionType:     model.QuestionType("FileUpload"),
		QuestionSequence: 99,
		QuestionText:     "Test handle answers without the colorcodes...",
		MaxScore:         9999.99,
		AuthorUserID:     123456789,
		WasCompared:      true,
	}
	db.Create(&q)
	q.ImportFile("question.xlsx", "FFFFFF00", true)
	assignment = model.Assignment{
		Title: "Test handle answers without the colorcodes...",
		State: "READY_FOR_GRADING",
	}
	db.Create(&assignment)
	db.Create(&model.QuestionAssignment{
		QuestionID:   q.ID,
		AssignmentID: assignment.ID,
	})
	for _, fn := range []string{"answer.xlsx", "answer.nocolor.xlsx"} {
		f := model.Source{
			FileName:     fn,
			S3BucketName: "studentanswers",
			S3Key:        fn,
		}
		db.Create(&f)
		a := model.Answer{
			SourceID:       model.NewNullInt64(f.ID),
			QuestionID:     model.NewNullInt64(q.ID),
			SubmissionTime: *parseTime("2018-09-14 14:42"),
		}
		db.Create(&a)
		model.ExtractBlocksFromFile(fn, "FFFFFF00", true, true, a.ID)
	}

	// var count int
	// db.Model(&model.Block{}).Where("is_reference = ?", true).Count(&count)
	// if expected := 8; count != expected {
	// 	t.Errorf("Expected %d blocks, got: %d", expected, count)
	// }

	q = model.Question{
		QuestionType:     model.QuestionType("FileUpload"),
		QuestionSequence: 99,
		QuestionText:     "Test handle answers without the colorcodes #2...",
		MaxScore:         9999.99,
		AuthorUserID:     123456789,
		WasCompared:      true,
		IsFormatting:     true,
	}
	db.Create(&q)
	q.ImportFile("Q1 Question different color.xlsx", "FFFFFF00", true)
	assignment = model.Assignment{
		Title: "Test handle answers without the colorcodes #2...",
		State: "READY_FOR_GRADING",
	}
	db.Create(&assignment)
	db.Create(&model.QuestionAssignment{
		QuestionID:   q.ID,
		AssignmentID: assignment.ID,
	})
	f = model.Source{
		FileName:     "Q1 Solution different color stud4.xlsx",
		S3BucketName: "studentanswers",
	}
	db.Create(&f)
	a = model.Answer{
		SourceID:       model.NewNullInt64(f.ID),
		QuestionID:     model.NewNullInt64(q.ID),
		SubmissionTime: *parseTime("2018-09-30 12:42"),
	}
	db.Create(&a)
	model.ExtractBlocksFromFile(f.FileName, "FFFFFF00", true, true, a.ID)
}

func testHandleNotcoloredQ3(t *testing.T) {
	var (
		q          model.Question
		assignment model.Assignment
		a          model.Answer
	)
	q = model.Question{
		QuestionType:     model.QuestionType("FileUpload"),
		QuestionSequence: 99,
		QuestionText:     "Test handle answers without the colorcodes #3...",
		MaxScore:         9999.99,
		AuthorUserID:     123456789,
		WasCompared:      true,
		Source:           model.Source{FileName: "Q3 Compounding1.xlsx"},
		IsFormatting:     true,
	}
	db.Create(&q)
	q.ImportFile("Q3 Compounding1.xlsx", "FFFFFF00", true)
	assignment = model.Assignment{
		Title: "Test handle answers without the colorcodes #3...",
		State: "READY_FOR_GRADING",
	}
	db.Create(&assignment)
	db.Create(&model.QuestionAssignment{
		QuestionID:   q.ID,
		AssignmentID: assignment.ID,
	})
	for _, r := range []struct {
		fn                     string
		blockCount, emptyCount int
	}{
		{"Answer stud 1 Q3 Compounding1.xlsx", 1, 99},
		{"Answer stud 2 Q3 Compounding1.xlsx", 2, 98},
		{"Answer stud 3 Q3 Compounding1.xlsx", 3, 97},
	} {
		a = model.Answer{
			Source:         model.Source{FileName: r.fn, S3BucketName: "studentanswers"},
			QuestionID:     model.NewNullInt64(q.ID),
			SubmissionTime: *parseTime("2018-09-30 12:42"),
		}
		db.Create(&a)
		wb, err := model.ExtractBlocksFromFile(r.fn, "FFFFFF00", true, true, a.ID)
		if err != nil {
			t.Error(err)
		}
		var ws model.Worksheet
		db.First(&ws, "workbook_id = ?", wb.ID)
		var count int
		db.Model(&model.Block{}).Where("worksheet_id = ? AND BlockFormula = ''", ws.ID).Count(&count)
		if expected := r.emptyCount; count != expected {
			t.Errorf("Empty block count: %d, expected: %d", count, expected)
		}
		db.Model(&model.Block{}).Where("worksheet_id = ? AND BlockFormula != ''", ws.ID).Count(&count)
		if expected := r.blockCount; count != expected {
			t.Errorf("Block count: %d, expected: %d", count, expected)
		}
	}
}

func testRowsToProcess(t *testing.T) {

	rows, _ := model.RowsToProcess()

	for _, r := range rows {
		if !strings.HasSuffix(r.FileName, ".xlsx") {
			t.Errorf("Expected only .xlsx extensions, got %q", r.FileName)
		}
	}
	if len(rows) != len(testFileNames) {
		t.Errorf("Expected %d rows, got %d", len(testFileNames), len(rows))
	}

}

type testManager struct{}

func (m testManager) Download(sourceName, s3BucketName, s3Key, dest string) (string, error) {
	return sourceName, nil
}

func (m testManager) Upload(sourceName, s3BucketName, s3Key string) (string, error) {
	return sourceName, nil
}

func testHandleAnswers(t *testing.T) {

	var countBefore, countAfter int
	db.Table("StudentAnswers").Where("was_xl_processed = ?", 0).Count(&countBefore)
	tm := testManager{}
	cmd.HandleAnswers(&tm)
	db.Table("StudentAnswers").Where("was_xl_processed = ?", 0).Count(&countAfter)
	if countBefore <= countAfter {
		t.Errorf(
			"Expeced that the number of answers to be processed dorps, got %d before, and %d after.",
			countBefore, countAfter)
	}
	var ws model.Worksheet
	db.Where("workbook_file_name = ?", "Sample-poi-file.xlsx").First(&ws)
	// if !ws.IsPlagiarised.Bool {
	if !ws.IsPlagiarised {
		t.Errorf("Expected that %#v will get marked as plagiarised.", ws)
	}
	// Auto-commenting
	var cells []model.Cell
	db.Find(&cells)
	for i, c := range cells {
		if i%5 <= 1 {
			db.Create(&model.AutoEvaluation{IsValueCorrect: i%2 == 0, CellID: c.ID})
		}
	}
	model.AutoCommentAnswerCells(12345)
}

func TestProcessing(t *testing.T) {

	db = createTestDB()
	defer db.Close()

	t.Run("QuestionsToProcess", testQuestionsToProcess)
	t.Run("RowsToProcess", testRowsToProcess)
	t.Run("FindBlocksInside", testFindBlocksInside)
	t.Run("ImportFile", testImportFile)
	t.Run("HandleAnswers", testHandleAnswers)
	t.Run("HandleNotcolored", testHandleNotcolored)
	t.Run("HandleNotcoloredQ3", testHandleNotcoloredQ3)
	t.Run("S3Downloading", testS3Downloading)
	t.Run("S3Uploading", testS3Uploading)
	t.Run("Questions", testQuestions)
	t.Run("HandleQuestions", testHandleQuestions)
	t.Run("CurruptedFiles", testCorruptedFiles)
	t.Run("ImportQuestionFile", testImportQuestionFile)
	t.Run("CWA175", testCWA175)
	t.Run("ImportWorksheets", testImportWorksheets)
	t.Run("POI", testPOI)
	t.Run("FullCycle", testFullCycle)
}

func testFullCycle(t *testing.T) {

	assignment := model.Assignment{
		Title: "Full Cycel Testing...",
		State: "READY_FOR_GRADING",
	}
	db.Create(&assignment)
	for _, r := range []struct {
		questionFileName, anserFileName, cr, ts string
		uid                                     int
		isPlagiarised                           bool
	}{
		{"Question_Stud1_4951.xlsx", "Answer_stud1_NOT_PLAGIARISED.xlsx", "KE4423", "2018-12-27 19:18:05", 4951, false},
		{"Question_Stud2_4952.xlsx", "Answer_stud2_PLAGIARISED.xlsx", "LI7010", "2018-12-27 19:51:29", 4952, true},
		{"Question_Stud3_4953.xlsx", "Answer_stud3_PLAGIARISED.xlsx", "AIP5821", "2018-12-28 07:59:21", 4953, true},
		{"Question_Stud1_4951.xlsx", "demo.xlsx", "", "", -1, true}, // "missing download"
	} {

		qf := model.Source{
			FileName:     r.questionFileName,
			S3BucketName: "studentanswers",
			S3Key:        r.questionFileName,
		}
		db.Create(&qf)
		q := model.Question{
			SourceID:     model.NewNullInt64(qf.ID),
			QuestionType: model.QuestionType("FileUpload"),
			QuestionText: r.questionFileName,
			MaxScore:     1010.88,
			AuthorUserID: 123456789,
			WasCompared:  false,
			IsFormatting: true,
		}
		if err := db.Create(&q).Error; err != nil {
			t.Error(err)
		}
		q.ImportFile(r.questionFileName, "FFFFFF00", true)
		sa := model.StudentAssignment{
			UserID:       r.uid,
			AssignmentID: assignment.ID,
		}
		db.Create(&sa)
		// answer
		af := model.Source{FileName: r.anserFileName, S3BucketName: "studentanswers"}
		db.Create(&af)
		a := model.Answer{
			SourceID:            model.NewNullInt64(af.ID),
			QuestionID:          model.NewNullInt64(q.ID),
			SubmissionTime:      *parseTime("2018-09-30 12:42"),
			StudentAssignmentID: sa.ID,
		}
		db.Create(&a)
		if r.uid > 0 {
			db.Create(&model.XLQTransformation{
				CellReference: r.cr,
				UserID:        r.uid,
				TimeStamp:     *parseTime(r.ts),
				QuestionID:    q.ID,
			})
			// Create extar entries for  non-pagiarised examples
			if !r.isPlagiarised {
				for i := 1; i < 10; i++ {
					db.Create(&model.XLQTransformation{
						CellReference: "A" + strconv.Itoa(i),
						UserID:        r.uid,
						TimeStamp:     *parseTime(r.ts),
						QuestionID:    q.ID,
					})
				}
			}
		}
		model.ExtractBlocksFromFile(r.anserFileName, "FFFFFF00", true, true, a.ID)
		// Test if is marked plagiarised:
		{
			var ws model.Worksheet
			db.Where("workbook_file_name = ?", r.anserFileName).First(&ws)
			if ws.IsPlagiarised != r.isPlagiarised {
				t.Errorf("Exected that %#v will get marked as plagiarised.", ws)
			}
		}
	}

	// Auto-commenting
	var cells []model.Cell
	db.Find(&cells)
	for i, c := range cells {
		switch i % 3 {
		case 0:
			db.Create(&model.AutoEvaluation{IsValueCorrect: false, CellID: c.ID})
		case 1:
			db.Create(&model.AutoEvaluation{IsValueCorrect: true, CellID: c.ID})
		}
	}
	var countBefore int
	if err := db.Model(&model.AutoEvaluation{}).Count(&countBefore).Error; err != nil {
		t.Error(err)
	}
	model.AutoCommentAnswerCells(12345)

	var countAfter int
	if err := db.Model(&model.AutoEvaluation{}).Count(&countAfter).Error; err != nil {
		t.Error(err)
	}
	if countAfter != countBefore {
		t.Errorf(
			"Exected unchanged rowcount of AutoEvaluation table. Expected: %d, got: %d",
			countBefore, countAfter)
	}
}

func testPOI(t *testing.T) {

	fileName := "POI Q1.xlsx"
	q := model.Question{
		Source: model.Source{
			FileName:     fileName,
			S3BucketName: "studentanswers",
			S3Key:        fileName,
		},
		QuestionType: model.QuestionType("FileUpload"),
		QuestionText: fileName,
		MaxScore:     1010.88,
		AuthorUserID: 123456789,
		WasCompared:  true,
		IsFormatting: true,
	}

	if err := db.Create(&q).Error; err != nil {
		t.Error(err)
	}
	assignment := model.Assignment{
		Title: "Testing...",
		State: "READY_FOR_GRADING",
	}
	db.Create(&assignment)
	sa := model.StudentAssignment{
		UserID:       4951,
		AssignmentID: assignment.ID,
	}
	db.Create(&sa)
	q.ImportFile(fileName, "FFFFFF00", true)

	// Answer
	fileName = "POI Q1_4951_null.xlsx"
	f := model.Source{FileName: fileName, S3BucketName: "studentanswers"}
	db.Create(&f)
	a := model.Answer{
		Source:              model.Source{FileName: fileName, S3BucketName: "studentanswers"},
		QuestionID:          model.NewNullInt64(q.ID),
		SubmissionTime:      *parseTime("2018-09-30 12:42"),
		StudentAssignmentID: sa.ID,
	}
	db.Create(&a)
	ts, _ := time.Parse(time.UnixDate, "Thu Dec 27 19:18:05 UTC 2088")
	db.Create(&model.XLQTransformation{
		CellReference: "KE4423",
		UserID:        4951,
		TimeStamp:     ts,
		QuestionID:    q.ID,
	})
	model.ExtractBlocksFromFile(fileName, "FFFFFF00", true, true, a.ID)

	var ws model.Worksheet
	db.Where("workbook_file_name = ?", fileName).First(&ws)
	if !ws.IsPlagiarised {
		t.Errorf("Expected that %#v will get marked as plagiarised.", ws)
	}
}

func testImportWorksheets(t *testing.T) {
	for _, fn := range []string{
		"CF ALL TYPES.xlsx",
		"Sorting ALL TYPES.xlsx",
		"Sorting Horizontal.xlsx",
		"Filter ALL TYPES.xlsx",
		"Salesman filter.xlsx",
		"Pivot 2.xlsx",
		"Multi text custom filter.xlsx",
	} {
		wb := model.Workbook{FileName: fn}
		db.Create(&wb)
		wb.ImportWorksheets(fn)
	}
}

func testFindBlocksInside(t *testing.T) {
	/* Expected Block and Formula:
	1. D8:D8 - B8*C8
	2. D9:D9 - PRODUCT(B9,C9)
	3. D10:D10 - PRODUCT(B10:C10)
	4. D11:D20 - C11*B11
	5. G9:G9 - SUM(B8,B9,B10,B11,B12,B13,B14,B15,B16,B17,B18,B19,B20)
	6. G10:G10 - SUM(D8:D20)
	7. G11:G11 - G10/G9
	*/
	source := model.Source{S3Key: "KEY", FileName: "test.xlsx"}
	ws := model.Worksheet{
		Workbook: model.Workbook{
			FileName: "find_blocks_inside_a_block.xlsx",
			Answer: model.Answer{
				Assignment: model.Assignment{Title: "TEST ASSIGNMENT", AssignmentSequence: 888},
				Marks:      98.7654,
				Source:     source,
				Question: model.Question{
					QuestionType: "FileUpload", Source: source, MaxScore: 98.76453,
					IsFormatting: true,
				},
			},
		},
	}
	db.Create(&ws)

	file, _ := xlsx.OpenFile("Q1 Solution different color stud4.xlsx")
	sheet := file.Sheets[0]
	ws.FindBlocksInside(sheet, model.Block{
		Range:       "D8:D20",
		LCol:        3,
		TRow:        7,
		RCol:        3,
		BRow:        19,
		IsReference: true,
	}, true)
	var blocks []model.Block
	db.Model(&ws).Related(&blocks)
	if expected, got := 4, len(blocks); expected != got {
		for _, b := range blocks {
			t.Log(b)
		}
		t.Errorf("Expected %d blocks, got: %d", expected, got)
	}
	// t.Log(ws)
	block := model.Block{
		Worksheet:   ws,
		Range:       "G9:G11",
		LCol:        6,
		TRow:        8,
		RCol:        6,
		BRow:        10,
		IsReference: true,
	}
	db.Create(&block)
	ws.FindBlocksInside(sheet, block, true)
	db.Model(&ws).Related(&blocks)
	if expected, got := 8, len(blocks); expected != got {
		for _, b := range blocks {
			t.Log(b)
		}
		t.Errorf("Got %d blocks, expected: %d", got, expected)
	}
}

func testCWA175(t *testing.T) {
	assignment := model.Assignment{Title: "TEST ASSIGNMENT", AssignmentSequence: 888}
	db.Create(&assignment)
	wb := model.Workbook{
		IsReference: true,
		FileName:    "reference-workbook.xlsx",
	}
	db.Create(&wb)
	ws := model.Worksheet{
		IsReference: true,
		Workbook:    wb,
		OrderNum:    0,
	}
	db.Create(&ws)
	q := model.Question{
		ReferenceID:  model.NewNullInt64(wb.ID),
		IsFormatting: true,
	}
	db.Create(&q)
	rb := model.Block{
		Range:       "C2:C8",
		TRow:        1,
		LCol:        2,
		BRow:        7,
		RCol:        2,
		IsReference: true,
		Worksheet:   ws,
	}
	db.Create(&rb)
	for _, r := range []struct {
		fn       string
		expected int
	}{
		{"CWA175-student1.xlsx", 58},
		{"CWA175-student2.xlsx", 21},
		{"CWA175-student3.xlsx", 25},
	} {
		answer := model.Answer{
			SubmissionTime: *parseTime("2017-01-01 14:42"),
			Assignment:     assignment,
			QuestionID:     model.NewNullInt64(q.ID),
			Marks:          98.7654,
			Source: model.Source{
				FileName:     r.fn,
				S3BucketName: "studentanswers",
				S3Key:        r.fn,
			},
		}
		db.Create(&answer)
		model.ExtractBlocksFromFile(r.fn, "", true, true, answer.ID)

		var wb model.Workbook

		if err := db.Preload("Worksheets").Preload("Worksheets.Blocks").First(&wb, "StudentAnswerID = ?", answer.ID).Error; err != nil {
			t.Error(err)
		}
		if wb.Worksheets == nil || len(wb.Worksheets) < 1 {
			t.Error("Missing worksheets in the workbook", wb)
			continue
		}
		blocks := wb.Worksheets[0].Blocks
		if blocks == nil {
			t.Error("Missing blocks in the workbook", wb)
			continue
		}
		if count := len(wb.Worksheets[0].Blocks); count != r.expected {
			t.Errorf("Expected %d blocks, got: %d", r.expected, count)
		}
	}
}

func testHandleQuestions(t *testing.T) {

	var fileID int
	db.DB().QueryRow("SELECT MAX(FileID)+1 AS LastFileID FROM FileSources").Scan(&fileID)
	f := model.Source{
		ID:           fileID,
		FileName:     "merged.xlsx",
		S3BucketName: "studentanswers",
		S3Key:        "merged.xlsx",
	}
	result := db.Create(&f)
	if result.Error != nil {
		t.Error(result.Error)
	}
	for _, qt := range []string{
		"Question with merged cells #1",
		"Question with merged cells #2 (duplicate file name)",
		"Question with merged cells #3 (duplicate file name)",
	} {
		result = db.Create(&model.Question{
			SourceID:     model.NewNullInt64(fileID),
			QuestionType: model.QuestionType("FileUpload"),
			QuestionText: qt,
			MaxScore:     8888.88,
			AuthorUserID: 123456789,
			WasCompared:  true,
			IsFormatting: true,
		})
		if result.Error != nil {
			t.Error(result.Error)
		}
	}

	tm := testManager{}
	cmd.HandleQuestions(&tm)

	var count int
	if err := db.Where("file_name = ?", "merged.xlsx").Model(&model.Workbook{}).Count(&count).Error; err != nil {
		t.Error(err)
	}
	if expected := 3; count != expected {
		t.Errorf("Expected %d question reference workbooks, got: %d", expected, count)
	}

	if err := db.DB().QueryRow(`
		SELECT COUNT(DISTINCT reference_id) AS ReferenceCount
		FROM Questions
		WHERE QuestionText LIKE 'Question with merged cells #%'`).Scan(&count); err != nil {
		t.Error(err)
	}
	if expected := 3; count != expected {
		t.Errorf("Expected %d questions, got: %d", expected, count)
	}

	var blockCount int
	if err := db.DB().QueryRow(`
		SELECT count(*) AS RowCount, count(distinct b.BlockCellRange) AS BlockCount
		FROM Questions AS q JOIN WorkSheets AS s ON s.workbook_id = q.reference_id
		JOIN ExcelBlocks AS b ON b.worksheet_id = s.id
		WHERE QuestionText LIKE 'Question with merged cells #%'`).Scan(&count, &blockCount); err != nil {
		t.Error(err)
	}
	if expected := 24; count != expected {
		t.Errorf("Expected %d rows, got: %d", expected, count)
	}
	if expected := 8; blockCount != expected {
		t.Errorf("Expected %d different blocks, got: %d", expected, blockCount)
	}
}

func testImportQuestionFile(t *testing.T) {

	var fileID int
	fileName := "Q1 Question different color.xlsx"
	db.DB().QueryRow("SELECT MAX(FileID)+1 AS LastFileID FROM FileSources").Scan(&fileID)
	f := model.Source{
		ID:           fileID,
		FileName:     fileName,
		S3BucketName: "studentanswers",
		S3Key:        fileName,
	}
	result := db.Create(&f)
	if result.Error != nil {
		t.Error(result.Error)
	}
	q := model.Question{
		Source:       f,
		QuestionType: model.QuestionType("FileUpload"),
		QuestionText: fileName,
		MaxScore:     1010.88,
		AuthorUserID: 123456789,
		WasCompared:  true,
		IsFormatting: true,
	}
	result = db.Create(&q)
	if result.Error != nil {
		t.Error(result.Error)
	}
	q.ImportFile(fileName, "FFFFFF00", true)
}

func testCorruptedFiles(t *testing.T) {
	q := model.Question{
		SourceID:         sql.NullInt64{},
		QuestionType:     model.QuestionType("FileUpload"),
		QuestionSequence: 77,
		QuestionText:     "Test handle answers without the colorcodes...",
		MaxScore:         7777.77,
		AuthorUserID:     987654321,
		WasCompared:      true,
		IsFormatting:     true,
	}
	db.Create(&q)
	assignment := model.Assignment{
		Title: "Test handle answers without the colorcodes...",
		State: "READY_FOR_GRADING",
	}
	db.Create(&assignment)
	db.Create(&model.QuestionAssignment{
		QuestionID:   q.ID,
		AssignmentID: assignment.ID,
	})
	for _, fn := range []string{"corrupt1.xlsx", "corrupt2.xlsx", "demo.xlsx"} {
		f := model.Source{
			FileName:     fn,
			S3BucketName: "studentanswers",
			S3Key:        fn,
		}
		db.Create(&f)
		a := model.Answer{
			SourceID:       model.NewNullInt64(f.ID),
			QuestionID:     model.NewNullInt64(q.ID),
			SubmissionTime: *parseTime("2018-09-14 14:42"),
		}
		db.Create(&a)
	}
	// model.ExtractBlocksFromFile(fn, "FFFFFF00", true, true, false, a.ID)

	// var count int
	// db.Model(&model.Block{}).Where("is_reference = ?", true).Count(&count)
	// if expected := 8; count != expected {
	// 	t.Errorf("Expected %d blocks, got: %d", expected, count)
	// }
	var countBefore, countAfter int
	db.Table("StudentAnswers").Where("was_xl_processed = ?", 0).Count(&countBefore)
	tm := testManager{}
	cmd.HandleAnswers(&tm)
	db.Table("StudentAnswers").Where("was_xl_processed = ?", 0).Count(&countAfter)
	if countBefore <= countAfter {
		t.Errorf(
			"Expeced that the number of answers to be processed dorps, got %d before, and %d afater.",
			countBefore, countAfter)
	}
	var answerCount int
	if err := db.DB().QueryRow(`
			SELECT COUNT(*) FROM StudentAnswers AS a
			WHERE a.FileID IS NULL
			  AND a.was_xl_processed = 0`).Scan(&answerCount); err != nil {
		t.Error(err)
	}
	if answerCount != 2 {
		t.Errorf("Expecte 2 answers rejected: %d", answerCount)
	}

}

func testQuestions(t *testing.T) {

	if testing.Short() {
		t.Skip("Skipping 'questions' testing...")
	}

	cmd.RootCmd.SetArgs([]string{"questions", "-U", url})
	cmd.Execute()

	if db == nil || db.DB() == nil {
		db, _ := model.OpenDb(url)
		defer db.Close()
	}

	var count int
	db.Model(&model.QuestionExcelData{}).Count(&count)
	if count != 1690 {
		t.Errorf("Expected 1690 blocks, got: %d", count)
	}
}

// Random number state.
// We generate random temporary file names so that there's a good
// chance the file doesn't exist yet - keeps the number of tries in
// TempFile to a minimum.
var rand uint32
var randmu sync.Mutex

func reseed() uint32 {
	return uint32(time.Now().UnixNano() + int64(os.Getpid()))
}

func nextRandomName() string {
	randmu.Lock()
	r := rand
	if r == 0 {
		r = reseed()
	}
	r = r*1664525 + 1013904223 // constants from Numerical Recipes
	rand = r
	randmu.Unlock()
	return strconv.Itoa(int(1e9 + r%1e9))[1:]
}

func testS3Downloading(t *testing.T) {

	if testing.Short() {
		t.Skip("Skipping S3 downloading testing...")
	}

	m := s3.NewManager("us-east-1", "rad")
	destName := path.Join(os.TempDir(), nextRandomName()+".xlsx")
	_, err := m.Download("test.xlsx", "studentanswers", "test.xlsx", destName)
	if err != nil {
		t.Error(err)
	}
	stat, err := os.Stat(destName)
	if os.IsNotExist(err) {
		t.Errorf("Expected to have file %q", destName)
	}
	if stat.Size() < 1000 {
		t.Errorf("Expected at least 5kB size file, got: %d bytes", stat.Size())
	}
}

func testS3Uploading(t *testing.T) {

	if testing.Short() {
		t.Skip("Skipping S3 uploading testing...")
	}

	m := s3.NewManager("us-east-1", "rad")
	location, err := m.Upload("upload.test.txt", "studentanswers", "upload.test.txt")
	if err != nil {
		t.Error(err)
	}
	t.Log("*** Uploaded to:", location)
}

func TestCommenting(t *testing.T) {

	db = createTestDB()
	defer db.Close()

	db.Exec(`
		INSERT INTO WorkBooks(file_name, StudentAnswerID)
		SELECT FileName, StudentAnswerID
		FROM FileSources NATURAL JOIN StudentAnswers`)

	db.Exec(`
		INSERT INTO WorkSheets (
			workbook_id,
			name,
			workbook_file_name,
			StudentAnswerID
		)
		SELECT wb.id, wsn.Name, FileName, sa.StudentAnswerID
		FROM FileSources NATURAL JOIN StudentAnswers AS sa
		JOIN WorkBooks AS wb ON wb.StudentAnswerID = sa.StudentAnswerID,
		(
			SELECT 'Sheet1' AS Name
			UNION SELECT 'Sheet2'
		) AS wsn`)
	db.Exec(`
		INSERT INTO ExcelBlocks(worksheet_id, BlockCellRange)
		SELECT id, r.v
		FROM WorkSheets AS s LEFT JOIN ExcelBlocks AS b
		ON b.worksheet_id = s.id,
		(
			SELECT 'A1' AS v
			UNION SELECT 'C3'
			UNION SELECT 'D1:F2'
			UNION SELECT 'C1:F2'
			UNION SELECT 'C12:E21'
			UNION SELECT 'C13:D16'
		) AS r
		WHERE b.ExcelBlockID IS NULL`)
	db.Create(&model.Assignment{Title: "ASSIGNMENT #1", State: "GRADED"})
	db.Create(&model.Assignment{Title: "ASSIGNMENT #2"})
	db.Exec(`
		UPDATE StudentAnswers SET QuestionID = StudentAnswerID%9+1
		WHERE QuestionID IS NULL OR QuestionID = 0`)
	db.Exec(`
		INSERT INTO QuestionAssignmentMapping(AssignmentID, QuestionID)
		SELECT AssignmentID, QuestionID
		FROM CourseAssignments, Questions
		WHERE QuestionID % 2 != AssignmentID % 2`)
	db.Exec(`
		INSERT INTO Comments (CommentText)
		VALUES ('COMMENT #1'), ('COMMENT #2'), ('COMMENT #3'), ('MULTILINE COMMENT:
2: 1234567890ABCDEF ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC
3: 123 1234 45676756 87585765 5767
4: 1234567890ABCDEF ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC 123'),
		('this is not correct, you have selected an extra row in both return and probability which is unwarranted.'),
		('an extra row has been selected which is not correct, even though your answer is coming correct')`)
	db.Exec(`
		INSERT INTO BlockCommentMapping(ExcelBlockID, ExcelCommentID)
		SELECT ExcelBlockID, CommentID
		FROM ExcelBlocks AS b, Comments AS c
		WHERE c.CommentID %  8 = b.ExcelBlockID %  8`)

	var assignment model.Assignment
	db.First(&assignment, "State = ?", "GRADED")
	var question model.Question
	db.Joins("JOIN QuestionAssignmentMapping AS qa ON qa.QuestionID = Questions.QuestionID").
		Where("qa.AssignmentID = ?", assignment.ID).
		First(&question)
	for _, fn := range []string{"commenting.test.xlsx", "indirect.test.xlsx"} {

		isIndirect := strings.HasPrefix("indirect", fn)

		f := model.Source{
			FileName:     fn,
			S3BucketName: "studentanswers",
			S3Key:        fn,
		}
		db.Create(&f)
		answer := model.Answer{
			// AssignmentID:   assignment.ID,
			SourceID:       model.NewNullInt64(f.ID),
			QuestionID:     model.NewNullInt64(question.ID),
			SubmissionTime: *parseTime("2017-01-01 14:42"),
		}
		db.Create(&answer)
		book := model.Workbook{FileName: fn, Answer: answer}
		db.Create(&book)
		for _, sn := range []string{"Sheet1", "Sheet2"} {
			sheet := model.Worksheet{Name: sn, Workbook: book, WorkbookFileName: book.FileName, Answer: answer}
			db.Create(&sheet)
			chart := model.Chart{Worksheet: sheet}
			db.Create(&chart)
			block := model.Block{Worksheet: sheet, Range: "ChartTitle", Formula: "TEST", RelativeFormula: "E15", Chart: chart}
			db.Create(&block)
			if !isIndirect {
				comment := model.Comment{Text: fmt.Sprintf("+++ Comment for CHART in %q for the range %q", sn, "E15")}
				db.Create(&comment)
				bcm := model.BlockCommentMapping{Block: block, Comment: comment}
				db.Create(&bcm)
			}
			for i, r := range []string{"A1", "C3", "D2:F13", "C2:F14", "C12:E21", "C13:D16"} {
				block := model.Block{Worksheet: sheet, Range: r, Formula: fmt.Sprintf("FORMULA #%d", i)}
				db.Create(&block)
				var ct string
				if !isIndirect {
					switch {
					case i < 3:
						ct = fmt.Sprintf("*** Comment in %q for the range %q", sn, r)
					case i == 4:
						ct = `MULTILINE COMMENT:
2: 1234567890ABCDEF ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC
3: 123 1234 45676756 87585765 5767
4: 1234567890ABCDEF ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC 123`
					case i == 5:
						ct = "this is not correct, you have selected an extra row in both return and probability which is unwarranted."
					default:
						ct = "an extra row has been selected which is not correct, even though your answer is coming correct"
					}
					comment := model.Comment{Text: ct}
					db.Create(&comment)
					bcm := model.BlockCommentMapping{Block: block, Comment: comment}
					db.Create(&bcm)
				}
			}
		}
	}
	err := db.Exec(`INSERT INTO Cells(block_id, worksheet_id, cell_range, value)
		SELECT b.ExcelBlockID, b.worksheet_id, r.range, 'CELL VALUE'
		FROM ExcelBlocks AS b JOIN
		(
			SELECT 'A1' AS v, 'A1' AS "range"
			UNION SELECT 'C3', 'C3'
			UNION SELECT 'D1:F2', 'D1'
			UNION SELECT 'D1:F2', 'E1'
			UNION SELECT 'D1:F2', 'F1'
			UNION SELECT 'D1:F2', 'D2'
			UNION SELECT 'D1:F2', 'E2'
			UNION SELECT 'D1:F2', 'F2'
			UNION SELECT 'C4:F5', 'C4'
			UNION SELECT 'C4:F5', 'D4'
			UNION SELECT 'C4:F5', 'E4'
			UNION SELECT 'C4:F5', 'F4'
			UNION SELECT 'C4:F5', 'C5'
			UNION SELECT 'C4:F5', 'D5'
			UNION SELECT 'C4:F5', 'E5'
			UNION SELECT 'C4:F5', 'F5'
			UNION SELECT 'D2:F13', 'F3'
			UNION SELECT 'D2:F13', 'F4'
			UNION SELECT 'D2:F13', 'F5'
			UNION SELECT 'D2:F13', 'F6'
			UNION SELECT 'D2:F13', 'F7'
			UNION SELECT 'D2:F13', 'F8'
			UNION SELECT 'D2:F13', 'F9'
			UNION SELECT 'D2:F13', 'F10'
			UNION SELECT 'D2:F13', 'F11'
			UNION SELECT 'D2:F13', 'F12'
			UNION SELECT 'D2:F13', 'F13'
	) AS r ON r.v = b.BlockCellRange
	LEFT JOIN Cells AS ce ON ce.block_id = b.ExcelBlockID AND r.range = ce.cell_range
	WHERE ce.id IS NULL`).Error
	if err != nil {
		t.Error(err)
	}
	var cells []model.Cell
	res := db.Where("CommentID IS NULL AND id %  2 = 1").Find(&cells)
	if res.Error != nil {
		t.Error(res.Error)
	}
	for no, c := range cells {
		c.Comment = model.Comment{Text: fmt.Sprintf("VALUE COMMENT FOR %q // %d", c.Range, no)}
		c.Value = fmt.Sprintf("CELL %q VALUE: %d", c.Range, no)
		db.Save(&c)
	}
	if err := db.Exec(`
		INSERT INTO StudentAnswerCommentMapping(StudentAnswerID, CommentID)
		SELECT DISTINCT wb.StudentAnswerID, bc.ExcelCommentID
		FROM WorkBooks AS wb, BlockCommentMapping AS bc
		JOIN ExcelBlocks AS b ON b.ExcelBlockID = bc.ExcelBlockID
		JOIN WorkSheets AS s ON s.id = b.worksheet_id
		WHERE s.workbook_file_name = 'commenting.test.xlsx'
			AND wb.file_name IN ('indirect.test.xlsx', 'commenting.test.xlsx')`).Error; err != nil {
		t.Error(err)
	}

	var blockUpdate, cellUpdate string
	if db.Dialect().GetName() == "sqlite3" {
		blockUpdate = `WITH b AS (
				SELECT
				b.ExcelBlockID AS id,
				SUBSTR(b.BlockCellRange, 1, INSTR(b.BlockCellRange,':')-1) AS s,
				SUBSTR(b.BlockCellRange, INSTR(b.BlockCellRange,':')+1) AS e
				FROM ExcelBlocks AS b
				WHERE INSTR(b.BlockCellRange, ':') > 0),
			u AS (
				SELECT
				b.id,
				CAST(LTRIM(s, RTRIM(s, '0123456789')) AS INTEGER)-1 AS tr,
				unicode(RTRIM(s, '0123456789'))-unicode('A') AS lc,
				CAST(LTRIM(e, RTRIM(e, '0123456789')) AS INTEGER)-1 AS br,
				unicode(RTRIM(e, '0123456789'))-unicode('A') AS rc
				FROM b)
			UPDATE ExcelBlocks
			SET
				t_row = (SELECT tr FROM u WHERE u.id = ExcelBlocks.ExcelBlockID),
				l_col = (SELECT lc FROM u WHERE u.id = ExcelBlocks.ExcelBlockID),
				b_row = (SELECT br FROM u WHERE u.id = ExcelBlocks.ExcelBlockID),
				r_col = (SELECT rc
			FROM u WHERE u.id = ExcelBlocks.ExcelBlockID)
			WHERE b_row IS NULL OR b_row <= 0`
		cellUpdate = `UPDATE Cells
			SET row=CAST(LTRIM(cell_range, RTRIM(cell_range, '0123456789')) AS INTEGER)-1,
			col=unicode(RTRIM(cell_range, '0123456789'))-unicode('A')
			WHERE col IS NULL or row IS NULL OR (col <= 0  AND row <= 0)`
	} else {
		if err := db.Exec("SET sql_mode = ?", "ANSI").Error; err != nil {
			t.Error(err)
		}
		blockUpdate = `UPDATE ExcelBlocks, (
			SELECT  b.*,
			CAST(
			CASE s REGEXP '[A-Za-z]{2}[0-9]+'
			WHEN 1 THEN right(s, length(s)-2)
			ELSE right(s, length(s)-1)
			END AS UNSIGNED INTEGER)-1 AS tr,
			ascii(CASE s REGEXP '[A-Za-z]{2}[0-9]+'
			WHEN 1 THEN left(s, 2)
			ELSE left(s, 1) END)-ascii('A') AS lc,
			CAST(
			CASE e REGEXP '[A-Za-z]{2}[0-9]+'
			WHEN 1 THEN right(e, length(e)-2)
			ELSE right(e, length(e)-1)
			END AS UNSIGNED INTEGER)-1 AS br,
			ascii(CASE e REGEXP '[A-Za-z]{2}[0-9]+'
			WHEN 1 THEN left(e, 2)
			ELSE left(e, 1) END)-ascii('A') AS rc
			FROM (
			SELECT
			b.ExcelBlockID,
			SUBSTR(b.BlockCellRange, 1, INSTR(b.BlockCellRange,':')-1) AS s,
			SUBSTR(b.BlockCellRange, INSTR(b.BlockCellRange,':')+1) AS e
			FROM ExcelBlocks AS b
			WHERE INSTR(b.BlockCellRange, ':') > 0) AS b
			) AS u
			SET t_row = tr, l_col = lc, b_row = br, r_col = rc
			WHERE u.ExcelBlockID = ExcelBlocks.ExcelBlockID
			AND (b_row IS NULL OR b_row <= 0)`
		cellUpdate = `
			UPDATE Cells, (
			SELECT id, cell_range,
			CAST(
			  CASE cell_range REGEXP '[A-Za-z]{2}[0-9]+'
			    WHEN 1 THEN right(cell_range, length(cell_range)-2)
			    ELSE right(cell_range, length(cell_range)-1)
			  END AS UNSIGNED INTEGER)-1 AS r,
			ascii(CASE cell_range REGEXP '[A-Za-z]{2}[0-9]+'
			WHEN 1 THEN left(cell_range, 2)
			ELSE left(cell_range, 1) END)-ascii('A') AS c
			FROM Cells
			WHERE "col" IS NULL or "row" IS NULL OR ("col" <= 0  AND "row" <= 0)
			) AS u
			SET Cells."row"=u.r, Cells.col=u.c
			WHERE Cells.id = u.id`
	}
	if err := db.Exec(blockUpdate).Error; err != nil {
		t.Log("*** SQL:\n", blockUpdate)
		t.Error(err)
	}
	if err := db.Exec(cellUpdate).Error; err != nil {
		t.Log("*** SQL:\n", cellUpdate)
		t.Error(err)
	}
	t.Run("Queries", testQueries)
	t.Run("GetCommentRows", testGetCommentRows)
	t.Run("RowsToComment", testRowsToComment)
	t.Run("Comments", testComments)
	t.Run("CellComments", testCellComments)
}

func testGetCommentRows(t *testing.T) {
	var ws model.Worksheet
	db.Where("workbook_file_name = ?", "commenting.test.xlsx").First(&ws)
	t.Log(ws)
	rows, err := ws.GetBlockComments()
	if err != nil {
		t.Error(err)
	} else {
		for i, r := range rows {
			t.Log(i, ": ", r)
		}
	}
	comments, err := ws.GetCellComments()
	if err != nil {
		t.Error(err)
	} else {
		for i, r := range comments {
			t.Log(i, ": ", r)
		}
	}
}

func testQueries(t *testing.T) {
	var book model.Workbook
	db.First(&book, "file_name LIKE ?", "%commenting.test.xlsx")
	if book.FileName != "commenting.test.xlsx" {
		t.Errorf("Expected 'commenting.test.xlsx', got: %q", book.FileName)
		t.Logf("%#v", book)
	}
	var answerCount int
	if err := db.DB().QueryRow(`
			SELECT COUNT(DISTINCT StudentAnswerID) AnswerCount
			FROM StudentAnswerCommentMapping`).Scan(&answerCount); err != nil {
		t.Error(err)
	}
	if answerCount != 2 {
		t.Errorf("Expecte 2 answers mapped got: %d", answerCount)
	}
}

func testComments(t *testing.T) {

	outputName := utils.TempFileName("", ".xlsx")
	t.Log("OUTPUT:", outputName)
	cmd.AddComments("commenting.test.xlsx", outputName)

	// xlFile, err := xlsx.OpenFile(outputName)
	// if err != nil {
	// 	t.Error(err)
	// }

	// TODO: fix comment loading...
	// sheet := xlFile.Sheets[0]
	// comment := sheet.Comment["D2"]
	// expect := `*** Comment in "Sheet1" for the range "D2:F13"`
	// if comment.Text != expect {
	// 	t.Errorf("Expected %q, got: %q", expect, comment.Text)
	// }

	outputName = utils.TempFileName("", ".xlsx")
	t.Log("OUTPUT:", outputName)
	cmd.RootCmd.SetArgs([]string{"comment", "-U", url, "commenting.test.xlsx", outputName})
	cmd.Execute()

	// xlFile, err = xlsx.OpenFile(outputName)
	// if err != nil {
	// 	t.Error(err)
	// }

	// TODO: fix comment loading...
	// sheet = xlFile.Sheets[0]
	// comment = sheet.Comment["D2"]
	// if comment.Text != expect {
	// 	t.Errorf("Expected %q, got: %q", expect, comment.Text)
	// }
}

func testCellComments(t *testing.T) {
	fileName := "cell_commenting.xlsx"
	f := model.Source{
		FileName:     fileName,
		S3BucketName: "studentanswers",
		S3Key:        fileName,
	}
	db.Create(&f)
	assignment := model.Assignment{Title: "ASSIGNMENT 'Test Cell Comments'"}
	db.Create(&assignment)
	question := model.Question{
		SourceID:     model.NewNullInt64(f.ID),
		QuestionType: model.QuestionType("FileUpload"),
		QuestionText: "Question wiht merged cells",
		MaxScore:     8888.88,
		AuthorUserID: 123456789,
		WasCompared:  true,
		IsFormatting: true,
	}
	db.Create(&question)
	answer := model.Answer{
		// AssignmentID:   assignment.ID,
		SourceID:       model.NewNullInt64(f.ID),
		QuestionID:     model.NewNullInt64(question.ID),
		SubmissionTime: *parseTime("2017-01-01 14:42"),
	}
	db.Create(&answer)
	book := model.Workbook{FileName: fileName, Answer: answer}
	db.Create(&book)

	sheet := model.Worksheet{
		Name:             "Sheet1",
		Workbook:         book,
		WorkbookFileName: book.FileName,
		Answer:           answer}
	db.Create(&sheet)
	chart := model.Chart{Worksheet: sheet}
	db.Create(&chart)
	block := model.Block{
		Worksheet:       sheet,
		Range:           "ChartTitle",
		Formula:         "TEST",
		RelativeFormula: "E15",
		Chart:           chart}
	db.Create(&block)
	block = model.Block{
		Worksheet: sheet,
		Range:     "D8:D20",
		Formula:   "B8*C8",
		TRow:      7,
		LCol:      3,
		BRow:      19,
		RCol:      3,
	}
	db.Create(&block)
	comments := []model.Comment{
		{Text: "cell comment 40 1742"},
		{Text: "cell comment 40 1743"},
		{Text: "block comment 1101"},
	}
	for i := range comments {
		db.Create(&comments[i])
		db.Create(&model.AnswerComment{Answer: answer, Comment: comments[i]})
	}
	cells := []model.Cell{
		{Range: "D8", Row: 7, Col: 3, Block: block, Formula: "B8*C8", Value: "80", Comment: comments[0], Worksheet: sheet},
		{Range: "D9", Row: 8, Col: 3, Block: block, Formula: "B9*C9", Value: "48", Comment: comments[1], Worksheet: sheet},
	}
	for i := range cells {
		db.Create(&cells[i])
	}
	db.Create(&model.BlockCommentMapping{Block: block, Comment: comments[2]})
	block = model.Block{
		Worksheet: sheet,
		Range:     "A2:B5",
		Formula:   "B8*C8",
		TRow:      1,
		LCol:      0,
		BRow:      4,
		RCol:      1,
	}
	db.Create(&block)
	comments = []model.Comment{
		{Text: "CELL COMMENT #1"},
		{Text: "CELL COMMENT #2"},
		{Text: "BLOCK COMMENT"},
	}
	for i := range comments {
		db.Create(&comments[i])
		db.Create(&model.AnswerComment{Answer: answer, Comment: comments[i]})
	}
	for _, c := range []model.Cell{
		{Range: "A2", Row: 1, Col: 0, Block: block, Formula: "B8*C8", Value: "80", Comment: comments[0], Worksheet: sheet},
		{Range: "B3", Row: 2, Col: 1, Block: block, Formula: "B9*C9", Value: "48", Comment: comments[1], Worksheet: sheet},
	} {
		db.Create(&c)
	}
	db.Create(&model.BlockCommentMapping{Block: block, Comment: comments[2]})
	outputName := utils.TempFileName("", ".xlsx")
	t.Log("OUTPUT:", outputName)
	if err := cmd.AddCommentsToFile(int(book.AnswerID.Int64), fileName, outputName, true); err != nil {
		log.Errorln(err)
	}
}

func testRowsToComment(t *testing.T) {

	rows, err := model.RowsToComment(-1)
	if err != nil {
		t.Fatal(err)
	}

	if expected, got := 10, len(rows); got != expected {
		t.Errorf("Expected to select %d files to comment, got: %d", expected, got)
	}
	if t.Failed() {
		for _, r := range rows {
			t.Log(r)
		}
	}
	rows, err = model.RowsToComment(999999)
	if err != nil {
		t.Fatal(err)
	}

	if expected, got := 0, len(rows); got != expected {
		t.Errorf("Expected to select %d files to comment, got: %d", expected, got)
	}
}

// TestNewUUID
func TestNewUUID(t *testing.T) {
	uuid1, err := utils.NewUUID()
	if err != nil {
		t.Error(err)
	}

	uuid2, err := utils.NewUUID()
	if err != nil {
		t.Error(err)
	}

	t.Logf("UUID: %q, %q", uuid1, uuid2)
	if uuid1 == uuid2 {
		t.Errorf("Expected different values: %q, %q", uuid1, uuid2)
	}
}
