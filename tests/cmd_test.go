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
	"Sample3_A2E1.xlsx",
	"Sample4_A2E1.xlsx",
	"test2.xlsx",
	"test.xlsx",
	"test_floats.xlsx",
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
	deletData()
	var wb model.Workbook
	cmd.RootCmd.SetArgs([]string{
		"run", "-U", url, "-t", "-f", "demo.xlsx"})
	cmd.Execute()

	db, _ := model.OpenDb(url)
	defer db.Close()

	result := db.First(&wb, "file_name = ?", "demo.xlsx")
	if result.Error != nil {
		t.Error(result.Error)
		t.Fail()
	}

	if wb.FileName != "demo.xlsx" {
		t.Logf("Missing workbook 'demo.xlsx'. Expected 'demo.xlsx', got: %q", wb.FileName)
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
}

func deletData() {

	if db == nil || db.DB() == nil {
		db, _ = model.OpenDb(url)
		defer db.Close()
	}
	for _, m := range []interface{}{
		&model.AnswerComment{},
		&model.BlockCommentMapping{},
		&model.QuestionAssignment{},
		&model.Comment{},
		&model.Cell{},
		&model.Block{},
		&model.Chart{},
		&model.Worksheet{},
		&model.Workbook{},
		&model.QuestionExcelData{},
		&model.Question{},
		&model.Answer{},
		&model.Source{}} {
		db.Delete(m)
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

	deletData()
	//db.LogMode(true)

	for _, fn := range testFileNames {
		f := model.Source{
			FileName:     fn,
			S3BucketName: "studentanswers",
			S3Key:        fn,
		}

		db.Create(&f)
		q := model.Question{
			SourceID:         sql.NullInt64{Int64: int64(f.ID), Valid: true},
			QuestionType:     model.QuestionType("FileUpload"),
			QuestionSequence: 123,
			QuestionText:     "QuestionText...",
			MaxScore:         9999.99,
			AuthorUserID:     123456789,
			WasCompared:      true,
		}
		db.Create(&q)
		db.Create(&model.Answer{
			SourceID:       f.ID,
			QuestionID:     sql.NullInt64{Int64: int64(q.ID), Valid: true},
			SubmissionTime: *parseTime("2017-01-01 14:42"),
		})
	}

	ignore := model.Source{FileName: "ignore.abc"}
	db.Create(&ignore)
	db.Create(&model.Answer{SourceID: ignore.ID, SubmissionTime: *parseTime("2017-01-01 14:42")})

	//db.LogMode(true)
	var lastFileId int
	db.DB().QueryRow("SELECT MAX(FileID) AS LastFileID FROM FileSources").Scan(&lastFileId)
	for i := lastFileId + 1; i < lastFileId+10; i++ {
		fn := "question" + strconv.Itoa(i)
		if i < lastFileId+4 {
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
	if len(rows) != 15 {
		t.Errorf("Expected 15 question rows, got %d", len(rows))
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
	if len(questions) != 9 {
		t.Errorf("Expected 9 rows, got %d", len(questions))
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
			"Expeced that the number of answers to be processed dorps, got %d before, and %d afater.",
			countBefore, countAfter)
	}
}

func TestProcessing(t *testing.T) {

	db = createTestDB()
	defer db.Close()

	t.Run("QuestionsToProcess", testQuestionsToProcess)
	t.Run("RowsToProcess", testRowsToProcess)
	t.Run("HandleAnswers", testHandleAnswers)
	t.Run("S3Downloading", testS3Downloading)
	t.Run("S3Uploading", testS3Uploading)
	t.Run("Questions", testQuestions)
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

	//db.LogMode(true)
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
		SELECT wb.id, 'Sheet1', FileName, sa.StudentAnswerID
		FROM FileSources NATURAL JOIN StudentAnswers AS sa
		JOIN WorkBooks AS wb ON wb.StudentAnswerID = sa.StudentAnswerID`)
	db.Exec(`
		INSERT INTO ExcelBlocks (worksheet_id, BlockCellRange)
		SELECT id, r.v
		FROM WorkSheets AS s LEFT JOIN ExcelBlocks AS b
		ON b.worksheet_id = s.id,
		(
			SELECT 'A1' AS v
			UNION SELECT 'C3'
			UNION SELECT 'D2:F1'
			UNION SELECT 'C2:F1'
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
4: 1234567890ABCDEF ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC 123')`)
	db.Exec(`
		INSERT INTO BlockCommentMapping(ExcelBlockID, ExcelCommentID)
		SELECT ExcelBlockID, CommentID
		FROM ExcelBlocks AS b, Comments AS c
		WHERE c.CommentID % 4 = b.ExcelBlockID % 4`)

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
			AssignmentID:   assignment.ID,
			SourceID:       f.ID,
			QuestionID:     sql.NullInt64{Int64: int64(question.ID), Valid: true},
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
			for i, r := range []string{"A1", "C3", "D2:F13", "C2:F14"} {
				block := model.Block{Worksheet: sheet, Range: r, Formula: fmt.Sprintf("FORMULA #%d", i)}
				db.Create(&block)
				var ct string
				if !isIndirect {
					if i < 3 {
						ct = fmt.Sprintf("*** Comment in %q for the range %q", sn, r)
					} else {
						ct = `MULTILINE COMMENT:
2: 1234567890ABCDEF ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC
3: 123 1234 45676756 87585765 5767
4: 1234567890ABCDEF ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC ABC 123`
					}
					comment := model.Comment{Text: ct}
					db.Create(&comment)
					bcm := model.BlockCommentMapping{Block: block, Comment: comment}
					db.Create(&bcm)
				}
			}
		}
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

	t.Run("Queries", testQueries)
	t.Run("RowsToComment", testRowsToComment)
	t.Run("Comments", testComments)
}

func testQueries(t *testing.T) {
	// db.LogMode(true)
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

	outputName := path.Join(os.TempDir(), nextRandomName()+".xlsx")
	t.Log("OUTPUT:", outputName)
	// db.LogMode(true)
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

	outputName = path.Join(os.TempDir(), nextRandomName()+".xlsx")
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

func testRowsToComment(t *testing.T) {

	// db.LogMode(true)
	rows, err := model.RowsToComment()
	if err != nil {
		t.Fatal(err)
	}

	if expected, got := 8, len(rows); got != expected {
		t.Errorf("Expected to select %d files to comment, got: %d", expected, got)
	}
	if t.Failed() {
		for _, r := range rows {
			t.Log(r)
		}
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
