package cmd_test

import (
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

	_ "github.com/go-sql-driver/mysql"
	"github.com/jinzhu/gorm"
	_ "github.com/jinzhu/gorm/dialects/sqlite"
	"github.com/nad2000/xlsx"
)

var testFileNames = []string{
	"demo.xlsx",
	"Sample3_A2E1.xlsx",
	"Sample4_A2E1.xlsx",
	"test2.xlsx",
	"test.xlsx",
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
	if count != 3 {
		t.Errorf("Expected 3 blocks, got: %d", count)
	}
	db.Model(&model.Cell{}).Count(&count)
	if count != 30 {
		t.Errorf("Expected 30 cells, got: %d", count)
	}
}

func deletData() {

	if db == nil || db.DB() == nil {
		db, _ = model.OpenDb(url)
		defer db.Close()
	}

	db.Delete(&model.Comment{})
	db.Delete(&model.Cell{})
	db.Delete(&model.Block{})
	db.Delete(&model.Worksheet{})
	db.Delete(&model.Workbook{})
	db.Delete(&model.QuestionExcelData{})
	db.Delete(&model.Question{})
	db.Delete(&model.Answer{})
	db.Delete(&model.Source{})
}

func createTestDB() *gorm.DB {
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
		db.Create(&model.Answer{
			FileID:         f.ID,
			SubmissionTime: *parseTime("2017-01-01 14:42"),
		})
	}

	ignore := model.Source{FileName: "ignore.abc"}
	db.Create(&ignore)
	db.Create(&model.Answer{FileID: ignore.ID, SubmissionTime: *parseTime("2017-01-01 14:42")})

	//db.LogMode(true)
	for i := 101; i < 110; i++ {
		fn := "question" + strconv.Itoa(i)
		if i < 104 {
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
			//FileID:           sql.NullInt64{Int64: int64(i), Valid: true},
			FileID:           model.NewNullInt64(i),
			QuestionType:     model.QuestionType("FileUpload"),
			QuestionSequence: 123,
			QuestionText:     "QuestionText...",
			MaxScore:         9999.99,
			AuthorUserID:     123456789,
			WasCompared:      true,
		})

	}
	db.Create(&model.Assignment{Title: "ASSIGNMENT #1", State: "GRADED"})
	db.Create(&model.Assignment{Title: "ASSIGNMENT #2"})
	db.Exec("UPDATE StudentAnswers SET QuestionID = StudentAnswerID%9+1")
	db.Exec(`
		INSERT INTO QuestionAssignmentMapping(AssignmentID, QuestionID)
		SELECT AssignmentID, QuestionID 
		FROM CourseAssignments, Questions 
		WHERE QuestionID % 2 != AssignmentID % 2`)
	db.Exec(`
		INSERT INTO Comments (CommentText)
		VALUES ('COMMENT #1'), ('COMMENT #2'), ('COMMENT #3')`)
	db.Exec(`
		INSERT INTO BlockCommentMapping(ExcelBlockID, ExcelCommentID)
		SELECT ExcelBlockID, CommentID
		FROM ExcelBlocks, Comments`)

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
	if len(rows) != 9 {
		t.Errorf("Expected 9 question rows, got %d", len(rows))
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
	if len(questions) != 3 {
		t.Errorf("Expected 3 rows, got %d", len(questions))
	}

}

func testRowsToProcess(t *testing.T) {

	rows, _ := model.RowsToProcess()

	for _, r := range rows {
		if !strings.HasSuffix(r.FileName, ".xlsx") {
			t.Errorf("Expected only .xlsx extensions, got %q", r.FileName)
		}
	}
	if len(rows) != 5 {
		t.Errorf("Expected 5 rows, got %d", len(rows))
	}

}

type testDownloader struct{}

func (d testDownloader) DownloadFile(sourceName, s3BucketName, s3Key, dest string) (string, error) {
	return sourceName, nil
}

func testHandleAnswers(t *testing.T) {
	td := testDownloader{}
	cmd.HandleAnswers(&td)
}

func TestProcessing(t *testing.T) {

	db = createTestDB()
	defer db.Close()

	t.Run("QuestionsToProcess", testQuestionsToProcess)
	t.Run("RowsToProcess", testRowsToProcess)
	t.Run("HandleAnswers", testHandleAnswers)
	t.Run("S3Downloader", testS3Downloader)
	t.Run("Questions", testQuestions)
	t.Run("RowsToComment", testRowsToComment)
	t.Run("Comments", testComments)
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
	if count != 72 {
		t.Errorf("Expected 72 blocks, got: %d", count)
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

func testS3Downloader(t *testing.T) {

	if testing.Short() {
		t.Skip("Skipping S3 downloaer testing...")
	}

	d := s3.NewS3Downloader("us-east-1", "rad")
	destName := path.Join(os.TempDir(), nextRandomName()+".xlsx")
	_, err := d.DownloadFile("test.xlsx", "studentanswers", "test.xlsx", destName)
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

func testComments(t *testing.T) {

	fileName := "commenting.test.xlsx"
	book := model.Workbook{FileName: fileName}
	db.Create(&book)
	for _, sn := range []string{"Sheet1", "Sheet2"} {
		sheet := model.Worksheet{Name: sn, Workbook: book, WorkbookFileName: book.FileName}
		db.Create(&sheet)
		for _, r := range []string{"A1", "C3", "D2:F13"} {
			block := model.Block{Worksheet: sheet, Range: r}
			db.Create(&block)
			comment := model.Comment{Text: fmt.Sprintf("*** Comment in %q for the range %q", sn, r)}
			db.Create(&comment)
			bcm := model.BlockCommentMapping{Block: block, Comment: comment}
			db.Create(&bcm)
		}
	}
	outputName := path.Join(os.TempDir(), nextRandomName()+".xlsx")
	t.Log("OUTPUT:", outputName)
	cmd.AddComments(book.FileName, outputName)

	xlFile, err := xlsx.OpenFile(outputName)
	if err != nil {
		t.Error(err)
	}

	sheet := xlFile.Sheets[0]
	comment := sheet.Comment["D2"]
	expect := `*** Comment in "Sheet1" for the range "D2:F13"`
	if comment.Text != expect {
		t.Errorf("Expected %q, got: %q", expect, comment.Text)
	}

	outputName = path.Join(os.TempDir(), nextRandomName()+".xlsx")
	t.Log("OUTPUT:", outputName)
	cmd.RootCmd.SetArgs([]string{"comment", fileName, outputName})
	cmd.Execute()

	xlFile, err = xlsx.OpenFile(outputName)
	if err != nil {
		t.Error(err)
	}

	sheet = xlFile.Sheets[0]
	comment = sheet.Comment["D2"]
	if comment.Text != expect {
		t.Errorf("Expected %q, got: %q", expect, comment.Text)
	}
}

func testRowsToComment(t *testing.T) {

	rows, err := model.RowsToComment()
	if err != nil {
		t.Fatal(err)
	}
	defer rows.Close()

	var count int
	for rows.Next() {
		count++
		var r model.RowsToProcessResult
		db.ScanRows(rows, &r)
		t.Log(r)
	}
	if count != 3 {
		t.Errorf("Expected to select 3 files to comment, got: %d", count)
	}
}
