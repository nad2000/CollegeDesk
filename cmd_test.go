package cmd_test

import (
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

	"github.com/jinzhu/gorm"
	_ "github.com/jinzhu/gorm/dialects/sqlite"
)

var testDb string = path.Join(os.TempDir(), "extract-block-test.db")

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
		relID := cmd.RelativeCellAddress(r.x, r.y, r.ID)
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
		relID := cmd.RelativeFormula(r.x, r.y, r.ID)
		if relID != r.expect {
			t.Errorf("Expecte %q for %#v; got %q", r.expect, r, relID)
		}
	}
}

func TestDemoFile(t *testing.T) {
	var wb cmd.Workbook
	cmd.RootCmd.SetArgs([]string{
		"-U", "sqlite3://" + testDb, "-t", "-d", "-f", "-v", "demo.xlsx"})
	cmd.Execute()
	db, _ := gorm.Open("sqlite3", testDb)
	defer db.Close()

	db.First(&wb, cmd.Workbook{FileName: "demo.xlsx"})
	if wb.FileName != "demo.xlsx" {
		t.Logf("Missing workbook 'demo.xlsx'. Expected 'demo.xlsx', got: %q", wb.FileName)
		t.Fail()
	}
	var count int
	db.Model(&cmd.Block{}).Count(&count)
	if count != 3 {
		t.Errorf("Expected 3 blocks, got: %d", count)
	}
	db.Model(&cmd.Cell{}).Count(&count)
	if count != 30 {
		t.Errorf("Expected 30 cells, got: %d", count)
	}
}

func createTestDB() *gorm.DB {
	db, _ = gorm.Open("sqlite3", testDb)
	cmd.Db = db
	cmd.SetDb()
	//db.LogMode(true)
	fileNames := []string{
		"demo.xlsx",
		"Sample3_A2E1.xlsx",
		"Sample4_A2E1.xlsx",
		"test2.xlsx",
		"test.xlsx",
	}

	for _, fn := range fileNames {
		f := cmd.Source{FileName: fn}
		db.Create(&f)
		db.Create(&cmd.Answer{FileID: f.ID, SubmissionTime: *parseTime("2017-01-01 14:42")})
	}

	ignore := cmd.Source{FileName: "ignore.abc"}
	db.Create(&ignore)
	db.Create(&cmd.Answer{FileID: ignore.ID, SubmissionTime: *parseTime("2017-01-01 14:42")})

	return db
}

var db *gorm.DB

func testRowsToProcess(t *testing.T) {

	rows, _ := cmd.RowsToProcess()

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

	t.Run("RowsToProcess", testRowsToProcess)
	t.Run("TestHandleAnswers", testHandleAnswers)
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

func TestS3Downloader(t *testing.T) {
	// if testing.Short() {
	// 	t.Skip("Skipping S3 downloaer testing...")
	// }

	d := cmd.NewS3Downloader("us-east-1", "rad")
	destName := nextRandomName() + ".xlsx"
	t.Logf("Downloading into %q", destName)
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
