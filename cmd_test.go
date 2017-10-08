package cmd_test

import (
	"os"
	"path"
	"strings"
	"testing"

	log "github.com/Sirupsen/logrus"

	"extract-blocks/cmd"

	"github.com/jinzhu/gorm"
	_ "github.com/jinzhu/gorm/dialects/sqlite"
)

var testDb string = path.Join(os.TempDir(), "extract-block-test.db")

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

func TestRowsToProcess(t *testing.T) {

	db, _ := gorm.Open("sqlite3", testDb)
	cmd.SetDb(db)
	//db.LogMode(true)
	defer db.Close()

	f1 := cmd.Source{FileName: "test.xlsx"}
	db.Create(&f1)
	db.Create(&cmd.Answer{FileID: f1.ID})
	f2 := cmd.Source{FileName: "test2.xlsx"}
	db.Create(&f2)
	db.Create(&cmd.Answer{FileID: f2.ID})
	ignore := cmd.Source{FileName: "ignore.abc"}
	db.Create(&ignore)
	db.Create(&cmd.Answer{FileID: ignore.ID})

	rows, _ := cmd.RowsToProcess(db)
	type Result struct {
		FileID          int    `gorm:"column:FileID"`
		S3BucketName    string `gorm:"column:S3BucketName"`
		S3Key           string `gorm:"column:S3Key"`
		FileName        string `gorm:"column:FileName"`
		StudentAnswerID int    `gorm:"column:StudentAnswerID"`
	}

	rowCount := 0
	for rows.Next() {
		rowCount += 1
		var r Result
		db.ScanRows(rows, &r)
		if !strings.HasSuffix(r.FileName, ".xlsx") {
			t.Errorf("Expected only .xlsx extensions, got %q", r.FileName)
		}
	}
	if rowCount != 2 {
		t.Errorf("Expected 2 rows, got %d", rowCount)
	}

}
