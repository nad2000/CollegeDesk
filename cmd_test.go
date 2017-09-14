package cmd_test

import (
	"os"
	"path"
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
}

func TestDemoFile(t *testing.T) {
	var wb cmd.Workbook
	cmd.RootCmd.SetArgs([]string{"-S", testDb, "-d", "-f", "-v", "demo.xlsx"})
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
