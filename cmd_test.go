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
	result := db.Count(&cmd.Block{})
	t.Log(result)
	t.Log(wb.FileName)
	t.Log("SUCCESS!")
}
