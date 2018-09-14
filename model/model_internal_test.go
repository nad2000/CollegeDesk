package model

import (
	"os"
	"path"
	"testing"

	log "github.com/Sirupsen/logrus"
	"github.com/nad2000/excelize"
)

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
