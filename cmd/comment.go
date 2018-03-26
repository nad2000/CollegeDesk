// Copyright © 2017 Radomirs Cirskis
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR ONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

package cmd

import (
	model "extract-blocks/model"
	"fmt"
	"path/filepath"
	"strings"

	log "github.com/Sirupsen/logrus"
	"github.com/spf13/cobra"
	//"github.com/tealeg/xlsx"
	"github.com/360EntSecGroup-Skylar/excelize"
)

// commentCmd represents the comment command
var commentCmd = &cobra.Command{
	Use:   "comment [INPUT] [OUTPUT]",
	Short: "Add comments from DB",
	Long: `
Adds comments to the answer Excel Workbooks either in batch
or to a sible file given as an input. If the out put also is give
the new file will be stored with the given name.`,
	Run: func(cmd *cobra.Command, args []string) {
		AddComments(args...)
	},
}

func init() {
	RootCmd.AddCommand(commentCmd)
}

// AddComments addes comments to the given file from the DB and stores file with the given name
func AddComments(fileNames ...string) {

	hasInputFile := len(fileNames) > 0
	if !hasInputFile {
		AddCommentsInBatch()
		return
	}

	var fileName, outputName string

	if hasInputFile {
		fileName = fileNames[0]
		if len(fileNames) > 1 {
			outputName = fileNames[1]
		}
	}
	xlsx, err := excelize.OpenFile(fileName)
	if err != nil {
		log.Errorf("Failed to open file %q", fileName)
		log.Errorln(err)
		return
	}

	var book model.Workbook
	base := filepath.Base(fileName)
	Db.Preload("Worksheets.Blocks.CommentMappings.Comment").First(&book, "file_name LIKE ?", "%"+base)
	for _, sheet := range book.Worksheets {
		for _, block := range sheet.Blocks {
			for _, bcm := range block.CommentMappings {
				comment := bcm.Comment
				rangeCells := strings.Split(block.Range, ":")
				xlsx.AddComment(
					sheet.Name,
					rangeCells[0],
					fmt.Sprintf(`{"author": %q,"text": %q}`, "????", comment.Text))
			}
		}
	}
	if fileName == outputName || outputName == "" {
		err = xlsx.Save()
	} else {
		err = xlsx.SaveAs(outputName)
	}
	if err != nil {
		if outputName != "" {
			log.Errorf("Failed to save file %q -> %q", fileName, outputName)
		} else {
			log.Errorf("Failed to save file %q", fileName)
		}
		log.Errorln(err)
	}
}

// AddCommentsInBatch addes comments to the answer files.
func AddCommentsInBatch() {
	// TODO: ...
	return
}
