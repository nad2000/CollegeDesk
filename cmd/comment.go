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
	Use:   "comment FILENAME",
	Short: "Add comments from DB",
	Long: `TODO
	TODO
	TODO.`,
	Run: func(cmd *cobra.Command, args []string) {
		if len(args) < 1 {
			log.Fatal("Missing file name.")
		}
		fileName := args[0]
		if verboseLevel > 0 || debug {
			log.Infof("Processing %q", fileName)
		}
		AddComments(fileName, "")
	},
}

func init() {
	RootCmd.AddCommand(commentCmd)
}

// AddComments addes comments to the given file from the DB and stores file with the given name
func AddComments(fileName, outputName string) {
	xlsx, err := excelize.OpenFile(fileName)
	if err != nil {
		log.Errorf("Fialed to open file %q", fileName)
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
				log.Info("****", comment)
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
