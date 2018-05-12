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
	"extract-blocks/s3"
	"extract-blocks/utils"
	"fmt"
	"path"
	"path/filepath"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	log "github.com/Sirupsen/logrus"
	"github.com/spf13/cobra"
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
		model.DebugLevel, model.VerboseLevel = debugLevel, verboseLevel
		getConfig()
		debugCmd(cmd)

		var err error
		Db, err = model.OpenDb(url)
		if err != nil {
			log.Error(err)
			log.Fatalf("failed to connect database %q", url)
		}
		defer Db.Close()
		if debugLevel > 1 {
			Db.LogMode(true)
		}

		if len(args) == 0 {
			manager := createS3Manager()
			AddCommentsInBatch(manager)
		} else {
			AddComments(args...)
		}
	},
}

func init() {
	RootCmd.AddCommand(commentCmd)
}

// AddComments addes comments to the given file from the DB and stores file with the given name
func AddComments(fileNames ...string) {

	var fileName, outputName string

	fileName = fileNames[0]
	if len(fileNames) > 1 {
		outputName = fileNames[1]
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
func AddCommentsInBatch(manager s3.FileManager) error {
	// TODO: ...

	rows, err := model.RowsToComment()
	if err != nil {
		log.Fatalf("Failed to retrieve list of question source files to process: %s",
			err.Error())
	}
	if len(rows) == 0 {
		log.Info("There is no files that can be commented.")
		return nil
	}
	var fileCount int
	for _, r := range rows {
		var a model.Answer
		err = Db.Preload("Source").First(&a, r.StudentAnswerID).Error

		var sheetName, blockRange, commentText string
		commens, err := Db.Raw(`
			SELECT
				ws.name, b.BlockCellRange, c.CommentText
			FROM StudentAnswers AS a
			JOIN StudentAnswerCommentMapping AS ac ON ac.StudentAnswerID = a.StudentAnswerID
			JOIN Comments AS c ON c.CommentID = ac.CommentID
			JOIN BlockCommentMapping AS bc ON bc.ExcelCommentID = c.CommentID
			JOIN ExcelBlocks AS b -- commented blocks
				ON b.ExcelBlockID = bc.ExcelBlockID
			JOIN WorkSheets AS cbws -- commented block worksheet
				ON cbws.id = b.worksheet_id
			JOIN StudentAnswers AS cba -- commented block answer
				ON cba.StudentAnswerID = cbws.StudentAnswerID
			JOIN ExcelBlocks AS ab -- answer blocks (that match...)
				ON ab.BlockFormula = b.BlockFormula
					-- AND ab.BlockCellRange = b.BlockCellRange
			JOIN WorkSheets AS ws -- answer work sheets
				ON ws.id = ab.worksheet_id
			WHERE
				a.QuestionID = cba.QuestionID
				AND a.StudentAnswerID = ?`, a.ID).Rows()
		if err != nil {
			log.Error(err)
			continue
		}
		defer commens.Close()

		// Download the file and open it
		fileName, err := a.Source.DownloadTo(manager, dest)
		if err != nil {
			log.Error(err)
			continue
		}

		// Iterate via assosiated comments and add them to the file
		xlsx, err := excelize.OpenFile(fileName)
		if err != nil {
			log.Errorf("Failed to open file %q", fileName)
			log.Errorln(err)
			continue
		}

		for commens.Next() {
			commens.Scan(&sheetName, &blockRange, &commentText)
			rangeCells := strings.Split(blockRange, ":")
			if debug {
				log.Debugf("Adding comment to %q sheet at %q: %s", sheetName, rangeCells[0], commentText)
			}
			xlsx.AddComment(
				sheetName,
				rangeCells[0],
				fmt.Sprintf(`{"text": %q}`, commentText))
		}

		basename, extension := filepath.Base(fileName), filepath.Ext(fileName)
		outputName := path.Join(dest, strings.TrimSuffix(basename, extension)+"_Reviewed"+extension)

		err = xlsx.SaveAs(outputName)
		log.Infof("Outpu saved to %q", outputName)

		// Upload the file
		newKey, err := utils.NewUUID()
		if err != nil {
			log.Error(err)
			continue
		}
		newKey += filepath.Ext(fileName)
		location, err := manager.Upload(outputName, a.Source.S3BucketName, newKey)
		if err != nil {
			log.Errorf("Failed to uploade the output file %q to %q with S3 key %q: %s",
				outputName, a.Source.S3BucketName, newKey, err)
			continue
		}
		log.Infof("Output file %q uploaded to bucket %q with S3 key %q, location: %q",
			outputName, a.Source.S3BucketName, newKey, location)

		// Associate the output file with the answer and mark the asnwer as 'COMMENTED'
		source := model.Source{
			FileName:     filepath.Base(outputName),
			S3BucketName: a.Source.S3BucketName,
			S3Key:        newKey,
		}
		Db.Create(&source)
		Db.Model(&a).UpdateColumns(model.Answer{Source: source, WasCommentProcessed: 1})

		fileCount++
	}
	log.Infof("Successfully commented %d Excel files.", fileCount)
	return nil
}
