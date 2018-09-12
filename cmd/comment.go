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
	"strconv"
	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"
	log "github.com/Sirupsen/logrus"
	"github.com/spf13/cobra"
)

var assignmentID int

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
		assignmentID, err = strconv.Atoi(flagString(cmd, "assignment"))
		if err != nil {
			log.Error(err)
			assignmentID = -1
		}

		Db, err = model.OpenDb(url)
		if err != nil {
			log.Error(err)
			log.Fatalf("Failed to connect database %q", url)
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
	flags := commentCmd.Flags()
	flags.IntP("assignment", "a", -1, "The assignment ID to process.")
}

// AddComments addes comments to the given file from the DB and stores file with the given name
func AddComments(fileNames ...string) {

	var fileName, outputName string

	fileName = fileNames[0]
	if len(fileNames) > 1 {
		outputName = fileNames[1]
	}

	var book model.Workbook
	base := filepath.Base(fileName)
	Db.First(&book, "file_name LIKE ?", "%"+base)
	if err := addCommentsToFile(book.AnswerID, fileName, outputName); err != nil {
		log.Errorln(err)
	}
}

// AddCommentsInBatch addes comments to the answer files.
func AddCommentsInBatch(manager s3.FileManager) error {

	rows, err := model.RowsToComment(assignmentID)
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

		if err := Db.Preload("Source").First(&a, r.StudentAnswerID).Error; err != nil {
			log.Error(err)
			continue
		}

		// Download the file and open it
		fileName, err := a.Source.DownloadTo(manager, dest)
		if err != nil {
			log.Error(err)
			continue
		}

		// Choose the output file name
		basename, extension := filepath.Base(fileName), filepath.Ext(fileName)
		outputName := path.Join(dest, strings.TrimSuffix(basename, extension)+"_Reviewed"+extension)

		if err := addCommentsToFile(a.ID, fileName, outputName); err != nil {
			log.Errorln(err)
		}

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

// AddCommentsToFile addes chart properties and comments to the answer files.
func addCommentsToFile(answerID int, fileName, outputName string) error {

	// Iterate via assosiated comments and add them to the file
	xlsx, err := excelize.OpenFile(fileName)
	if err != nil {
		return fmt.Errorf("Failed to open file %q: %s", fileName, err.Error())
	}

	if err := addChartProperties(xlsx, answerID); err != nil {
		return err
	}

	var sheetName, blockRange, commentText string
	commens, err := Db.Raw(`
        SELECT DISTINCT
          ws.name,
          CASE
            WHEN b.chart_id IS NULL THEN b.BlockCellRange
            ELSE b.relative_formula
          END AS CellRange,
          c.CommentText
        FROM StudentAnswers AS a
        JOIN WorkSheets AS ws
          ON ws.StudentAnswerID = a.StudentAnswerID
        JOIN ExcelBlocks AS b
          ON b.worksheet_id = ws.id
        JOIN BlockCommentMapping AS bc
          ON bc.ExcelBlockID = b.ExcelBlockID
        JOIN Comments AS c
          ON c.CommentID = bc.ExcelCommentID
        WHERE a.StudentAnswerID = ?
		`, answerID).Rows()
	if err != nil {
		return err
	}
	defer commens.Close()

	if err := addChartProperties(xlsx, answerID); err != nil {
		return err
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
			fmt.Sprintf(`{"author":"Grader: ", "text":%q}`, commentText))
	}

	if fileName == outputName || outputName == "" {
		err = xlsx.Save()
	} else {
		err = xlsx.SaveAs(outputName)
	}

	if err != nil {
		if outputName != "" {
			return fmt.Errorf("Failed to save file %q -> %q: %s", fileName, outputName, err.Error())
		}
		return fmt.Errorf("Failed to save file %q: %s", fileName, err.Error())
	}
	log.Infof("Outpu saved to %q", outputName)

	return nil
}

func addChartProperties(xlsx *excelize.File, answerID int) error {
	chartProperties, err := Db.Raw(`
			SELECT
				ws.name,
				b.relative_formula,
				b.BlockCellRange,
				b.BlockFormula
			FROM ExcelBlocks AS b
			JOIN WorkSheets AS ws
				ON ws.id = b.worksheet_id
			JOIN charts AS c
				ON c.id = b.chart_id
			WHERE ws.StudentAnswerID = ?
				AND b.relative_formula IS NOT NULL
				AND b.relative_formula <> ''`, answerID).Rows()
	if err != nil {
		return err
	}
	defer chartProperties.Close()
	var sheetName, propName, propValue, cellAddress string

	for chartProperties.Next() {
		chartProperties.Scan(&sheetName, &cellAddress, &propName, &propValue)
		xlsx.SetCellStr(sheetName, cellAddress, propName)
		nextAddress, err := model.RelCellAddress(cellAddress, 0, 1)
		if err != nil {
			return err
		}
		xlsx.SetCellStr(sheetName, nextAddress, propValue)
	}
	return nil
}
