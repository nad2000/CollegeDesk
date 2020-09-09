// Copyright © 2017,2018 Radomirs Cirskis
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
	"sort"
	"strings"

	log "github.com/Sirupsen/logrus"
	"github.com/nad2000/excelize"
	"github.com/nad2000/xlsx"
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
	flags.IntVarP(&assignmentID, "assignment", "a", -1, "The assignment ID to process (-1 - process all assignments)")
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
	res := Db.Order("ID DESC").First(&book, "StudentAnswerID IS NOT NULL AND file_name LIKE ?", "%"+base)
	if res.RecordNotFound() {
		log.Errorf("the workbook record not found for %q.", fileName)
		return
	}
	if res.Error != nil {
		log.Errorln(res.Error)
		return
	}
	if err := AddCommentsToFile(int(book.AnswerID.Int64), fileName, outputName, true); err != nil {
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

		if err := AddCommentsToFile(a.ID, fileName, outputName, true); err != nil {
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
			log.Errorf("failed to uploade the output file %q to %q with S3 key %q: %s",
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
		if err := Db.Model(&a).UpdateColumns(model.Answer{
			GradedFileID:        model.NewNullInt64(source.ID),
			WasCommentProcessed: 1,
		}).Error; err != nil {
			log.Errorf("Failed to update the answer entry: %s", err)
		}
		fileCount++

	}
	log.Infof("Successfully commented %d Excel files.", fileCount)
	return nil
}

type commentEntry struct {
	address     string
	row, col    int
	commentText string
	boxRow      int
}

// addCommentsToWorksheet
func addCommentsToWorksheet(file *excelize.File, sheetName string, comments map[int][]commentEntry) {
	cols := make([]int, 0)
	for c, column := range comments {
		if column != nil {
			cols = append(cols, c)
		}
	}
	sort.IntSlice(cols).Sort()

	for i := len(cols) - 1; i >= 0; i-- {
		col := cols[i]
		log.Debug("+++ COL: ", col, " WITH BOX AT ", i*2)
		addCommentsToColumn(file, sheetName, comments[col], i*2)
	}
}

// addCommentsToColumn
func addCommentsToColumn(file *excelize.File, sheetName string, column []commentEntry, boxCol int) {

	if column == nil || len(column) < 1 {
		return
	}
	if boxCol < 0 {
		boxCol = column[0].col + 1
	}
	address := column[0].address
	col, _, err := xlsx.GetCoordsFromCellIDString(address)
	if err != nil {
		log.Errorf("error occured while adding commment to column starting with %q: %s", address, err)
		col = column[0].col
	}

	nextBoxRow := 1
	for i := range column {
		cell := &column[i]
		hight := file.CommmentCalloutBoxHightAt(
			sheetName, fmt.Sprintf(`{"author":"Grader: ", "text":%q}`, cell.commentText), col, 0)
		log.Debug("-- ROW: ", nextBoxRow, ", HIGHT: ", hight)
		cell.boxRow, nextBoxRow = nextBoxRow, nextBoxRow+int(hight+0.5)
	}
	for _, cell := range column {
		if debug {
			log.Debugf(
				"Adding comment to %q sheet at %q: %s (box: %q)",
				sheetName, cell.address, cell.commentText,
				// Save the file w/o comments and reopen it:
				xlsx.GetCellIDStringFromCoords(boxCol, cell.boxRow))
		}
		log.Debugf("*** Adding a comment at (%d, %d): %v", cell.boxRow, boxCol, cell.commentText)
		file.AddCommentAt(
			sheetName,
			cell.address,
			fmt.Sprintf(`{"author":"Grader: ", "text":%q}`, cell.commentText),
			boxCol, cell.boxRow)
	}
}

// AddCommentsToFile addes chart properties and comments to the answer files.
func AddCommentsToFile(answerID int, fileName, outputName string, deleteComments bool) error {

	// Iterate via assosiated comments and add them to the file
	file, err := excelize.OpenFile(fileName)
	if err != nil {
		return fmt.Errorf("failed to open file %q: %s", fileName, err.Error())
	}
	if deleteComments {
		if model.DeleteAllComments(file) {
			// Save the file w/o comments and reopen it:
			fileName := utils.TempFileName("", filepath.Ext(fileName))
			log.Infof("Inermediate file saved to %q", fileName)
			err = file.SaveAs(fileName)
			if err != nil {
				return fmt.Errorf("failed to remove comments from file %q: %s", fileName, err.Error())
			}

			file, err = excelize.OpenFile(fileName)
			if err != nil {
				return fmt.Errorf("failed to open file %q: %s", fileName, err.Error())
			}
		}
	}
	if err := addChartProperties(file, answerID); err != nil {
		return err
	}

	var (
		address, commentText string
		answer               model.Answer
	)

	Db.Preload("Worksheets").First(&answer, answerID)
	log.Debug("*** Answer: ", answer)

	for _, sheet := range answer.Worksheets {
		log.Debug("*** Worksheet: ", sheet)
		blockComments, err := sheet.GetBlockComments()
		if err != nil {
			log.Error("Failed to retrieve the block comments: ", err)
			continue
		}
		cellComments, err := sheet.GetCellComments()
		if err != nil {
			log.Error("Failed to retrieve the comment comments: ", err)
			continue
		}
		cellCommentMap := make(map[string]model.CellCommentRow, len(cellComments))
		for _, cc := range cellComments {
			log.Debug("*** Cell Comment: ", cc)
			cellCommentMap[cc.Range] = cc
		}

		comments := make(map[int][]commentEntry)

		for bcCol, bcInCol := range blockComments {

			log.Debug("+++ Column: ", bcCol)

			for _, bc := range bcInCol {
				log.Debug("*** Block: ", bc)

				for col := bc.LCol; col <= bc.RCol; col++ {
					for row := bc.TRow; row <= bc.BRow; row++ {
						if comments[bcCol] == nil {
							comments[bcCol] = make([]commentEntry, 0)
						}
						address = model.CellAddress(row, col)

						commentText = fmt.Sprintf("Points = %.2f. %s", bc.Marks, bc.CommentText)
						if cc, ok := cellCommentMap[address]; ok {
							if commentText != "" {
								commentText += "\n"
							}
							commentText += cc.CommentText
						}

						log.Debugf("COMMENT: %q, %q, %q", sheet.Name, address, commentText)

						if commentText != "" {
							comments[bcCol] = append(comments[bcCol],
								commentEntry{address, row, col, commentText, -1})
						}
					}
				}
			}
		}
		log.Debug("*** Collected comments:", comments)
		addCommentsToWorksheet(file, sheet.Name, comments)
	}

	if fileName == outputName || outputName == "" {
		err = file.Save()
	} else {
		err = file.SaveAs(outputName)
	}

	if err != nil {
		if outputName != "" {
			return fmt.Errorf("failed to save file %q -> %q: %s", fileName, outputName, err.Error())
		}
		return fmt.Errorf("failed to save file %q: %s", fileName, err.Error())
	}
	log.Infof("Outpu saved to %q", outputName)

	return nil
}

func addChartProperties(file *excelize.File, answerID int) error {
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
		file.SetCellStr(sheetName, cellAddress, propName)
		nextAddress, err := model.RelCellAddress(cellAddress, 0, 1)
		if err != nil {
			return err
		}
		file.SetCellStr(sheetName, nextAddress, propValue)
	}
	return nil
}
