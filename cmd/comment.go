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
		if len(args) == 0 {

			downloader := createS3Downloader()
			AddCommentsInBatch(downloader)
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
func AddCommentsInBatch(downloader FileDownloader) error {
	// TODO: ...

	rows, err := model.QuestionsToProcess()
	if err != nil {
		log.Fatalf("Failed to retrieve list of question source files to process: %s",
			err.Error())
	}
	var fileCount int
	for _, q := range rows {
		var s model.Source
		Db.Model(&q).Related(&s, "FileID")
		destinationName := path.Join(dest, s.FileName)
		log.Infof(
			"Downloading %q (%q) form %q into %q",
			s.S3Key, s.FileName, s.S3BucketName, destinationName)
		fileName, err := downloader.DownloadFile(
			s.FileName, s.S3BucketName, s.S3Key, destinationName)
		if err != nil {
			log.Errorf(
				"Failed to retrieve file %q from %q into %q: %s",
				s.S3Key, s.S3BucketName, destinationName, err.Error())
			continue
		}
		log.Infof("Processing %q", fileName)
		err = q.ImportFile(fileName)
		if err != nil {
			log.Errorf(
				"Failed to import %q for the question %#v: %s", fileName, q, Db.Error.Error())
			continue
		}
		q.IsProcessed = true
		Db.Save(&q)
		if Db.Error != nil {
			log.Errorf(
				"Failed update question entry %#v for %q: %s", q, fileName, Db.Error.Error())
			continue
		}
		fileCount++
	}
	log.Infof("Downloaded and loaded %d Excel files.", fileCount)
	if len(rows) != fileCount {
		log.Infof("Failed to download and load %d file(s)", len(rows)-fileCount)
	}
	return nil
}
