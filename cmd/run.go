// Copyright © 2017 Radomirs Cirskis <nad2000@gmail.com>
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

package cmd

import (
	"database/sql"
	model "extract-blocks/model"
	"extract-blocks/s3"
	"path"

	log "github.com/Sirupsen/logrus"
	"github.com/spf13/cobra"
	"github.com/spf13/viper"
)

// runCmd represents the run command
var runCmd = &cobra.Command{
	Use:   "run",
	Short: "Process submitted student answer worksheets.",
	Long: `Retrievs list of the file sources for submitted answers, downloads the Excel
workbooks and extracts Cell Formula Blocks from Excel file and writes to MySQL.

Conditions that define Cell Formula Block -
    (i) Any contiguous (unbroken) range of excel cells containing cell formula
   (ii) Contiguous cells could be either in a row or in a column or in row+column cell block.
  (iii) The formula in the range of cells should be the same except the changes due to relative cell references.

Connection should be defined using connection URL notation: DRIVER://CONNECIONT_PARAMETERS,
where DRIVER is either "mysql" or "sqlite", e.g., mysql://user:password@/dbname?charset=utf8&parseTime=True&loc=Local.
More examples on connection parameter you can find at: https://github.com/go-sql-driver/mysql#examples.`,
	Run: extractBlocks,
}

func init() {
	RootCmd.AddCommand(runCmd)
	flags := runCmd.Flags()

	flags.BoolP("force", "f", false, "Repeat extraction if files were already handle.")
	flags.StringP("color", "c", defaultColor, "The block filling color.")

	viper.BindPFlag("color", flags.Lookup("color"))
	viper.BindPFlag("force", flags.Lookup("force"))
}

func extractBlocks(cmd *cobra.Command, args []string) {

	getConfig()
	debugCmd(cmd)

	var err error

	Db, err = model.OpenDb(url)
	if err != nil {
		log.Error(err)
		log.Fatalf("failed to connect database %q", url)
	}
	defer Db.Close()
	model.DebugLevel, model.VerboseLevel = debugLevel, verboseLevel

	if testing || len(args) > 0 {
		// read up the file list from the arguments
		for _, excelFileName := range args {
			q := model.Question{
				QuestionType:      "ShortAnswer",
				QuestionSequence:  0,
				QuestionText:      "DUMMY",
				AnswerExplanation: sql.NullString{String: "DUMMY", Valid: true},
				MaxScore:          999.99,
			}
			if !model.DryRun {
				Db.FirstOrCreate(&q, &q)
			}
			// Create Student answer entry
			a := model.Answer{
				ShortAnswer:    excelFileName,
				SubmissionTime: *parseTime("2017-01-01 14:42"),
				QuestionID:     model.NewNullInt64(q.ID),
			}
			if !model.DryRun {
				Db.FirstOrCreate(&a, &a)
			}
			model.ExtractBlocksFromFile(excelFileName, color, force, verbose, a.ID)
		}
	} else {
		manager := createS3Manager()
		HandleAnswers(manager)
	}
}

// HandleAnswers - iterates through student answers and retrievs answer workbooks
// it thaks the funcion that actuatualy performs file download from S3 bucket
// and returns the downloades file name or an error.
func HandleAnswers(manager s3.FileManager) error {

	model.ModelAnswerUserID = modelAnswerUserID
	rows, err := model.RowsToProcess()
	if err != nil {
		log.WithError(err).Fatalf("Failed to retrieve list of source files to process.")
	}
	var fileCount int
	for _, r := range rows {
		var a model.Answer
		if err := Db.First(&a, r.StudentAnswerID).Error; err != nil {
			log.WithError(err).Error("Failded to retrieve the answer entry with ID: ", r.StudentAnswerID)
			continue
		}
		destinationName := path.Join(dest, r.FileName)
		log.Infof(
			"Downloading %q (%q) form %q into %q",
			r.S3Key, r.FileName, r.S3BucketName, destinationName)
		fileName, err := manager.Download(
			r.FileName, r.S3BucketName, r.S3Key, destinationName)
		if err != nil {
			log.WithError(err).Errorf(
				"Failed to retrieve file %q from %q into %q",
				r.S3Key, r.S3BucketName, destinationName)
			continue
		}
		log.Infof("Processing %q", fileName)
		if _, err := model.ExtractBlocksFromFile(fileName, color, force, verbose, r.StudentAnswerID); err != nil {
			log.WithError(err).Errorln("Failed to process file: ", fileName)
		} else {
			fileCount++
		}
	}
	log.Infof("Downloaded and loaded %d Excel files.", fileCount)
	if missed := len(rows) - fileCount; missed > 0 {
		log.WithField("missed", missed).Infof("Failed to download and load %d file(s)", missed)
	}
	return nil
}
