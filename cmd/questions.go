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
	model "extract-blocks/model"
	"extract-blocks/s3"

	log "github.com/Sirupsen/logrus"
	"github.com/spf13/cobra"
)

// questionsCmd represents the questions command
var questionsCmd = &cobra.Command{
	Use:   "questions",
	Short: "Process questions and questions workbooks.",
	Long:  `Process questions and questions workbooks.`,
	Run:   processQuestions,
}

func processQuestions(cmd *cobra.Command, args []string) {
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

	manager := createS3Manager()
	HandleQuestions(manager)
}

// HandleQuestions - iterates through questions, downloads the all files
// and inport all cells into DB
func HandleQuestions(manager s3.FileManager) error {

	rows, err := model.QuestionsToProcess()
	if err != nil {
		log.Fatalf("Failed to retrieve list of question source files to process: %s",
			err.Error())
	}
	var fileCount int
	for _, q := range rows {
		result := Db.Where("QuestionID = ?", q.ID).Delete(&model.QuestionExcelData{})
		if result.Error != nil {
			log.Error(result.Error)
		}
		var s model.Source
		Db.Model(&q).Related(&s, "FileID")
		fileName, err := s.DownloadTo(manager, dest)
		if err != nil {
			log.Error(err)
			continue
		}
		log.Infof("Processing %q", fileName)
		err = q.ImportFile(fileName)
		if err != nil {
			log.Errorf(
				"Failed to import %q for the question %#v: %s", fileName, q, Db.Error)
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

func init() {
	RootCmd.AddCommand(questionsCmd)
}
