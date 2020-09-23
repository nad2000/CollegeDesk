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
	"github.com/spf13/viper"
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

	manager := createManager()
	HandleQuestions(manager)
}

// HandleQuestions - iterates through questions, downloads the all files
// and inport all cells into DB
func HandleQuestions(manager s3.FileManager) error {

	rows, err := model.QuestionsToProcess()
	if err != nil {
		log.WithError(err).Fatalf("Failed to retrieve list of question source files to process.")
	}
	var fileCount int
	for _, q := range rows {
		if err := Db.Where("QuestionID = ?", q.ID).Delete(&model.QuestionExcelData{}).Error; err != nil {
			log.WithError(err).Errorln("Failed to delete existing question data of the qustion: ", q)
		}
		var s model.Source
		if err := Db.Model(&q).Related(&s, "FileID").Error; err != nil {
			log.WithError(err).Errorln("Failed to retrieve source file data entry for the question: ", q)
			continue
		}
		fileName, err := s.DownloadTo(manager, dest)
		if err != nil {
			log.WithError(err).Errorln("Failed to download the file: ", s)
			continue
		}
		log.Infof("Processing %q", fileName)

		if err := q.ImportFile(fileName, color, verbose, skipHidden); err != nil {
			log.WithError(err).Errorf("Failed to import %q for the question %#v", fileName, q)
			continue
		}
		q.IsProcessed = true

		if err := Db.Save(&q).Error; err != nil {
			log.WithError(err).Errorf("Failed update question entry %#v for %q.", q, fileName)
			continue
		}
		fileCount++
	}
	log.WithField("filecount", fileCount).
		Infof("Downloaded and loaded %d Excel files.", fileCount)
	if missedCount := len(rows) - fileCount; missedCount > 0 {
		log.WithField("missed", missedCount).
			Infof("Failed to download and load %d file(s)", missedCount)
	}
	return nil
}

func init() {
	RootCmd.AddCommand(questionsCmd)
	flags := questionsCmd.Flags()
	flags.StringVarP(&color, "color", "c", defaultColor, "The block filling color")
	viper.BindPFlag("color", flags.Lookup("color"))
}
