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
	"extract-blocks/model"
	"extract-blocks/s3"

	log "github.com/Sirupsen/logrus"
	"github.com/spf13/cobra"
)

// problemsCmd represents the questions command
var problemsCmd = &cobra.Command{
	Use:   "problems",
	Short: "Process problems and questions workbooks.",
	Long:  `Process problems and questions workbooks.`,
	Run:   processProblems,
}

func processProblems(cmd *cobra.Command, args []string) {
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
	HandleProblems(manager)
}

// HandleProblems - iterates through questions, downloads the all files
// and inport all cells into DB
func HandleProblems(manager s3.FileManager) (err error) {

	var problems []model.Problem
	// Db.LogMode(true)
	err = (Db.
		Preload("Source").
		Joins("JOIN FileSources ON FileSources.FileID = Problems.FileID").
		Where("IsProcessed = ?", 0).
		Where("FileSources.FileName LIKE ?", "%.xlsx").
		Find(&problems)).Error
	if err != nil {
		log.WithError(err).Fatalf("Failed to retrieve list of question source files to process.")
	}
	var fileCount int
	for _, p := range problems {
		if err := Db.Where("problem_id = ?", p.ID).Delete(&model.ProblemSheetData{}).Error; err != nil {
			log.WithError(err).Errorln("Failed to delete existing problem worksheet data of the problem: ", p)
		}
		if err := Db.Where("problem_id = ?", p.ID).Delete(&model.ProblemSheet{}).Error; err != nil {
			log.WithError(err).Errorln("Failed to delete existing problem worksheet entries of the problem: ", p)
		}
		fileName, err := p.Source.DownloadTo(manager, dest)
		if err != nil {
			log.WithError(err).Errorln("Failed to download the file: ", p.Source)
			continue
		}
		log.Infof("Processing %q", fileName)

		if err := p.ImportFile(fileName, color, verbose, manager); err != nil {
			log.WithError(err).Errorf("Failed to import %q for the question %#v", fileName, p)
			continue
		}

		if err := Db.Model(&p).UpdateColumn("IsProcessed", true).Error; err != nil {
			log.WithError(err).Errorf("Failed update problem entry %#v for %q.", p, fileName)
			continue
		}
		fileCount++
	}
	log.WithField("filecount", fileCount).
		Infof("Downloaded and loaded %d Excel files.", fileCount)
	if missedCount := len(problems) - fileCount; missedCount > 0 {
		log.WithField("missed", missedCount).
			Infof("Failed to download and load %d file(s)", missedCount)
	}
	return nil
}

func init() {
	RootCmd.AddCommand(problemsCmd)
}
