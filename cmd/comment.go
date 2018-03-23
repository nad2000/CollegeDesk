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

	// model "extract-blocks/model"

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
		addComments(fileName)
	},
}

func init() {
	RootCmd.AddCommand(commentCmd)
}

func addComments(fileName string) {
	xlsx, err := excelize.OpenFile(fileName)
	if err != nil {
		log.Errorf("Fialed to open file %q", fileName)
		log.Errorln(err)
		return
	}
	xlsx.AddComment("Sheet1", "A1", `{"author":"Excelize: ","text":"This is a comment."}`)
	err = xlsx.SaveAs("NEW_" + fileName)
	if err != nil {
		log.Errorf("Fialed to save file %q", fileName)
		log.Errorln(err)
	}
}
