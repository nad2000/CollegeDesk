// Copyright © 2018 Radomirs Cirskis <nad2000@gmail.com>
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

	log "github.com/Sirupsen/logrus"
	"github.com/spf13/cobra"
)

// autocommentCmd represents the autocomment command
var autocommentCmd = &cobra.Command{
	Use:   "autocomment",
	Short: "Insert hard coded comments",
	Long: `This command inserts hard coded comments to the student answers:

The command does the following

For all StudentAnswerID where StudentAnswers.was_autocommented = 0, retrieve the CellIds belonging to the StudentAnswerID and implement the following steps.

*(1)* For a given Cell ID, IF WorkSheets.is_plagiarised = 0 AND AutoEvaluation.IsValueCorrect = 1
Then
(a) Insert a new row in comments table with the CommentText = *"Answer is correct"* , Marks = 1 
(b) insert a new row in StudentAnswerCommentMapping with CommentID retreived in step (a) and the StudentAnswerID of the CellID
(c) Update Cells table, set Cells.CommentID = CommentID retreived in step (a) 

*(2)* For a given Cell ID, IF WorkSheets.is_plagiarised = 0 AND AutoEvaluation.IsValueCorrect = 0
Then
(a) Insert a new row in comments table with the  CommentText = *"Answer is wrong"* , Marks = 0 
(b) insert a new row in StudentAnswerCommentMapping with CommentID retreived in step (a) and the StudentAnswerID of the CellID 
(c) Update Cells table, set Cells.CommentID = CommentID retreived in step (a) 
 
*(3)* For a given Cell ID, IF WorkSheets.is_plagiarised = 1 then 
(a) insert a new row in StudentAnswerCommentMapping with CommentID  = 12345 and the StudentAnswerID of the CellID
(b) Update Cells table, set Cells.CommentID = 12345

*(4)* After processing all the CellIds for the given StudentAnswerID , set StudentAnswers.was_autocommented = 1.`,
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

		model.AutoCommentAnswerCells()
	},
}

func init() {
	RootCmd.AddCommand(autocommentCmd)

	// Here you will define your flags and configuration settings.

	// Cobra supports Persistent Flags which will work for this command
	// and all subcommands, e.g.:
	// autocommentCmd.PersistentFlags().String("foo", "", "A help for foo")

	// Cobra supports local flags which will only run when this command
	// is called directly, e.g.:
	// autocommentCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")
}
