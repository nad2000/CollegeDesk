package model

import (
	"database/sql"
	"strconv"

	log "github.com/Sirupsen/logrus"
	"github.com/nad2000/xlsx"
)

const gradingAssistanceSheetName = "GA"

// GARow - GradingAssistance entry
type GARow struct {
	userID, sheetID, sequence int
	name                      string
}

// GetGAEntries - build the GA entry list (map) for the question and answer file
func (q *Question) GetGAEntries(file *xlsx.File) (sheetsToUserIDs map[int]GARow, err error) {

	GA, ok := file.Sheet[gradingAssistanceSheetName]
	if ok {
		sheetsToUserIDs = make(map[int]GARow)
		for i := 1; ; i++ {
			value := GA.Cell(i, 0).Value
			if value == "" {
				break
			}
			sheetNo, err := strconv.Atoi(value)
			if err != nil {
				log.Error(err)
				continue
			}
			value = GA.Cell(i, 2).Value
			if value == "" {
				break
			}
			userID, err := strconv.Atoi(value)
			if err != nil {
				log.Error(err)
				continue
			}
			sheetsToUserIDs[sheetNo] = GARow{userID: userID}

		}
		userIDs := make([]int, 0, len(sheetsToUserIDs))
		for _, v := range sheetsToUserIDs {
			userIDs = append(userIDs, v.userID)
		}

		var rows *sql.Rows
		rows, err = Db.Raw(`SELECT
				qs.ProblemWorkSheetsID,
				qs.Sheet_Sequence,
				qs.Sheet_Name
			FROM XLQTransformation AS xt 
			JOIN QuestionFileWorkSheets AS qs ON qs.QuestionFileID = xt.questionfile_id
			WHERE xt.UserID IN (?) AND xt.QuestionID = ?`, userIDs, q.ID).Rows()
		if err != nil {
			return
		}
		defer rows.Close()

		for rows.Next() {
			var (
				sheetID, sequence int
				name              string
			)

			rows.Scan(&sheetID, &sequence, &name)
			r, ok := sheetsToUserIDs[sequence]
			if !ok {
				log.Errorf("missing entry in the 'Details' spreadsheet for %q", name)
				continue
			}
			r.name = name
			r.sequence = sequence
			r.sheetID = sheetID
			sheetsToUserIDs[sequence] = r
		}
		return
	}
	return nil, nil
}
