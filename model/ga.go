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
	isNotPlagiarised          bool
}

// GetGAEntries - build the GA entry list (map) for the question and answer file
func (q *Question) GetGAEntries(file *xlsx.File, userID int) (entries map[int]GARow, err error) {

	GA, ok := file.Sheet[gradingAssistanceSheetName]
	if ok {
		entries = make(map[int]GARow)
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
			name := GA.Cell(i, 1).Value
			if value == "" {
				break
			}
			value = GA.Cell(i, 2).Value
			if value == "" {
				break
			}
			sheetID, err := strconv.Atoi(value)
			if err != nil {
				log.Error(err)
				continue
			}
			sheetID -= userID
			entries[sheetNo] = GARow{sheetID: sheetID, userID: userID, name: name}
		}
		userIDs := make([]int, 0, len(entries))
		for _, v := range entries {
			userIDs = append(userIDs, v.userID)
		}

		sheetIDs := make([]int, 0, len(entries))
		for _, v := range entries {
			sheetIDs = append(sheetIDs, v.sheetID)
		}

		var rows *sql.Rows
		rows, err = Db.Raw(`
			SELECT DISTINCT
				qs.ProblemWorkSheetsID,
				qs.Sheet_Sequence,
				qs.Sheet_Name
			FROM XLQTransformation AS xt 
			JOIN QuestionFileWorkSheets AS qs ON qs.QuestionFileID = xt.questionfile_id
			WHERE xt.UserID = ? AND xt.QuestionID = ? AND qs.ProblemWorkSheetsID`,
			userIDs, q.ID, sheetIDs).Rows()
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
			r, ok := entries[sequence]
			if !ok {
				log.Errorf("missing entry in the 'Details' spreadsheet for %q", name)
				continue
			}
			if r.name == name && r.sequence == sequence {
				r.isNotPlagiarised = true
				entries[sequence] = r
			}
		}
		return
	}
	return nil, nil
}
