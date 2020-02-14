package tests

import (
	model "extract-blocks/model"
	"testing"

	"github.com/nad2000/xlsx"
)

func testGradingAssistanceData(t *testing.T) {

	const fileName = "data/StudentAnswer.xlsx"
	file, err := xlsx.OpenFile(fileName)
	if err != nil {
		t.Fatal(err)
	}

	var (
		s model.Source
		q model.Question
		p model.Problem
		u model.User
	)
	db.First(&s)
	db.First(&u)
	db.First(&q)
	p.SourceID = s.ID
	db.Create(&p)

	for _, r := range []struct {
		sequence, sheetID int
		name              string
	}{
		{2, 11684, "Q1"},
		{3, 11688, "Q2"},
		{4, 11691, "Q3"},
		{5, 11711, "Q4"},
		{6, 11781, "Q5"},
		{7, 11881, "Q6"},
		{8, 11981, "Q7"},
		{9, 11891, "Q8"},
		{10, 10684, "Q9"},
		{11, 10688, "Q10"},
	} {

		qf := model.QuestionFile{SourceID: s.ID, QuestionID: q.ID}
		db.Create(&qf)
		ps := model.ProblemSheet{ProblemID: p.ID, Name: r.name, SequenceNumber: r.sequence}
		db.Create(&ps)
		sheetID := r.sheetID - u.ID
		qs := model.QuestionFileSheet{ID: sheetID, Sequence: r.sequence, Name: r.name, QuestionFileID: qf.ID, ProblemSheetID: ps.ID, ProblemID: p.ID}
		db.Create(&qs)
		xt := model.XLQTransformation{UserID: u.ID, QuestionID: q.ID, SourceID: s.ID, QuestionFileID: qf.ID}
		db.Create(&xt)

	}
	entires, err := q.GetGAEntries(file, u.ID)
	if err != nil {
		t.Fatal(err)
	}
	if entires == nil {
		t.Fatal("Expected to get a populated map with GA data entries")
	}

	if expected, count := 10, len(entires); count != expected {
		t.Errorf("Expected %d entries, got: %d", expected, count)
	}
}
