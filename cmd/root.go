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
	"fmt"
	"os"
	"path"
	"regexp"
	"strings"
	"time"

	"github.com/jinzhu/gorm"

	log "github.com/Sirupsen/logrus"
	homedir "github.com/mitchellh/go-homedir"
	"github.com/spf13/cobra"
	"github.com/spf13/viper"
	//"github.com/tealeg/xlsx"
	"github.com/nad2000/xlsx"
)

const (
	defaultColor = "FFFFFF00"
	defaultURL   = "sqlite://blocks.db"
)

var (
	awsRegion          string
	awsAccessKeyID     string
	awsSecretAccessKey string
	cfgFile            string
	// Db - shared DB connection
	Db       *gorm.DB
	debug    bool
	verbose  bool
	testing  bool
	force    bool
	color    string
	cellIDRe = regexp.MustCompile("\\$?[A-Z]+\\$?[0-9]+")
)

func cellAddress(rowIndex, colIndex int) string {
	return xlsx.GetCellIDStringFromCoords(colIndex, rowIndex)

}

// RelativeCellAddress converts cell ID into a relative R1C1 representation
func RelativeCellAddress(rowIndex, colIndex int, cellID string) string {
	x, y, err := xlsx.GetCoordsFromCellIDString(cellID)
	if err != nil {
		log.Fatalf("Failed to find coordinates for %q: %s", cellID, err.Error())
	}
	var r1c1 string

	if strings.Contains(cellID[1:], "$") {
		r1c1 = fmt.Sprintf("R[%d]", y)
	} else {
		r1c1 = fmt.Sprintf("R[%+d]", y-rowIndex)
	}

	if cellID[0] == '$' {
		r1c1 += fmt.Sprintf("C[%d]", x)
	} else {
		r1c1 += fmt.Sprintf("C[%+d]", x-colIndex)
	}
	//return fmt.Sprintf("R[%d]C[%d]", y-rowIndex, x-colIndex)
	return r1c1
}

// RelativeFormula transforms the cell formula into the relative in R1C1 notation
func RelativeFormula(rowIndex, colIndex int, formula string) string {
	cellIDs := cellIDRe.FindAllString(formula, -1)
	for _, cellID := range cellIDs {
		relCellID := RelativeCellAddress(rowIndex, colIndex, cellID)
		log.Debugf("Replacing %q with %q at (%d, %d)", cellID, relCellID, rowIndex, colIndex)
		formula = strings.Replace(formula, cellID, relCellID, -1)
	}
	return formula
}

// Source - student answer file sources
type Source struct {
	ID           int    `gorm:"column:FileID;primary_key;AUTO_INCREMENT"`
	S3BucketName string `gorm:"column:S3BucketName"`
	S3Key        string `gorm:"column:S3Key"`
	FileName     string `gorm:"column:FileName"`
	ContentType  string `gorm:"column:ContentType"`
	FileSize     int    `gorm:"column:FileSize"`
	Answers      []Answer
}

// TableName overrides default table name for the model
func (Source) TableName() string {
	return "FileSources"
}

// Answer - student submitted answers
type Answer struct {
	ID             int         `gorm:"column:StudentAnswerID;primary_key;AUTO_INCREMENT"`
	AssignmentID   int         `gorm:"column:StudentAssignmentID"`
	QuestionID     int         `gorm:"column:QuestionID"`
	MCQOptionID    int         `gorm:"column:MCQOptionID"`
	ShortAnswer    string      `gorm:"column:ShortAnswerText"`
	Marks          string      `gorm:"column:Marks"`
	SubmissionTime time.Time   `gorm:"column:SubmissionTime"`
	FileID         int         `gorm:"column:FileID"`
	Worksheets     []Worksheet `gorm:"ForeignKey:StudentAnswerID;AssociationForeignKey:Refer"`
	Source         Source      `gorm:"ForeignKey:FileID"`
}

// TableName overrides default table name for the model
func (Answer) TableName() string {
	return "StudentAnswers"
}

// Workbook - Excel file / workbook
type Workbook struct {
	ID         int
	FileName   string
	CreatedAt  time.Time
	Worksheets []Worksheet `gorm:"ForeignKey:WorkbookID;AssociationForeignKey:Refer"`
}

// reset deletes all underlying objects: worksheets, blocks, and cells
func (wb *Workbook) reset() {

	var worksheets []Worksheet
	Db.Model(&wb).Related(&worksheets)
	log.Debugf("Deleting worksheets: %#v", worksheets)
	for ws := range worksheets {
		var blocks []Block
		Db.Model(&ws).Related(&blocks)
		for _, b := range blocks {
			log.Debugf("Deleting blocks: %#v", blocks)
			Db.Where("block_id = ?", b.ID).Delete(Cell{})
			Db.Delete(b)
		}
	}

	Db.Where("workbook_id = ?", wb.ID).Delete(Worksheet{})
}

// Worksheet - Excel workbook worksheet
type Worksheet struct {
	ID               int
	WorkbookID       int `gorm:"index"`
	Name             string
	WorkbookFileName string
	Blocks           []Block `gorm:"ForeignKey:worksheet_id;AssociationForeignKey:Refer"`
}

// Block - the univormly filled with specific color block
type Block struct {
	ID              int `gorm:"column:ExcelBlockID;primary_key;AUTO_INCREMENT"`
	WorksheetID     int `gorm:"index"`
	Color           string
	Range           string `gorm:"column:BlockCellRange"`
	Formula         string `gorm:"column:BlockFormula"` // first block cell formula
	RelativeFormula string // first block cell relative formula formula
	Cells           []Cell `gorm:"ForeignKey:block_id;AssociationForeignKey:Refer"`

	s struct{ r, c int } `gorm:"-"` // Top-left cell
	e struct{ r, c int } `gorm:"-"` //  Bottom-right cell
}

func (b Block) String() string {
	return fmt.Sprintf("Block {Range: %q, Color: %q, Formula: %q, Relative Formula: %q}",
		b.Range, b.Color, b.Formula, b.RelativeFormula)
}

// TableName overrides default table name for the model
func (b Block) TableName() string {
	return "ExcelBlocks"
}

func (b *Block) save() {
	b.Range = b.address()
	Db.Save(b)
}

func (b *Block) address() string {
	return cellAddress(b.s.r, b.s.c) + ":" + cellAddress(b.e.r, b.e.c)
}

//  getCellComment returns cell comment text value
func getCellComment(file *xlsx.File, cellID string) string {
	if file.Comments != nil {
		for _, c := range file.Comments {
			if cellID == c.Ref {
				return c.Text
			}
		}
	}
	return ""
}

// fildWhole finds whole range of the specified color
// and the same "relative" formula starting with the set top-left cell.
func (b *Block) findWhole(sheet *xlsx.Sheet, color string) {

	b.e = b.s
	for i, row := range sheet.Rows {

		// skip all rows until the first block row
		if i < b.s.r {
			continue
		}

		log.Debugf("Total cells: %d at %d", len(row.Cells), i)
		// Range is discontinued or of a differnt color
		//log.Infof("*** b.e.c: %d, len: %d, %#v", b.e.c, len(row.Cells), row.Cells)
		if len(row.Cells) <= b.e.c ||
			row.Cells[b.e.c].GetStyle().Fill.FgColor != color ||
			RelativeFormula(i, b.e.c, row.Cells[b.e.c].Formula()) != b.RelativeFormula {
			log.Debugf("Reached the edge row of the block at row %d", i)
			b.e.r = i - 1
			break
		} else {
			b.e.r = i
		}

		for j, cell := range row.Cells {
			// skip columns until the start:
			if j < b.s.c {
				continue
			}

			fgColor := cell.GetStyle().Fill.FgColor
			relFormula := RelativeFormula(i, j, cell.Formula())
			// Reached the top-right corner:
			if fgColor == color && relFormula == b.RelativeFormula {
				cellID := cellAddress(i, j)
				commentText := ""
				comment, ok := sheet.Comment[cellID]
				if ok {
					commentText = comment.Text
				}
				c := Cell{
					BlockID: b.ID,
					Formula: cell.Formula(),
					Value:   cell.Value,
					Range:   cellID,
					Comment: commentText,
				}
				Db.Create(&c)
				b.e.c = j
			} else {
				log.Debugf("Reached the edge column  of the block at column %d", j)
				if j > b.e.c {
					b.e.c = j - 1
				}
				break
			}
		}
	}
}

func (b *Block) isInside(r, c int) bool {
	return (b.s.r <= r &&
		r <= b.e.r &&
		b.s.c <= c &&
		c <= b.e.c)
}

type blockList []Block

// alreadyFound tests if the range containing the cell
// coordinates hhas been already found.
func (bl *blockList) alreadyFound(r, c int) bool {
	for _, b := range *bl {
		if b.isInside(r, c) {
			return true
		}
	}
	return false
}

// Cell - a sigle cell of the block
type Cell struct {
	ID      int
	BlockID int `gorm:"index"`
	Range   string
	Formula string
	Value   string
	Comment string
}

// RootCmd represents the base command when called without any subcommands
var RootCmd = &cobra.Command{
	Use:   "extract-blocks",
	Short: "Extracts Cell Formula Blocks from Excel file and writes to MySQL",
	Long: `Extracts Cell Formula Blocks from Excel file and writes to MySQL.
	
Conditions that define Cell Formula Block - 
    (i) Any contiguous (unbroken) range of excel cells containing cell formula
   (ii) Contiguous cells could be either in a row or in a column or in row+column cell block.
  (iii) The formula in the range of cells should be the same except the changes due to relative cell references.
  
Connection should be defined using connection URL notation: DRIVER://CONNECIONT_PARAMETERS, 
where DRIVER is either "mysql" or "sqlite", e.g., mysql://user:password@/dbname?charset=utf8&parseTime=True&loc=Local.
More examples on connection parameter you can find at: https://github.com/go-sql-driver/mysql#examples.`,
	Run: extractBlocks,
}

// SetDb initializes DB
func SetDb() {
	// Migrate the schema
	log.Debug("Add to automigrate...")
	Db.AutoMigrate(&Source{})
	Db.AutoMigrate(&Answer{})
	Db.AutoMigrate(&Workbook{})
	Db.AutoMigrate(&Worksheet{})
	Db.AutoMigrate(&Block{})
	Db.AutoMigrate(&Cell{})
	if strings.HasPrefix(Db.Dialect().GetName(), "mysql") {
		Db.Model(&Worksheet{}).AddForeignKey("StudentAnswerID", "worksheets(StudentAnswerID)", "CASCADE", "CASCADE")
		Db.Model(&Cell{}).AddForeignKey("block_id", "ExcelBlocks(ExcelBlockID)", "CASCADE", "CASCADE")
		Db.Model(&Block{}).AddForeignKey("worksheet_id", "worksheets(id)", "CASCADE", "CASCADE")
		Db.Model(&Worksheet{}).AddForeignKey("workbook_id", "workbooks(id)", "CASCADE", "CASCADE")
	}
}

// RowsToProcessResult stores query resut
type RowsToProcessResult struct {
	FileID          int    `gorm:"column:FileID"`
	S3BucketName    string `gorm:"column:S3BucketName"`
	S3Key           string `gorm:"column:S3Key"`
	FileName        string `gorm:"column:FileName"`
	StudentAnswerID int    `gorm:"column:StudentAnswerID"`
}

// RowsToProcess returns answer file sources
func RowsToProcess() ([]RowsToProcessResult, error) {

	currentTime := time.Now()
	midnight := time.Date(
		currentTime.Year(),
		currentTime.Month(),
		currentTime.Day(),
		0, 0, 0, 0, time.UTC)

	// TODO: select file links from StudentAnswers and download them form S3 buckets..."
	rows, err := Db.Table("FileSources").Select(
		"FileSources.FileID, S3BucketName, S3Key, FileName, StudentAnswerID").Joins(
		"JOIN StudentAnswers ON StudentAnswers.FileID = FileSources.FileID").Where(
		"FileName IS NOT NULL").Where(
		"FileName != ''").Where(
		"FileName LIKE '%.xlsx'").Where(
		"SubmissionTime <= ?", midnight).Rows()
	defer rows.Close()

	if err != nil {
		return nil, err
	}

	var results []RowsToProcessResult
	for rows.Next() {
		var r RowsToProcessResult
		Db.ScanRows(rows, &r)
		results = append(results, r)
	}

	return results, nil
}

func extractBlocks(cmd *cobra.Command, args []string) {

	debugCmd(cmd)
	var err error
	testing = flagBool(cmd, "test")
	force = flagBool(cmd, "force")
	color = flagString(cmd, "color")
	profile := flagString(cmd, "aws-profile")
	region := flagString(cmd, "aws-region")

	url := flagString(cmd, "url")
	parts := strings.Split(flagString(cmd, "url"), "://")
	if len(parts) < 2 {
		log.Warnf("Driver name not given in %q, assuming 'mysql'.", url)
		parts = []string{"mysql", parts[0]}
	}

	switch parts[0] {
	case "sqlite", "sqlite3":
		log.Debugf("Connecting to Sqlite3 DB: %q.", parts[1])
		parts[0] = "sqlite3"
	case "mysql":
		log.Debugf("Connecting to MySQL DB: %q.", parts[1])
	default:
		log.Fatalf("Unsupported driver: %q. It should be either 'mysql' or 'sqlite'.", parts[0])
	}
	Db, err = gorm.Open(parts[0], parts[1])
	if err != nil {
		log.Error(err)
		log.Fatalf("failed to connect database %q", url)
	}
	defer Db.Close()
	SetDb()
	//db.LogMode(true)

	if testing {
		// read up the file list from the arguments
		for _, excelFileName := range args {
			extractBlocksFromFile(excelFileName)
		}
	} else {
		downloader := NewS3Downloader(region, profile)
		HandleAnswers(downloader)
	}
}

// HandleAnswers - iterates through student answers and retrievs answer workbooks
// it thaks the funcion that actuatualy performs file download from S3 bucket
// and returns the downloades file name or an error.
func HandleAnswers(downloader FileDownloader) error {

	rows, err := RowsToProcess()
	if err != nil {
		log.Fatalf("Failed to retrieve list of source files to process: %s", err.Error())
	}
	var fileCount int
	for _, r := range rows {
		destinationName := path.Join(os.TempDir(), r.FileName)
		log.Infof(
			"Downloading %q (%q) form %q into %q",
			r.S3Key, r.FileName, r.S3BucketName, destinationName)
		fileName, err := downloader.DownloadFile(
			r.FileName, r.S3BucketName, r.S3Key, destinationName)
		if err != nil {
			log.Errorf(
				"Failed to retrieve file %q from %q into %q: %s",
				r.S3Key, r.S3BucketName, destinationName, err.Error())
			continue
		}
		log.Infof("Processing %q", fileName)
		extractBlocksFromFile(fileName)
		fileCount++
	}
	log.Infof("Downloaded and loaded %d Excel files.", fileCount)
	if len(rows) != fileCount {
		log.Infof("Failed to download and load %d file(s)", len(rows)-fileCount)
	}
	return nil
}

func extractBlocksFromFile(fileName string) (wb Workbook) {
	xlFile, err := xlsx.OpenFile(fileName)
	if err != nil {
		log.Fatal(err)
	}

	result := Db.First(&wb, Workbook{FileName: fileName})
	if !result.RecordNotFound() {
		if !force {
			log.Errorf("File %q was already processed.", fileName)
			return
		}
		log.Warnf("File %q was already processed.", fileName)
		wb.reset()
	} else {
		wb = Workbook{FileName: fileName}
		Db.Create(&wb)
	}

	if verbose {
		log.Infof("*** Processing workbook: %s", fileName)
	}

	for _, sheet := range xlFile.Sheets {

		if sheet.Hidden {
			log.Infof("Skipping hidden worksheet %q", sheet.Name)
			continue
		}

		if verbose {
			log.Infof("Processing worksheet %q", sheet.Name)
		}

		var ws Worksheet
		Db.FirstOrCreate(&ws, Worksheet{
			Name:             sheet.Name,
			WorkbookID:       wb.ID,
			WorkbookFileName: wb.FileName,
		})
		blocks := blockList{}
		sheetFillColors := []string{}

		for i, row := range sheet.Rows {
			for j, cell := range row.Cells {

				if blocks.alreadyFound(i, j) {
					continue
				}
				style := cell.GetStyle()
				fgColor := style.Fill.FgColor
				if fgColor != "" {
					for _, c := range sheetFillColors {
						if c == fgColor {
							goto MATCH
						}
					}
					sheetFillColors = append(sheetFillColors, fgColor)
				}
			MATCH:

				if fgColor == color {

					b := Block{
						WorksheetID:     ws.ID,
						Color:           color,
						Formula:         cell.Formula(),
						RelativeFormula: RelativeFormula(i, j, cell.Formula()),
					}
					b.s.r, b.s.c = i, j

					Db.Create(&b)

					b.findWhole(sheet, color)
					b.save()
					blocks = append(blocks, b)
					if verbose {
						log.Infof("Found: %s", b)
					}

				}
			}
		}
		if len(blocks) == 0 {
			log.Warningf("No block found ot the worksheet %q of the workbook %q with color %q", sheet.Name, fileName, color)
			if len(sheetFillColors) > 0 {
				log.Infof("Following colors were found in the worksheet you could use: %v", sheetFillColors)
			}
		}
	}
	return
}

// Execute adds all child commands to the root command and sets flags appropriately.
// This is called by main.main(). It only needs to happen once to the rootCmd.
func Execute() {

	if err := RootCmd.Execute(); err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}

func init() {
	cobra.OnInitialize(initConfig)
	RootCmd.PersistentFlags().StringVar(&cfgFile, "config", "", "config file (default is $HOME/.extract-blocks.yaml)")
	RootCmd.Flags().BoolP("test", "t", false, "Run in testing ignoring 'StudentAnswers'.")
	RootCmd.PersistentFlags().BoolP("debug", "d", false, "Show full stack trace on error.")
	RootCmd.PersistentFlags().BoolP("verbose", "v", false, "Verbose mode. Produce more output about what the program does.")
	RootCmd.PersistentFlags().StringP("url", "U", defaultURL, "Database URL connection string, e.g., mysql://user:password@/dbname?charset=utf8&parseTime=True&loc=Local (More examples at: https://github.com/go-sql-driver/mysql#examples).")
	RootCmd.PersistentFlags().BoolP("force", "f", false, "Repeat extraction if files were already handle.")
	RootCmd.PersistentFlags().StringP("color", "c", defaultColor, "The block filling color.")

	RootCmd.PersistentFlags().String("aws-profile", "default", "AWS Configuration Profile (see: http://docs.aws.amazon.com/cli/latest/userguide/cli-chap-getting-started.html)")
	RootCmd.PersistentFlags().String("aws-region", "ap-south-1", "AWS Region.")
	// RootCmd.PersistentFlags().String("aws-access-key-id", "", "AWS Access Key ID.")
	// RootCmd.PersistentFlags().String("aws-secret-access-key", "", "AWS Secret Access Key.")

	viper.BindPFlag("url", RootCmd.PersistentFlags().Lookup("url"))
	viper.BindPFlag("color", RootCmd.PersistentFlags().Lookup("color"))
	viper.BindPFlag("force", RootCmd.PersistentFlags().Lookup("force"))
	viper.BindPFlag("aws-profile", RootCmd.PersistentFlags().Lookup("aws-profile"))
	viper.BindPFlag("aws-region", RootCmd.PersistentFlags().Lookup("aws-region"))
	viper.BindEnv("aws-region", "AWS_REGION")
	// viper.BindPFlag("aws-access-key-id", RootCmd.PersistentFlags().Lookup("aws-access-key-id"))
	// viper.BindPFlag("aws-secret-access-key", RootCmd.PersistentFlags().Lookup("aws-secret-access-key"))

}

// initConfig reads in config file and ENV variables if set.
func initConfig() {
	if cfgFile != "" {
		// Use config file from the flag.
		viper.SetConfigFile(cfgFile)
	} else {
		// Find home directory.
		home, err := homedir.Dir()
		if err != nil {
			log.Fatal(err)
		}

		// Search config in home directory with name ".extract-blocks" (without extension).
		viper.AddConfigPath(home)
		viper.SetConfigName(".extract-blocks")
	}

	viper.AutomaticEnv() // read in environment variables that match

	// If a config file is found, read it in.
	if err := viper.ReadInConfig(); err == nil {
		log.Info("Using config file:", viper.ConfigFileUsed())
	}
}

func flagString(cmd *cobra.Command, name string) string {

	value := cmd.Flag(name).Value.String()
	if value != "" {
		return value
	}
	conf := viper.Get(name)
	if conf == nil {
		return ""
	}
	return conf.(string)
}

func flagStringSlice(cmd *cobra.Command, name string) (val []string) {
	val, err := cmd.Flags().GetStringSlice(name)
	if err != nil {
		log.Fatal(err)
	}
	return
}

func flagStringArray(cmd *cobra.Command, name string) (val []string) {
	val, err := cmd.Flags().GetStringArray(name)
	if err != nil {
		log.Fatal(err)
	}
	return
}

func flagBool(cmd *cobra.Command, name string) (val bool) {
	val, err := cmd.Flags().GetBool(name)
	if err != nil {
		log.Fatal(err)
	}
	return
}

func flagInt(cmd *cobra.Command, name string) (val int) {
	val, err := cmd.Flags().GetInt(name)
	if err != nil {
		log.Fatal(err)
	}
	return
}

func debugCmd(cmd *cobra.Command) {
	debug = flagBool(cmd, "debug")
	verbose = flagBool(cmd, "verbose")

	if debug {
		log.SetLevel(log.DebugLevel)
		title := fmt.Sprintf("Command %q called with flags:", cmd.Name())
		log.Info(title)
		log.Info(strings.Repeat("=", len(title)))
		cmd.DebugFlags()
	}
}
