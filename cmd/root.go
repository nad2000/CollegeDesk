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
	"strconv"
	"strings"

	"github.com/jinzhu/gorm"

	log "github.com/Sirupsen/logrus"
	homedir "github.com/mitchellh/go-homedir"
	"github.com/spf13/cobra"
	"github.com/spf13/viper"
	"github.com/tealeg/xlsx"
)

const defaultColor = "FFFFFF00"

// colCode - Excel column code
func colCode(colIndex int) string {

	if colIndex > 25 {
		return colCode(colIndex/26-1) + colCode(colIndex%26)
	}
	return string(colIndex + int('A'))
}

func cellAddress(rowIndex, colIndex int) string {
	return colCode(colIndex) + strconv.Itoa(rowIndex+1)
}

// Workbook - Excel file / workbook
type Workbook struct {
	ID         int
	FileName   string
	Worksheets []Worksheet `gorm:"ForeignKey:WorkbookID;AssociationForeignKey:Refer"`
}

// reset deletes all underlying objects: worksheets, blocks, and cells
func (wb *Workbook) reset() {

	var worksheets []Worksheet
	db.Model(&wb).Related(&worksheets)
	log.Debugf("Deleting worksheets: %#v", worksheets)
	for ws := range worksheets {
		var blocks []Block
		db.Model(&ws).Related(&blocks)
		for _, b := range blocks {
			log.Debugf("Deleting blocks: %#v", blocks)
			db.Where("block_id = ?", b.ID).Delete(Cell{})
			db.Delete(b)
		}
	}

	db.Where("workbook_id = ?", wb.ID).Delete(Worksheet{})
}

// Worksheet - Excel workbook worksheet
type Worksheet struct {
	ID               int
	WorkbookID       int `gorm:"index"`
	Name             string
	WorkbookFileName string
	Blocks           []Block
}

// Block - the univormly filled with specific color block
type Block struct {
	ID          int `gorm:"column:ExcelBlockID;primary_key;AUTO_INCREMENT"`
	WorksheetID int `gorm:"index"`
	Color       string
	Range       string `gorm:"column:BlockCellRange"`
	Formula     string `gorm:"column:BlockFormula"` // first block cell formula
	Cells       []Cell

	s struct{ r, c int } `gorm:"-"` // Top-left cell
	e struct{ r, c int } `gorm:"-"` //  Bottom-right cell
}

func (b Block) String() string {
	return fmt.Sprintf("Block {Range: %q, Color: %q, Formula: %q}", b.Range, b.Color, b.Formula)
}

func (b Block) TableName() string {
	return "ExcelBlocks"
}

func (b *Block) save() {
	b.Range = b.address()
	db.Save(b)
}

func (b *Block) address() string {
	return cellAddress(b.s.r, b.s.c) + ":" + cellAddress(b.e.r, b.e.c)
}

// fildWhole finds whole range of the specified color
// starting with the set top-left cell.
func (b *Block) findWhole(sheet *xlsx.Sheet, color string) {
	b.e = b.s
	for i, row := range sheet.Rows {

		// skip all rows until the first block row
		if i < b.s.r {
			continue
		}

		log.Debugf("Total cells: %d at %d", len(row.Cells), i)
		// Range is discontinued or of a differnt color
		if len(row.Cells) < b.e.c ||
			row.Cells[b.e.c].GetStyle().Fill.FgColor != color {
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
			// Reached the top-right corner:
			if fgColor == color {
				c := Cell{
					BlockID: b.ID,
					Formula: cell.Formula(),
					Value:   cell.Value,
					Range:   cellAddress(i, j),
				}
				db.Create(&c)
				b.e.c = j
			} else {
				log.Debugf("Reached the edge column  of the block at column %d", j)
				b.e.c = j - 1
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
}

var (
	cfgFile string
	db      *gorm.DB
	debug   bool
	verbose bool
)

// RootCmd represents the base command when called without any subcommands
var RootCmd = &cobra.Command{
	Use:   "extract-blocks",
	Short: "Extracts Cell Formula Blocks from Excel file and writes to MySQL",
	Long: `Extracts Cell Formula Blocks from Excel file and writes to MySQL.
	
Conditions that define Cell Formula Block - 
  (i) Any contiguous (unbroken) range of excel cells containing cell formula
  (ii) Contiguous cells could be either in a row or in a column or in row+column cell block.
  (iii) The formula in the range of cells should be the same except the changes due to relative cell references.`,
	Run: extractBlocks,
}

func SetDb(db *gorm.DB) {
	// Migrate the schema
	log.Debug("Add to automigrate...")
	db.AutoMigrate(&Workbook{})
	db.AutoMigrate(&Worksheet{})
	db.AutoMigrate(&Block{})
	db.AutoMigrate(&Cell{})
	if strings.HasPrefix(db.Dialect().GetName(), "myslq") {
		db.Model(&Cell{}).AddForeignKey("block_id", "ExcelBlocks(ExcelBlockID)", "CASCADE", "CASCADE")
		db.Model(&Block{}).AddForeignKey("worksheet_id", "worksheets(id)", "CASCADE", "CASCADE")
		db.Model(&Worksheet{}).AddForeignKey("workbook_id", "workbooks(id)", "CASCADE", "CASCADE")
	}
}

func extractBlocks(cmd *cobra.Command, args []string) {

	debugCmd(cmd)
	var err error

	mysql := flagString(cmd, "mysql")
	if mysql == "" {
		sqlite := flagString(cmd, "sqlite")
		db, err = gorm.Open("sqlite3", sqlite)
		log.Debugf("Connecting to Sqlite3 DB: %s", sqlite)
	} else {
		log.Debugf("Connecting to MySQL DB: %s", mysql)
		db, err = gorm.Open("mysql", mysql)
	}

	if err != nil {
		log.Error(err)
		log.Fatalf("failed to connect database %q", mysql)
	}
	defer db.Close()
	SetDb(db)
	//db.LogMode(true)

	color := flagString(cmd, "color")
	force := flagBool(cmd, "force")
	for _, excelFileName := range args {
		xlFile, err := xlsx.OpenFile(excelFileName)
		if err != nil {
			log.Error(err)
		}
		var wb Workbook
		result := db.FirstOrCreate(&wb, Workbook{FileName: excelFileName})
		// result := db.First(&wb, Workbook{FileName: excelFileName})

		if !result.RecordNotFound() {
			if !force {
				log.Errorf("File %q was already processed.", excelFileName)
				return
			} else {
				log.Warnf("File %q was already processed.", excelFileName)
				wb.reset()
			}
		}

		if verbose {
			log.Infof("*** Processing workbook: %s", excelFileName)
		}

		for _, sheet := range xlFile.Sheets {

			if verbose {
				log.Infof("Processing worksheet: %s", sheet.Name)
			}

			var ws Worksheet
			db.FirstOrCreate(&ws, Worksheet{
				Name:             sheet.Name,
				WorkbookID:       wb.ID,
				WorkbookFileName: wb.FileName,
			})
			blocks := blockList{}
			sheetFillColors := []string{}

			for i, row := range sheet.Rows {
				if debug {
					log.Printf("\n\nROW %d\n=========\n", i)
				}
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
							WorksheetID: ws.ID,
							Color:       color,
							Formula:     cell.Formula(),
						}
						b.s.r, b.s.c = i, j

						db.Create(&b)

						b.findWhole(sheet, color)
						b.save()
						blocks = append(blocks, b)
						if verbose {
							log.Infof("Found: %s", b)
						}

					}
				}
				if debug {
					log.Println()
				}
			}
			if len(blocks) == 0 {
				log.Warningf("No block found ot the worksheet %q of the workbook %q with color %q", sheet.Name, excelFileName, color)
				if len(sheetFillColors) > 0 {
					log.Infof("Following colors were found in the worksheet you could use: %v", sheetFillColors)
				}
			}
		}

	}
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
	RootCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")
	RootCmd.PersistentFlags().BoolP("debug", "d", false, "Show full stack trace on error.")
	RootCmd.PersistentFlags().BoolP("verbose", "v", false, "Verbose mode. Produce more output about what the program does.")
	RootCmd.PersistentFlags().StringP("mysql", "M", "", "MySQL connection string, e.g., https://github.com/go-sql-driver/mysql#examples.")
	RootCmd.PersistentFlags().StringP("sqlite", "S", "blocks.db", "Sqlite3 database file.")
	// RootCmd.PersistentFlags().StringP("user", "u", "", "The MySQL user name to use when connecting to the server.")
	// RootCmd.PersistentFlags().StringP("password", "p", "", "The password to use when connecting to the server.")

	// RootCmd.PersistentFlags().StringP("database", "D", "", "Database to use.")
	// RootCmd.PersistentFlags().StringP("host", "H", "", "Connect to host.")
	RootCmd.PersistentFlags().BoolP("force", "f", false, "Repeat extraction if files were already handle.")
	RootCmd.PersistentFlags().StringP("color", "c", defaultColor, "The block filling color.")
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
			fmt.Println(err)
			os.Exit(1)
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
