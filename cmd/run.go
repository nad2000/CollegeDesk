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
	"github.com/spf13/cobra"
	"github.com/spf13/viper"
)

// runCmd represents the run command
var runCmd = &cobra.Command{
	Use:   "run",
	Short: "A brief description of your command",
	Long: `Retrievs list of the file sources for submitted answers, downloads the Excel 
workbooks and extracts Cell Formula Blocks from Excel file and writes to MySQL.
	
Conditions that define Cell Formula Block - 
    (i) Any contiguous (unbroken) range of excel cells containing cell formula
   (ii) Contiguous cells could be either in a row or in a column or in row+column cell block.
  (iii) The formula in the range of cells should be the same except the changes due to relative cell references.
  
Connection should be defined using connection URL notation: DRIVER://CONNECIONT_PARAMETERS, 
where DRIVER is either "mysql" or "sqlite", e.g., mysql://user:password@/dbname?charset=utf8&parseTime=True&loc=Local.
More examples on connection parameter you can find at: https://github.com/go-sql-driver/mysql#examples.`,
	Run: extractBlocks,
}

func init() {
	RootCmd.AddCommand(runCmd)
	flags := runCmd.Flags()

	flags.StringP("url", "U", defaultURL, "Database URL connection string, e.g., mysql://user:password@/dbname?charset=utf8&parseTime=True&loc=Local (More examples at: https://github.com/go-sql-driver/mysql#examples).")
	flags.BoolP("force", "f", false, "Repeat extraction if files were already handle.")
	flags.StringP("color", "c", defaultColor, "The block filling color.")

	viper.BindPFlag("url", flags.Lookup("url"))
	viper.BindPFlag("color", flags.Lookup("color"))
	viper.BindPFlag("force", flags.Lookup("force"))
}
