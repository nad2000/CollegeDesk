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
	"os"
	"strings"

	model "extract-blocks/model"

	log "github.com/Sirupsen/logrus"
	"github.com/jinzhu/gorm"
	homedir "github.com/mitchellh/go-homedir"
	"github.com/spf13/cobra"
	"github.com/spf13/viper"
	//"github.com/tealeg/xlsx"
)

const (
	defaultColor = "FFFFFF00"
	defaultURL   = "sqlite://blocks.db"
)

// Db - shared DB connection
var Db *gorm.DB
var (
	awsAccessKeyID         string
	awsProfile             string
	awsRegion              string
	awsSecretAccessKey     string
	cfgFile                string
	color                  = defaultColor
	debug                  bool
	skipHidden             bool
	debugLevel             int
	dest                   string
	force                  bool
	testing                bool
	url                    string
	verbose                bool
	verboseLevel           int
	isPlagiarisedCommentID int
	modelAnswerUserID      int
)

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
where DRIVER is either "mysql" or "sqlite", e.g., mysql://user:password@tcp(SERVER:PORT)/dbname?charset=utf8&parseTime=True&loc=Local.
More examples on connection parameter you can find at: https://github.com/go-sql-driver/mysql#examples.`,
	// Run: func(c *cobra.Command, args []Args) {},
}

func getConfig() {
	awsProfile = viper.GetString("aws-profile")
	awsRegion = viper.GetString("aws-region")
	awsAccessKeyID = viper.GetString("aws-access-key-id")
	awsSecretAccessKey = viper.GetString("aws-secret-access-key")
	url = viper.GetString("url")
	color = viper.GetString("color")
	force = viper.GetBool("force")
	dest = viper.GetString("dest")
	skipHidden = viper.GetBool("skip-hidden")
	isPlagiarisedCommentID = viper.GetInt("is-plagiarised-comment-id")
	modelAnswerUserID = viper.GetInt("model-answer-user-id")
	if !strings.HasPrefix(dest, "/") {
		dest += "/"
	}
}

// Execute adds all child commands to the root command and sets flags appropriately.
// This is called by main.main(). It only needs to happen once to the rootCmd.
func Execute() {

	if err := RootCmd.Execute(); err != nil {
		log.Error(err)
		os.Exit(1)
	}
}

func init() {
	cobra.OnInitialize(initConfig)
	flags := RootCmd.PersistentFlags()
	flags.StringVar(&cfgFile, "config", "", "config file (default is $HOME/.extract-blocks.yaml)")
	flags.BoolVarP(&testing, "test", "t", false, "Run in testing ignoring 'StudentAnswers'.")
	flags.BoolVarP(&skipHidden, "skip-hidden", "", false, "Skip hidden worksheets.")
	flags.CountVarP(&debugLevel, "debug", "d", "Show full stack trace on error.")
	flags.CountVarP(&verboseLevel, "verbose", "v", "Verbose mode. Produce more output about what the program does.")
	flags.BoolVarP(&model.DryRun, "dry", "D", false, "Dry run, run commands without performing and DB update or file changes.")
	flags.StringP("url", "U", defaultURL, "Database URL connection string, e.g., mysql://user:password@/dbname?charset=utf8&parseTime=True&loc=Local (More examples at: https://github.com/go-sql-driver/mysql#examples).")
	flags.String("aws-profile", "default", "AWS Configuration Profile (see: http://docs.aws.amazon.com/cli/latest/userguide/cli-chap-getting-started.html)")
	flags.String("aws-region", "ap-south-1", "AWS Region.")
	flags.String("aws-access-key-id", "", "AWS Access Key ID.")
	flags.String("aws-secret-access-key", "", "AWS Secret Access Key.")
	flags.String("dest", os.TempDir(), "The destionation directory for download files from AWS S3.")

	viper.BindPFlag("url", flags.Lookup("url"))
	viper.BindPFlag("aws-profile", flags.Lookup("aws-profile"))
	viper.BindPFlag("aws-region", flags.Lookup("aws-region"))
	viper.BindPFlag("aws-access-key-id", flags.Lookup("aws-access-key-id"))
	viper.BindPFlag("aws-secret-access-key", flags.Lookup("aws-secret-access-key"))
	viper.BindEnv("aws-region", "AWS_REGION")
	viper.BindEnv("aws-access-key-id", "AWS_ACCESS_KEY_ID")
	viper.BindEnv("aws-secret-access-key", "AWS_SECRET_ACCESS_KEY")
	viper.SetDefault("aws-region", "ap-south-1")
	viper.SetDefault("dest", os.TempDir())
	viper.SetDefault("is-plagiarised-comment-id", 12345)
	viper.SetDefault("model-answer-user-id", 10000)
	viper.SetDefault("skip-hidden", false)
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
		viper.AddConfigPath(".")
		viper.SetConfigType("yaml")
		viper.SetConfigName(".extract-blocks")
	}

	viper.AutomaticEnv() // read in environment variables that match

	// If a config file is found, read it in.
	if err := viper.ReadInConfig(); err == nil {
		log.Info("Using config file:", viper.ConfigFileUsed())
	} else {
		if !strings.Contains(err.Error(), "Not Found") {
			log.Error(err)
		}
	}
}
