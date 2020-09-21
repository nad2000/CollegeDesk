package cmd

import (
	"extract-blocks/s3"
	"fmt"
	"strings"
	"time"

	log "github.com/Sirupsen/logrus"
	"github.com/jinzhu/now"
	"github.com/spf13/cobra"
	"github.com/spf13/viper"
)

func parseTime(str string) *time.Time {
	t := now.New(time.Now().UTC()).MustParse(str)
	return &t
}

func createManager() (d s3.FileManager) {
	if sourceDir != "" {
		if destinationDir == "" {
			destinationDir = sourceDir
		}
		return s3.LocalManager{SourceDirectory: sourceDir, DestinationDirctory: destinationDir}
	}
	if awsAccessKeyID == "" && awsProfile != "" || awsProfile != "default" {
		return s3.NewManager(awsRegion, awsProfile)
	} else if awsAccessKeyID != "" && awsSecretAccessKey != "" {
		return s3.NewManagerWithCredentials(
			awsAccessKeyID, awsSecretAccessKey, awsRegion)
	}
	log.Fatal("AWS credential information missing!")
	return
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

func debugCmd(cmd *cobra.Command) {

	if debugLevel > 0 {
		debug = true
	}

	if verboseLevel > 0 {
		verbose = true
	}

	if debug {
		log.SetLevel(log.DebugLevel)
		title := fmt.Sprintf("Command %q called with flags:", cmd.Name())
		log.Info(title)
		log.Info(strings.Repeat("=", len(title)))
		cmd.DebugFlags()
		viper.Debug()
	}
}
