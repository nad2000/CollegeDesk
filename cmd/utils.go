package cmd

import (
	"extract-blocks/s3"
	"fmt"
	"strings"

	log "github.com/Sirupsen/logrus"
	"github.com/spf13/cobra"
	"github.com/spf13/viper"
)

func createS3Downloader() (d s3.Downloader) {
	if awsAccessKeyID == "" && awsProfile != "" || awsProfile != "default" {
		return s3.NewDownloader(awsRegion, awsProfile)
	} else if awsAccessKeyID != "" && awsSecretAccessKey != "" {
		return s3.NewDownloaderWithCredentials(
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
