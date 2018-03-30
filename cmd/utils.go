package cmd

import (
	"crypto/rand"
	"extract-blocks/s3"
	"fmt"
	"io"
	"strings"

	log "github.com/Sirupsen/logrus"
	"github.com/spf13/cobra"
	"github.com/spf13/viper"
)

func createS3Manager() (d s3.Manager) {
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

// newUUID generates a random UUID according to RFC 4122
func newUUID() (string, error) {
	uuid := make([]byte, 16)
	n, err := io.ReadFull(rand.Reader, uuid)
	if n != len(uuid) || err != nil {
		return "", err
	}
	// variant bits; see section 4.1.1
	uuid[8] = uuid[8]&^0xc0 | 0x80
	// version 4 (pseudo-random); see section 4.1.3
	uuid[6] = uuid[6]&^0xf0 | 0x40
	return fmt.Sprintf("%x-%x-%x-%x-%x", uuid[0:4], uuid[4:6], uuid[6:8], uuid[8:10], uuid[10:]), nil
}
