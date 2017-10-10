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
	"fmt"

	log "github.com/Sirupsen/logrus"
	"github.com/spf13/cobra"
)

// lsbCmd represents the lsb command
var lsbCmd = &cobra.Command{
	Use:   "lsb",
	Short: "List the content of a S3 bucket",
	Long:  `List the content of a S3 bucket.`,
	Run: func(cmd *cobra.Command, args []string) {

		debugCmd(cmd)
		bucket := flagString(cmd, "bucket")
		prefix := flagString(cmd, "prefix")

		downloader := NewS3Downloader(region, profile)
		list, err := downloader.List(bucket, prefix)
		if err != nil {
			log.Fatalf("Error occured listing %q: %s", bucket, err.Error())
		}

		if debug {
			fmt.Println("Key / Name\tOwner\tSize\tInternal Representation")
			fmt.Println("====================================================================")
		} else {
			fmt.Println("Key / Name\tOwner\tSize")
			fmt.Println("=============================================================")
		}

		for _, e := range list {
			if debug {
				fmt.Printf("%s\t%s\t%d\t%s\n", e.Name, e.Owner, e.Size, e.Repr)
			} else {
				fmt.Printf("%s\t%s\t%d\n", e.Name, e.Owner, e.Size)
			}
		}
	},
}

func init() {
	RootCmd.AddCommand(lsbCmd)

	lsbCmd.PersistentFlags().StringP("bucket", "b", "collegedesk", "S3 bucket")
	lsbCmd.PersistentFlags().StringP("prefix", "p", "", "Entry key prefix")
}
