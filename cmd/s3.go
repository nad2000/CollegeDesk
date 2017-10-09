// Package cmd imlements submitted student answer file processing.
//
// Uses the default AWS SDK Credentials;  e.g. via the environment
// AWS_REGION=region AWS_ACCESS_KEY_ID=key AWS_SECRET_ACCESS_KEY=secret
// OR in the AWS SDK credential configurtion ~/.aws/credentials:
//
// aws_access_key_id = AKID
// aws_secret_access_key = SECRET
// aws_session_token = TOKEN
//
// See: https://docs.aws.amazon.com/sdk-for-go/api/aws/session/#pkg-overview
package cmd

import (
	"os"

	log "github.com/Sirupsen/logrus"
	"github.com/aws/aws-sdk-go/aws"
	"github.com/aws/aws-sdk-go/aws/session"
	"github.com/aws/aws-sdk-go/service/s3"
	"github.com/aws/aws-sdk-go/service/s3/s3manager"
)

// FileDownloader generic file downloader interfacae
type FileDownloader interface {
	DownloadFile(fileName, containerName, sourceName, destinationName string) (string, error)
}

// S3Downloader AWS S3 file downloader
type S3Downloader struct {
	s3Downloader *s3manager.Downloader
}

func (d *S3Downloader) setUp(region, profile string) {
	if profile == "" {
		profile = "default"
	}
	log.Debugf("Using region: %q, profile: %q", region, profile)
	sess := session.Must(session.NewSessionWithOptions(
		session.Options{
			Profile: profile,
			Config:  aws.Config{Region: aws.String(region)},
		}))
	d.s3Downloader = s3manager.NewDownloader(sess)
}

// NewS3Downloader instantiates an AWS S3 file downloader
func NewS3Downloader(region, profile string) S3Downloader {
	d := S3Downloader{}
	d.setUp(region, profile)
	return d
}

// DownloadFile downloads a file form the given bucket to the destination file.
func (d S3Downloader) DownloadFile(
	SourceName, S3BucketName, S3Key, DestinationFileName string) (string, error) {

	f, err := os.Create(DestinationFileName)
	if err != nil {
		return "", err
	}
	defer f.Close()

	numBytes, err := d.s3Downloader.Download(f,
		&s3.GetObjectInput{
			Bucket: aws.String(S3BucketName),
			Key:    aws.String(S3Key),
		})
	if err != nil {
		return "", err
	}

	log.Debug("Downloaded file", f.Name(), numBytes, "bytes")
	return DestinationFileName, nil
}
