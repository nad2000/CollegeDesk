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
	"github.com/aws/aws-sdk-go/aws/credentials"
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

// NewS3DownloaderWithCredentials instantiates an AWS S3 file downloader
func NewS3DownloaderWithCredentials(accessKeyID, secretAccessKey, region string) S3Downloader {
	d := S3Downloader{}
	creds := credentials.NewStaticCredentials(accessKeyID, secretAccessKey, "")
	_, err := creds.Get()
	if err != nil {
		log.Fatalf("Bad AWS credentials: %s", err.Error())
	}
	cfg := aws.NewConfig().WithRegion(region).WithCredentials(creds)

	sess := session.Must(session.NewSessionWithOptions(session.Options{Config: *cfg}))
	d.s3Downloader = s3manager.NewDownloader(sess)
	return d
}

// NewS3Downloader instantiates an AWS S3 file downloader
func NewS3Downloader(region, profile string) S3Downloader {
	d := S3Downloader{}
	d.setUp(region, profile)
	return d
}

// S3Entry S3 entry returned by List
type S3Entry struct {
	Name, Owner, Repr string
	Size              int64
}

// List lists content of a S3 bucket
func (d S3Downloader) List(
	bucket, prefix string) ([]S3Entry, error) {

	params := &s3.ListObjectsInput{
		Bucket: aws.String(bucket),
		Prefix: aws.String(prefix),
	}

	resp, err := d.s3Downloader.S3.ListObjects(params)
	if err != nil {
		return nil, err
	}

	list := make([]S3Entry, len(resp.Contents))
	for i, key := range resp.Contents {

		list[i].Name = *key.Key
		list[i].Size = *key.Size
		list[i].Repr = key.String()
		owner := key.Owner
		if owner != nil {
			list[i].Owner = *owner.DisplayName
		}
	}
	return list, nil
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
