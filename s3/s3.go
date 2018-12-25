// Package s3 imlements wrappers for AWS S3 service
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
package s3

import (
	"os"

	log "github.com/Sirupsen/logrus"
	"github.com/aws/aws-sdk-go/aws"
	"github.com/aws/aws-sdk-go/aws/credentials"
	"github.com/aws/aws-sdk-go/aws/session"
	"github.com/aws/aws-sdk-go/service/s3"
	"github.com/aws/aws-sdk-go/service/s3/s3manager"
)

// FileManager generic file downloader interfacae
type FileManager interface {
	Download(fileName, containerName, sourceName, destinationName string) (string, error)
	Upload(fileName, containerName, key string) (string, error)
}

// Manager AWS S3 file downloader
type Manager struct {
	s3Downloader *s3manager.Downloader
	s3Uploader   *s3manager.Uploader
}

func (m *Manager) setUp(region, profile string) {
	if profile == "" {
		profile = "default"
	}
	log.Debugf("Using region: %q, profile: %q", region, profile)

	sess := session.Must(session.NewSessionWithOptions(
		session.Options{
			Profile: profile,
			Config:  aws.Config{Region: aws.String(region)},
		}))
	m.s3Downloader = s3manager.NewDownloader(sess)
	m.s3Uploader = s3manager.NewUploader(sess)
}

func newAwsSession(accessKeyID, secretAccessKey, region string) (*session.Session, error) {
	creds := credentials.NewStaticCredentials(accessKeyID, secretAccessKey, "")
	_, err := creds.Get()
	if err != nil {
		return nil, err
	}

	return session.NewSession(&aws.Config{
		Region:      aws.String(region),
		Credentials: creds,
	})
}

// NewManagerWithCredentials instantiates an AWS S3 file manager
func NewManagerWithCredentials(accessKeyID, secretAccessKey, region string) Manager {
	m := Manager{}
	sess, err := newAwsSession(accessKeyID, secretAccessKey, region)
	if err != nil {
		log.Errorln("Failed to connect to AWS: ", err.Error())
	}
	m.s3Downloader = s3manager.NewDownloader(sess)
	m.s3Uploader = s3manager.NewUploader(sess)
	return m
}

// NewManager instantiates an AWS S3 file Manager
func NewManager(region, profile string) Manager {
	m := Manager{}
	m.setUp(region, profile)
	return m
}

// Entry S3 entry returned by List
type Entry struct {
	Name, Owner, Repr string
	Size              int64
}

// List lists content of a S3 bucket
func (m Manager) List(
	bucket, prefix string) ([]Entry, error) {

	params := &s3.ListObjectsInput{
		Bucket: aws.String(bucket),
		Prefix: aws.String(prefix),
	}

	resp, err := m.s3Downloader.S3.ListObjects(params)
	if err != nil {
		return nil, err
	}

	list := make([]Entry, len(resp.Contents))
	for i, key := range resp.Contents {

		//log.Debugf("i=%d, %#v", i, key)
		list[i].Name = *key.Key
		list[i].Size = *key.Size
		list[i].Repr = key.String()
		owner := key.Owner
		if owner != nil {
			if owner.DisplayName != nil {
				list[i].Owner = *owner.DisplayName
			} else if owner.ID != nil {
				list[i].Owner = *owner.ID
			}
		}
	}
	return list, nil
}

// Download downloads a file form the given bucket to the destination file.
func (m Manager) Download(
	SourceName, S3BucketName, S3Key, DestinationFileName string) (string, error) {

	f, err := os.Create(DestinationFileName)
	if err != nil {
		return "", err
	}
	defer f.Close()

	numBytes, err := m.s3Downloader.Download(f,
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

// Upload upload a file to the given bucket.
func (m Manager) Upload(fileName, bucket, key string) (string, error) {

	f, err := os.Open(fileName)
	if err != nil {
		return "", err
	}
	defer f.Close()
	result, err := m.s3Uploader.Upload(&s3manager.UploadInput{
		Bucket: aws.String(bucket),
		Key:    aws.String(key),
		Body:   f,
	})

	if err != nil {
		return "", err
	}
	log.Debugf("Uploaded file %q to %q", f.Name(), result.Location)
	return result.Location, nil
}
