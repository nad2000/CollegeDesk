package utils

import (
	"crypto/rand"
	"encoding/hex"
	"fmt"
	"io"
	rnd "math/rand"
	"os"
	"path/filepath"
	"time"
)

func init() {
	rnd.Seed(time.Now().UnixNano())
}

// NewUUID generates a random UUID according to RFC 4122
func NewUUID() (string, error) {
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

const validS3KeyCharacters = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-"
const s3KeyLength = 8

// NewS3Key generates a new S3 Key value
func NewS3Key() string {
	b := make([]byte, s3KeyLength)
	for i := range b {
		b[i] = validS3KeyCharacters[rnd.Intn(len(validS3KeyCharacters))]
	}
	return string(b)
}

// TempFileName generates a temporary filename for use in testing or whatever
func TempFileName(prefix, suffix string) string {
	randBytes := make([]byte, 8)
	rnd.Read(randBytes)
	return filepath.Join(os.TempDir(), prefix+hex.EncodeToString(randBytes)+suffix)
}
