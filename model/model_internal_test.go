package model

import (
	"testing"
)

func TestNormalizeFloatRepr(t *testing.T) {
	if expected, got := "0.16", normalizeFloatRepr("0.16000000000000003"); got != expected {
		t.Errorf("Failed not normalizeFloatRepr, expected: %q, got: %q", expected, got)
	}
	if expected, got := "-0.08", normalizeFloatRepr("-8.0000000000000016E-2"); got != expected {
		t.Errorf("Failed not normalizeFloatRepr, expected: %q, got: %q", expected, got)
	}
}
