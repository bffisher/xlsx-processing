package util

import (
	"testing"
)

func Test_Axis(t *testing.T) {
	if r := Axis(0, 0); r != "A1" {
		t.Fatal("Expect A1, result:", r)
	}

	if r := Axis(134, 3); r != "D135" {
		t.Fatal("Expect D135, result:", r)
	}

	if r := Axis(5, 25); r != "Z6" {
		t.Fatal("Expect Z6, result:", r)
	}

	if r := Axis(223, 28); r != "AC224" {
		t.Fatal("Expect AC224, result:", r)
	}

	if r := Axis(223, 51); r != "AZ224" {
		t.Fatal("Expect AZ224, result:", r)
	}
}
