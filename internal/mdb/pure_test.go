package mdb

import (
	"math"
	"testing"
)

// IsLongValue

func TestIsLongValue_memoIsLong(t *testing.T) {
	col := &Column{Type: ColTypeMemo}
	if !col.IsLongValue() {
		t.Error("expected Memo to be LongValue")
	}
}

func TestIsLongValue_oleIsLong(t *testing.T) {
	col := &Column{Type: ColTypeOLE}
	if !col.IsLongValue() {
		t.Error("expected OLE to be LongValue")
	}
}

func TestIsLongValue_textIsNotLong(t *testing.T) {
	col := &Column{Type: ColTypeText}
	if col.IsLongValue() {
		t.Error("expected Text to NOT be LongValue")
	}
}

func TestIsLongValue_longIsNotLong(t *testing.T) {
	col := &Column{Type: ColTypeLong}
	if col.IsLongValue() {
		t.Error("expected Long (integer) to NOT be LongValue")
	}
}

// ColTypeName

func TestColTypeName_knownTypes(t *testing.T) {
	cases := []struct {
		typ  byte
		want string
	}{
		{ColTypeBool, "Bool"},
		{ColTypeByte, "Byte"},
		{ColTypeInt, "Int"},
		{ColTypeLong, "Long"},
		{ColTypeMoney, "Money"},
		{ColTypeFloat, "Float"},
		{ColTypeDouble, "Double"},
		{ColTypeDatetime, "DateTime"},
		{ColTypeBinary, "Binary"},
		{ColTypeText, "Text"},
		{ColTypeOLE, "OLE"},
		{ColTypeMemo, "Memo"},
		{ColTypeGUID, "GUID"},
		{ColTypeNumeric, "Numeric"},
	}

	for _, tc := range cases {
		if got := ColTypeName(tc.typ); got != tc.want {
			t.Errorf("ColTypeName(%#x) = %q, want %q", tc.typ, got, tc.want)
		}
	}
}

func TestColTypeName_unknownType(t *testing.T) {
	got := ColTypeName(0xAB)
	if got == "" {
		t.Error("expected non-empty result for unknown type")
	}

	// Should contain the hex value
	if got == "Bool" || got == "Text" {
		t.Errorf("unknown type should not match a known name: %q", got)
	}
}

// intField

func TestIntField_int32(t *testing.T) {
	row := Row{"x": int32(42)}
	if got := intField(row, "x"); got != 42 {
		t.Errorf("expected 42, got %d", got)
	}
}

func TestIntField_int16(t *testing.T) {
	row := Row{"x": int16(100)}
	if got := intField(row, "x"); got != 100 {
		t.Errorf("expected 100, got %d", got)
	}
}

func TestIntField_int8(t *testing.T) {
	row := Row{"x": int8(7)}
	if got := intField(row, "x"); got != 7 {
		t.Errorf("expected 7, got %d", got)
	}
}

func TestIntField_uint32(t *testing.T) {
	row := Row{"x": uint32(999)}
	if got := intField(row, "x"); got != 999 {
		t.Errorf("expected 999, got %d", got)
	}
}

func TestIntField_uint32Overflow(t *testing.T) {
	row := Row{"x": uint32(math.MaxUint32)}
	if got := intField(row, "x"); got != 0 {
		t.Errorf("expected 0 for overflow uint32, got %d", got)
	}
}

func TestIntField_missingKey(t *testing.T) {
	row := Row{}
	if got := intField(row, "missing"); got != 0 {
		t.Errorf("expected 0 for missing key, got %d", got)
	}
}

func TestIntField_wrongType(t *testing.T) {
	row := Row{"x": "not an int"}
	if got := intField(row, "x"); got != 0 {
		t.Errorf("expected 0 for string value, got %d", got)
	}
}

// isKnownPageType

func TestIsKnownPageType_knownTypes(t *testing.T) {
	known := []byte{PageTypeDB, PageTypeData, PageTypeTDEF, PageTypeIIdx, PageTypeLIdx, PageTypeUsage}
	for _, pt := range known {
		if !isKnownPageType(pt) {
			t.Errorf("expected page type %#x to be known", pt)
		}
	}
}

func TestIsKnownPageType_unknownType(t *testing.T) {
	if isKnownPageType(0xFF) {
		t.Error("expected 0xFF to be unknown page type")
	}
}
