// Package mdb implements a reader for the Microsoft Jet4 MDB database format.
package mdb

import (
	"encoding/binary"
	"fmt"
	"io"
	"os"
)

const (
	PageSize = 4096

	// Jet version values at offset 0x14 (little-endian uint32).
	JetVersion3    = 0x00 // Access 97
	JetVersion4    = 0x01 // Access 2000
	JetVersion4x   = 0x02 // Access 2002/2003
	JetVersionACE  = 0x03 // Access 2007
	JetVersionACE4 = 0x04 // Access 2010
	JetVersionACE5 = 0x05 // Access 2013
	JetVersionACE6 = 0x06 // Access 2016+

	// Page types (byte 0 of each page).
	PageTypeDB    = 0x00
	PageTypeData  = 0x01
	PageTypeTDEF  = 0x02
	PageTypeIIdx  = 0x03
	PageTypeLIdx  = 0x04
	PageTypeUsage = 0x05

	// Header offsets.
	offsetMagic      = 0x00
	offsetDBName     = 0x04
	offsetJetVersion = 0x14
	offsetCodePage   = 0x3C
	offsetDBKey      = 0x3E
	offsetSortOrder  = 0x6E
)

var magicBytes = [4]byte{0x00, 0x01, 0x00, 0x00}

// Header holds parsed database header fields from page 0.
type Header struct {
	JetVersion uint32
	DBName     string
	CodePage   uint16
	DBKey      uint32
	SortOrder  uint32
}

// Database represents an open MDB file.
type Database struct {
	f         *os.File
	Header    Header
	pageCount int64
}

// Open opens an MDB file and parses its header.
func Open(path string) (*Database, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, fmt.Errorf("mdb: open: %w", err)
	}

	fi, err := f.Stat()
	if err != nil {
		f.Close()
		return nil, fmt.Errorf("mdb: stat: %w", err)
	}

	if fi.Size() < PageSize {
		f.Close()
		return nil, fmt.Errorf("mdb: file too small (%d bytes)", fi.Size())
	}

	db := &Database{
		f:         f,
		pageCount: fi.Size() / PageSize,
	}

	if err := db.parseHeader(); err != nil {
		f.Close()
		return nil, err
	}

	return db, nil
}

// Close closes the database file.
func (db *Database) Close() error {
	return db.f.Close()
}

// PageCount returns the total number of pages in the database.
func (db *Database) PageCount() int64 {
	return db.pageCount
}

func (db *Database) parseHeader() error {
	page := make([]byte, PageSize)
	if _, err := db.f.ReadAt(page, 0); err != nil {
		return fmt.Errorf("mdb: read header: %w", err)
	}

	// Validate magic bytes.
	var magic [4]byte
	copy(magic[:], page[offsetMagic:offsetMagic+4])

	if magic != magicBytes {
		return fmt.Errorf("mdb: invalid magic bytes: %x", magic)
	}

	// DB name/identifier string (null-terminated, up to 16 bytes).
	nameEnd := offsetDBName
	for nameEnd < offsetDBName+16 && page[nameEnd] != 0 {
		nameEnd++
	}

	db.Header.DBName = string(page[offsetDBName:nameEnd])

	db.Header.JetVersion = binary.LittleEndian.Uint32(page[offsetJetVersion:])
	db.Header.CodePage = binary.LittleEndian.Uint16(page[offsetCodePage:])
	db.Header.DBKey = binary.LittleEndian.Uint32(page[offsetDBKey:])
	db.Header.SortOrder = binary.LittleEndian.Uint32(page[offsetSortOrder:])

	return nil
}

// IsJet4 returns true if the database uses Jet4 or later format.
func (db *Database) IsJet4() bool {
	return db.Header.JetVersion >= JetVersion4
}

// ReadPage reads a single page by page number.
func (db *Database) ReadPage(pageNum int64) ([]byte, error) {
	if pageNum < 0 || pageNum >= db.pageCount {
		return nil, fmt.Errorf("mdb: page %d out of range (0..%d)", pageNum, db.pageCount-1)
	}

	page := make([]byte, PageSize)

	_, err := db.f.ReadAt(page, pageNum*PageSize)
	if err != nil && err != io.EOF {
		return nil, fmt.Errorf("mdb: read page %d: %w", pageNum, err)
	}

	return page, nil
}

// PageType returns the type byte of a page.
func PageType(page []byte) byte {
	return page[0]
}
