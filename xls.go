package xls

import (
	"bytes"
	"errors"
	"io"
	"io/ioutil"
	"os"

	"github.com/extrame/ole2"
)

// Open open one xls file with some charset
func Open(file string, charset string) (*WorkBook, error) {
	fi, err := os.Open(file)
	if err != nil {
		return nil, err
	}
	return OpenReader(fi, charset)
}

// OpenWithBuffer open one xls file with memory buffer
func OpenWithBuffer(file string, charset string) (*WorkBook, error) {
	fi, err := ioutil.ReadFile(file)
	if err != nil {
		return nil, err
	}
	return OpenReader(bytes.NewReader(fi), charset)
}

// OpenWithCloser open one xls file and return the closer
func OpenWithCloser(file string, charset string) (*WorkBook, io.Closer, error) {
	fi, err := os.Open(file)
	if err != nil {
		return nil, nil, err
	}

	wb, err := OpenReader(fi, charset)
	if err != nil {
		return nil, nil, err
	}

	return wb, fi, nil
}

// OpenReader open xls file from reader
func OpenReader(reader io.ReadSeeker, charset string) (wb *WorkBook, err error) {
	ole, err := ole2.Open(reader, charset)
	if err != nil {
		return nil, err
	}

	dir, err := ole.ListDir()
	if err != nil {
		return nil, err
	}

	var book *ole2.File
	var root *ole2.File
	for _, file := range dir {
		name := file.Name()
		if name == "Workbook" {
			book = file
			// break
		}
		if name == "Book" {
			book = file
			// break
		}
		if name == "Root Entry" {
			root = file
		}
	}

	if book != nil {
		return newWorkBookFromOle2(ole.OpenFile(book, root)), nil
	}

	return nil, errors.New("book not found")
}
