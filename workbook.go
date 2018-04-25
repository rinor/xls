package xls

import (
	"bytes"
	"encoding/binary"
	"io"
	"strings"
	"unicode/utf16"
)

// WorkBook represents xls workbook type
type WorkBook struct {
	Is5ver        bool
	Type          uint16
	Codepage      uint16
	Xfs           []XF
	Fonts         []Font
	Formats       map[uint16]*Format
	sheets        []*WorkSheet
	Author        string
	rs            io.ReadSeeker
	sst           []string
	ref           *extSheetRef
	continueUtf16 uint16
	continueRich  uint16
	continueApsb  uint32
	dateMode      uint16
}

// newWorkBookFromOle2 read workbook from ole2 file
func newWorkBookFromOle2(rs io.ReadSeeker) *WorkBook {
	var wb = &WorkBook{
		rs:      rs,
		ref:     new(extSheetRef),
		sheets:  make([]*WorkSheet, 0),
		Formats: make(map[uint16]*Format),
	}

	wb.parse(rs)
	wb.prepare()

	return wb
}

func (w *WorkBook) parse(buf io.ReadSeeker) {
	b := new(bof)
	bp := new(bof)
	offset := 0

	for {
		if err := binary.Read(buf, binary.LittleEndian, b); err == nil {
			bp, b, offset = w.parseBof(buf, b, bp, offset)
		} else {
			break
		}
	}
}

func (wb *WorkBook) parseBof(buf io.ReadSeeker, b *bof, pre *bof, offsetPre int) (after *bof, afterUsing *bof, offset int) {
	after = b
	afterUsing = pre
	var bts = make([]byte, b.Size)
	binary.Read(buf, binary.LittleEndian, bts)
	item := bytes.NewReader(bts)
	switch b.Id {
	case 0x0809: // BOF
		bif := new(biffHeader)
		binary.Read(item, binary.LittleEndian, bif)
		if bif.Ver != 0x600 {
			wb.Is5ver = true
		}
		wb.Type = bif.Type
	case 0x0042: // CODEPAGE
		binary.Read(item, binary.LittleEndian, &wb.Codepage)
	case 0x3C: // CONTINUE
		if pre.Id == 0xfc {
			var size uint16
			var err error
			if wb.continueUtf16 >= 1 {
				size = wb.continueUtf16
				wb.continueUtf16 = 0
			} else {
				err = binary.Read(item, binary.LittleEndian, &size)
			}
			for err == nil && offsetPre < len(wb.sst) {
				var str string
				if size > 0 {
					str, err = wb.parseString(item, size)
					wb.sst[offsetPre] = wb.sst[offsetPre] + str
				}

				if err == io.EOF {
					break
				}

				offsetPre++
				err = binary.Read(item, binary.LittleEndian, &size)
			}
		}
		offset = offsetPre
		after = pre
		afterUsing = b
	case 0x00FC: // SST
		info := new(SstInfo)
		binary.Read(item, binary.LittleEndian, info)
		wb.sst = make([]string, info.Count)
		var size uint16
		var i = 0
		for ; i < int(info.Count); i++ {
			var err error
			if err = binary.Read(item, binary.LittleEndian, &size); err == nil {
				var str string
				str, err = wb.parseString(item, size)
				wb.sst[i] = wb.sst[i] + str
			}

			if err == io.EOF {
				break
			}
		}
		offset = i
	case 0x0085: // SHEET
		var bs = new(boundsheet)
		binary.Read(item, binary.LittleEndian, bs)
		// different for BIFF5 and BIFF8
		wb.addSheet(bs, item)
	case 0x0017: // EXTERNSHEET
		if !wb.Is5ver {
			binary.Read(item, binary.LittleEndian, &wb.ref.Num)
			wb.ref.Info = make([]ExtSheetInfo, wb.ref.Num)
			binary.Read(item, binary.LittleEndian, &wb.ref.Info)
		}
	case 0x00e0: // XF
		if wb.Is5ver {
			xf := new(Xf5)
			binary.Read(item, binary.LittleEndian, xf)
			wb.addXf(xf)
		} else {
			xf := new(Xf8)
			binary.Read(item, binary.LittleEndian, xf)
			wb.addXf(xf)
		}
	case 0x0031: // FONT
		f := new(FontInfo)
		binary.Read(item, binary.LittleEndian, f)
		wb.addFont(f, item)
	case 0x041E: //FORMAT
		format := new(Format)
		binary.Read(item, binary.LittleEndian, &format.Head)
		if raw, err := wb.parseString(item, format.Head.Size); nil == err && "" != raw {
			format.Raw = strings.Split(raw, ";")
		} else {
			format.Raw = []string{}
		}

		wb.addFormat(format)
	case 0x0022: //DATEMODE
		binary.Read(item, binary.LittleEndian, &wb.dateMode)
	}
	return
}

func (w *WorkBook) addXf(xf XF) {
	w.Xfs = append(w.Xfs, xf)
}

func (w *WorkBook) addFont(font *FontInfo, buf io.ReadSeeker) {
	name, _ := w.parseString(buf, uint16(font.NameB))
	w.Fonts = append(w.Fonts, Font{Info: font, Name: name})
}

func (w *WorkBook) addFormat(format *Format) {
	w.Formats[format.Head.Index] = format
}

func (w *WorkBook) addSheet(sheet *boundsheet, buf io.ReadSeeker) {
	name, _ := w.parseString(buf, uint16(sheet.Name))
	w.sheets = append(w.sheets, &WorkSheet{id: len(w.sheets), bs: sheet, Name: name, wb: w})
}

// prepare process workbook struct
func (w *WorkBook) prepare() {
	for k, v := range builtInNumFmt {
		if _, ok := w.Formats[k]; !ok {
			w.Formats[k] = &Format{
				Raw: strings.Split(v, ";"),
			}
		}
	}
	for _, v := range w.Formats {
		v.Prepare()
	}
}

//prepareSheet reads a sheet from the compress file to memory, you should call this before you try to get anything from sheet
func (w *WorkBook) prepareSheet(sheet *WorkSheet) {
	w.rs.Seek(int64(sheet.bs.Filepos), 0)
	sheet.parse(w.rs)
}

func (w *WorkBook) parseString(buf io.ReadSeeker, size uint16) (res string, err error) {
	if w.Is5ver {
		var bts = make([]byte, size)
		_, err = buf.Read(bts)
		res = string(bytes.Trim(bts, "\r\n\t "))
	} else {
		var richtextNum = uint16(0)
		var phoneticSize = uint32(0)
		var flag byte
		err = binary.Read(buf, binary.LittleEndian, &flag)
		if flag&0x8 != 0 {
			err = binary.Read(buf, binary.LittleEndian, &richtextNum)
		} else if w.continueRich > 0 {
			richtextNum = w.continueRich
			w.continueRich = 0
		}
		if flag&0x4 != 0 {
			err = binary.Read(buf, binary.LittleEndian, &phoneticSize)
		} else if w.continueApsb > 0 {
			phoneticSize = w.continueApsb
			w.continueApsb = 0
		}
		if flag&0x1 != 0 {
			var bts = make([]uint16, size)
			var i = uint16(0)
			for ; i < size && err == nil; i++ {
				err = binary.Read(buf, binary.LittleEndian, &bts[i])
			}
			runes := utf16.Decode(bts[:i])
			res = strings.Trim(string(runes), "\r\n\t ")
			if i < size {
				w.continueUtf16 = size - i + 1
			}
		} else {
			var bts = make([]byte, size)
			var n int
			n, err = buf.Read(bts)
			if uint16(n) < size {
				w.continueUtf16 = size - uint16(n)
				err = io.EOF
			}

			var bts1 = make([]uint16, n)
			for k, v := range bts[:n] {
				bts1[k] = uint16(v)
			}
			runes := utf16.Decode(bts1)
			res = strings.Trim(string(runes), "\r\n\t ")
		}
		if richtextNum > 0 {
			var bts []byte
			var ss int64
			if w.Is5ver {
				ss = int64(2 * richtextNum)
			} else {
				ss = int64(4 * richtextNum)
			}
			bts = make([]byte, ss)
			err = binary.Read(buf, binary.LittleEndian, bts)
			if err == io.EOF {
				w.continueRich = richtextNum
			}
		}
		if phoneticSize > 0 {
			var bts []byte
			bts = make([]byte, phoneticSize)
			err = binary.Read(buf, binary.LittleEndian, bts)
			if err == io.EOF {
				w.continueApsb = phoneticSize
			}
		}
	}
	return
}

// Format formats value to string
func (w *WorkBook) Format(xf uint16, v float64) (string, bool) {
	var val string
	var idx = int(xf)
	if len(w.Xfs) > idx {
		if formatter := w.Formats[w.Xfs[idx].FormatNo()]; nil != formatter {
			return formatter.String(v), true
		}
	}

	return val, false
}

// GetSheet gets one sheet by its number
func (w *WorkBook) GetSheet(num int) *WorkSheet {
	if num < len(w.sheets) {
		s := w.sheets[num]
		if !s.parsed {
			w.prepareSheet(s)
		}
		return s
	}
	return nil
}

// NumSheets get the number of all sheets, look into example
func (w *WorkBook) NumSheets() int {
	return len(w.sheets)
}

// ReadAllCells helper function to read all cells from file
// Notice: the max value is the limit of the max capacity of lines.
// Warning: the helper function will need big memory if file is large.
func (w *WorkBook) ReadAllCells(max int) (res [][]string) {
	res = make([][]string, 0)
	for _, sheet := range w.sheets {
		if len(res) < max {
			max = max - len(res)
			w.prepareSheet(sheet)
			if sheet.MaxRow != 0 {
				length := int(sheet.MaxRow) + 1
				if max < length {
					length = max
				}
				temp := make([][]string, length)
				for k, row := range sheet.rows {
					data := make([]string, 0)
					if len(row.cols) > 0 {
						for _, col := range row.cols {
							if uint16(len(data)) <= col.LastCol() {
								data = append(data, make([]string, col.LastCol()-uint16(len(data))+1)...)
							}
							str := col.String(w)

							for i := uint16(0); i < col.LastCol()-col.FirstCol()+1; i++ {
								data[col.FirstCol()+i] = str[i]
							}
						}
						if length > int(k) {
							temp[k] = data
						}
					}
				}
				res = append(res, temp...)
			}
		}
	}
	return
}
