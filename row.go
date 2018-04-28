package xls

type rowInfo struct {
	Index    uint16
	First    uint16
	Last     uint16
	Height   uint16
	Notused  uint16
	Notused2 uint16
	Flags    uint32
}

// Row the data of one row
type Row struct {
	wb   *WorkBook
	info *rowInfo
	cols map[uint16]contentHandler
}

// Col Get the Nth Column on the Row.
func (r *Row) Col(i int) string {
	var val string
	var serial = uint16(i)

	if ch, ok := r.cols[serial]; ok {
		val = ch.String(r.wb)[0]
	} else {
		for _, v := range r.cols {
			if v.FirstCol() <= serial && v.LastCol() >= serial {
				val = v.String(r.wb)[serial-v.FirstCol()]

				break
			}
		}
	}

	return val
}

// Column will return the index i cell of the row
func (r *Row) Column(i int) cell {
	var serial = uint16(i)

	if ch, ok := r.cols[serial]; ok {
		return cell{ch}
	}

	for _, v := range r.cols {
		if v.FirstCol() <= serial && v.LastCol() >= serial {
			return cell{v}
		}
	}

	return cell{nil}
}

// FirstCol Get the number of First Col of the Row.
func (r *Row) FirstCol() int {
	return int(r.info.First)
}

//LastCol Get the number of Last Col of the Row.
func (r *Row) LastCol() int {
	return int(r.info.Last)
}
