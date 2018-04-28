package xls

// Ranger is range type of multi rows
type Ranger interface {
	FirstRow() uint16
	LastRow() uint16
}

// CellRange is range type of multi cells in multi rows
type CellRange struct {
	FirstRowB uint16
	LastRowB  uint16
	FristColB uint16
	LastColB  uint16
}

// FirstRow return first row index
func (c *CellRange) FirstRow() uint16 {
	return c.FirstRowB
}

// LastRow return last row index
func (c *CellRange) LastRow() uint16 {
	return c.LastRowB
}

// FirstCol return first column index
func (c *CellRange) FirstCol() uint16 {
	return c.FristColB
}

// LastCol return last column index
func (c *CellRange) LastCol() uint16 {
	return c.LastColB
}
