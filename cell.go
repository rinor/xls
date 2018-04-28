package xls

type cell struct {
	contentHandler
}

// IsValid return true if the cell is valid
func (c cell) IsValid() bool {
	return c.contentHandler != nil
}

// IsHyperLink return true if the cell is a hyperlink
func (c cell) IsHyperLink() bool {
	_, ok := c.contentHandler.(*HyperLink)
	return ok
}

// MustHpyerLink always return hyperlink
func (c cell) MustHpyerLink() *HyperLink {
	return c.contentHandler.(*HyperLink)
}

// IsNumber return true if the cell is the number
func (c cell) IsNumber() bool {
	_, ok := c.contentHandler.(*NumberCol)
	return ok
}

// MustNumber always return number
func (c cell) MustNumber() *NumberCol {
	return c.contentHandler.(*NumberCol)
}

// IsFormula return true if the cell is formula
func (c cell) IsFormula() bool {
	_, ok := c.contentHandler.(*FormulaCol)
	return ok
}

// MustFormula always return formula
func (c cell) MustFormula() *FormulaCol {
	return c.contentHandler.(*FormulaCol)
}

// IsBlank return true if the cell is blank
func (c cell) IsBlank() bool {
	_, ok := c.contentHandler.(*BlankCol)
	return ok
}

// MustBlank always return blank
func (c cell) MustBlank() *BlankCol {
	return c.contentHandler.(*BlankCol)
}

// IsRk return true if the cell is rk
func (c cell) IsRk() bool {
	_, ok := c.contentHandler.(*RkCol)
	return ok
}

// MustRk alwasy return RkCol
func (c cell) MustRk() *RkCol {
	return c.contentHandler.(*RkCol)
}
