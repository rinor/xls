package xls

type cell struct {
	contentHandler
}

func (c cell) IsValid() bool {
	return c.contentHandler != nil
}

func (c cell) IsHyperLink() bool {
	_, ok := c.contentHandler.(*HyperLink)
	return ok
}

func (c cell) MustHpyerLink() *HyperLink {
	return c.contentHandler.(*HyperLink)
}

func (c cell) IsNumber() bool {
	_, ok := c.contentHandler.(*NumberCol)
	return ok
}

func (c cell) MustNumber() *NumberCol {
	return c.contentHandler.(*NumberCol)
}

func (c cell) IsFormula() bool {
	_, ok := c.contentHandler.(*FormulaCol)
	return ok
}

func (c cell) MustFormula() *FormulaCol {
	return c.contentHandler.(*FormulaCol)
}
