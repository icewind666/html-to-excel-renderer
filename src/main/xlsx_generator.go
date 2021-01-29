package main

type XlsxGenerator interface {
	Open(filename string) bool
	Close()
	AddSheet(sheetName string)
	SetRowHeight(rowHeight int)
	Save(filename string)
	Create()
	ApplyCellStyle(style *HtmlStyle)
	AddRow()
}
