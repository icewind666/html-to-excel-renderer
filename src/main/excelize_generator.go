package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

/**

 */
type ExcelizeGenerator struct {
	openedFile *excelize.File
	filename string
	currentSheet string
	currentCol int
	currentRow int
}

func (x *ExcelizeGenerator) Create() {
	x.openedFile = excelize.NewFile()
}


func (x *ExcelizeGenerator) Open(filename string) bool {
	f,err := excelize.OpenFile(filename)

	if err != nil {
		fmt.Println(err)
		return false
	}

	x.openedFile = f
	x.currentCol = 0 // current X in worksheet
	x.currentRow = 0 // current Y in worksheet
	return true
}

func (x *ExcelizeGenerator) Save(filename string) {
	x.openedFile.SaveAs(filename)
}

/**
Adds new sheet and makes it active
 */
func (x *ExcelizeGenerator) AddSheet(sheetName string) {
	index := x.openedFile.NewSheet(sheetName)
	x.openedFile.SetActiveSheet(index)
	x.currentSheet = sheetName
}

/**
Adds new sheet and makes it active
*/
func (x *ExcelizeGenerator) AddRow() {
	x.currentRow += 1
	err := x.openedFile.InsertRow(x.currentSheet, x.currentRow)
	if err != nil {
		fmt.Println(err)
	}

}

func (x *ExcelizeGenerator) GetCurrentRow() int {
	return x.currentRow
}

/**
Adds new sheet and makes it active
*/
func (x *ExcelizeGenerator) AddCell() {

}

/**
Set font size and font bold
 */
func (x *ExcelizeGenerator) SetCellFont(style *HtmlStyle)  {
	newStyle, err := x.openedFile.NewStyle(FontToExcelizeString(style))

	if err != nil {
		fmt.Println(err)
	}

	cell := x.GetCell()
	x.openedFile.SetCellStyle(x.currentSheet, cell, cell, newStyle)
}

func (x *ExcelizeGenerator) GetCell() string {
	return fmt.Sprintf("%d%d", x.currentCol, x.currentRow)
}

func (x *ExcelizeGenerator) SetRowHeight(rowHeight int) {
	x.openedFile.SetRowHeight(x.currentSheet, x.currentRow, float64(rowHeight))
}

/**
Sets cospan style for column.
Current column indexes are taken from currentRow & currentCol of  ExcelizeGenerator
instance
 */
func (x *ExcelizeGenerator) SetColspan() {

}

/**
Returns style for current cell
 */
func (x *ExcelizeGenerator) GetCellStyle() HtmlStyle {
	index,err := x.openedFile.GetCellStyle(x.currentSheet, x.GetCell())
	if err != nil {
		fmt.Println(err)
	}

	xlsxXf := x.openedFile.Styles.CellXfs.Xf[index]
	fontID := xlsxXf.FontID // font id in inner style table of workbook
	borderID := xlsxXf.BorderID
	font := x.openedFile.Styles.Fonts.Font[*fontID]
	isBold := font.B // bold
	fontSize := font.Sz.Val
	horizontalAlignment := xlsxXf.Alignment.Horizontal // horizontal cell alignment
	wrapWords := xlsxXf.Alignment.WrapText

	cellName,err := excelize.CoordinatesToCellName(x.currentCol, x.currentRow)
	if err != nil {
		fmt.Println(err)
	}

	width,err := x.openedFile.GetColWidth(x.currentSheet, cellName)
	if err != nil {
		fmt.Println(err)
	}

	rowHeight,err := x.openedFile.GetRowHeight(x.currentSheet, x.currentRow)
	if err != nil {
		fmt.Println(err)
	}

	isBordered := false
	border := x.openedFile.Styles.Borders.Border[*borderID]

	if border.Outline || border.DiagonalDown || border.DiagonalUp || border.Bottom.Style != "" || border.Top.Style != "" ||
		border.Left.Style != "" || border.Right.Style != "" {
		isBordered = true
	}

	//TODO: add merge cells
	//mergedCells := x.openedFile.GetMergeCells()


	htmlStyle := HtmlStyle{
		TextAlign:         horizontalAlignment,
		WordWrap:          wrapWords,
		Width:             width,
		Height:            rowHeight,
		BorderInheritance: false,
		BorderStyle:       isBordered,
		FontSize:          *fontSize,
		IsBold:            *isBold,
		Colspan:           0,
	}

	return htmlStyle
}

/**
Returns alignment json string
 */
func (x *ExcelizeGenerator) getCellAlignment(style *HtmlStyle) string{
	//fmt.Println(AlignmentToExcelizeString(style))
	return AlignmentToExcelizeString(style)
	//x.openedFile.SetCellStyle(x.currentSheet, cell, cell, newStyle)
}

/**
Returns aligment json string
*/
func (x *ExcelizeGenerator) getCellFont(style *HtmlStyle) string{
	//fmt.Println(FontToExcelizeString(style))
	return FontToExcelizeString(style)
	//x.openedFile.SetCellStyle(x.currentSheet, cell, cell, newStyle)
}

/**
Returns border json string
*/
func (x *ExcelizeGenerator) getCellBorders(style *HtmlStyle) string{
	//fmt.Println(BordersToExcelizeString(style))
	return BordersToExcelizeString(style)
}


func (x *ExcelizeGenerator) ApplyCellStyle(style *HtmlStyle) {
	styleJson := fmt.Sprintf(`
				{
					"font": %s,
				
					"border": %s,				
				
					"alignment": %s
				}`,
				x.getCellFont(style),
				x.getCellBorders(style),
				x.getCellAlignment(style),
				)


	newStyle, err := x.openedFile.NewStyle(styleJson)
	if err != nil {
		fmt.Println(err)
	}
	cell,err := excelize.CoordinatesToCellName(x.currentCol, x.currentRow)
	if err != nil {
		fmt.Println(err)
	}

	err = x.openedFile.SetCellStyle(x.currentSheet, cell, cell, newStyle)
	if err != nil {
		fmt.Println(err)
	}

}

func (x *ExcelizeGenerator) ApplyRowStyle(style *HtmlStyle) {
	if style.Height > 0 {
		x.SetRowHeight(int(style.Height))
	} else {
		x.SetRowHeight(15) // default
	}
	//TODO: set alignment and border on all cells
}

func (x *ExcelizeGenerator) ApplyColumnStyle(style *HtmlStyle) {
	styleJson := fmt.Sprintf(`{"alignment": %s}`,
		x.getCellAlignment(style),
	)

	newStyle, err := x.openedFile.NewStyle(styleJson)
	if err != nil {
		fmt.Println(err)
	}

	colName,err := excelize.ColumnNumberToName(x.currentCol)

	if style.Width > 0 {
		err := x.openedFile.SetColWidth(x.currentSheet, colName, colName, style.Width)
		if err != nil {
			fmt.Println(err)
		}

	}

	if err != nil {
		fmt.Println(err)
	}

	err = x.openedFile.SetColStyle(x.currentSheet, colName, newStyle)
	if err != nil {
		fmt.Println(err)
	}
}

func (x *ExcelizeGenerator) SetCellValue(value interface{}) {
	cellName,_ := excelize.CoordinatesToCellName(x.currentCol,x.currentRow)
	err := x.openedFile.SetCellValue(x.currentSheet, cellName, value)
	if err != nil {
		fmt.Println(err)
	}

}


/**
	"font": {
        "bold": true,
        "italic": false,
        "family": "Times New Roman",
        "size": 12,
        "color": "#777777"
    }
 */
func FontToExcelizeString(style *HtmlStyle) string  {
	isBold := "false"
	if style.IsBold {
		isBold = "true"
	}
	return fmt.Sprintf(`{"bold": %s,"size":%f}`, isBold, style.FontSize)
}

/**
	"alignment": {
        "horizontal": "center",
        "ident": 1,
        "justify_last_line": true,
        "reading_order": 0,
        "relative_indent": 1,
        "shrink_to_fit": true,
        "text_rotation": 45,
        "vertical": "",
        "wrap_text": true
    }
*/
func AlignmentToExcelizeString(style *HtmlStyle) string {
	isWrapped := "false"
	if style.WordWrap {
		isWrapped = "true"
	}
	return fmt.Sprintf(`{"horizontal": "%s","wrap_text": %s}`, style.TextAlign, isWrapped)
}


/**
"border": [
	{
		"type": "left",
		"color": "0000FF",
		"style": 3
	},
	{
		"type": "top",
		"color": "00FF00",
		"style": 4
	},
	{
		"type": "bottom",
		"color": "FFFF00",
		"style": 5
	},
	{
		"type": "right",
		"color": "FF0000",
		"style": 6
	},
	{
		"type": "diagonalDown",
		"color": "A020F0",
		"style": 7
	},
	{
		"type": "diagonalUp",
		"color": "A020F0",
		"style": 8
	}
],

 */
func BordersToExcelizeString(style *HtmlStyle) string {
	if style.BorderStyle {
		return `[
			{
				"type": "left",
				"style": 1
			},
			{
				"type": "top",
				"style": 1
			},
			{
				"type": "bottom",
				"style": 1
			},
			{
				"type": "right",
				"style": 1
			}]`
	}
	return "[]"
}