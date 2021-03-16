package generator

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/icewind666/html-to-excel-render/src/types"
)

/**

 */
type ExcelizeGenerator struct {
	OpenedFile   *excelize.File
	Filename     string
	CurrentSheet string
	CurrentCol   int
	CurrentRow   int
}

func (x *ExcelizeGenerator) Create() {
	x.OpenedFile = excelize.NewFile()
}


func (x *ExcelizeGenerator) Open(filename string) bool {
	f,err := excelize.OpenFile(filename)

	if err != nil {
		fmt.Println(err)
		return false
	}

	x.OpenedFile = f
	x.CurrentCol = 0 // current X in worksheet
	x.CurrentRow = 0 // current Y in worksheet
	return true
}

func (x *ExcelizeGenerator) Save(filename string) {
	err := x.OpenedFile.SaveAs(filename)

	if err != nil {
		fmt.Println(err.Error())
	}
}

/**

Adds new sheet and makes it active
 */
func (x *ExcelizeGenerator) AddSheet(sheetName string) {
	index := x.OpenedFile.NewSheet(sheetName)
	x.OpenedFile.SetActiveSheet(index)
	x.CurrentSheet = sheetName
}

/**

Adds new sheet and makes it active
*/
func (x *ExcelizeGenerator) SetSheetName(oldSheetName string, sheetName string) {
	x.OpenedFile.SetSheetName(oldSheetName, sheetName)
	x.CurrentSheet = sheetName
}

/**
Adds new sheet and makes it active
*/
func (x *ExcelizeGenerator) AddRow() {
	x.CurrentRow += 1
	err := x.OpenedFile.InsertRow(x.CurrentSheet, x.CurrentRow)

	if err != nil {
		fmt.Println(err)
	}

}

func (x *ExcelizeGenerator) GetSheetAtIndex(index int) string {
	return x.OpenedFile.GetSheetName(index)
}

func (x *ExcelizeGenerator) GetCurrentRow() int {
	return x.CurrentRow
}

/**
Adds new sheet and makes it active
*/
func (x *ExcelizeGenerator) AddCell() {

}

/**
Set font size and font bold
 */
func (x *ExcelizeGenerator) SetCellFont(style *types.HtmlStyle)  {
	newStyle, err := x.OpenedFile.NewStyle(FontToExcelizeString(style))

	if err != nil {
		fmt.Println(err)
	}

	cell := x.GetCell()
	err = x.OpenedFile.SetCellStyle(x.CurrentSheet, cell, cell, newStyle)

	if err != nil {
		fmt.Println(err)
	}
}

func (x *ExcelizeGenerator) GetCell() string {
	return fmt.Sprintf("%d%d", x.CurrentCol, x.CurrentRow)
}

func (x *ExcelizeGenerator) SetRowHeight(rowHeight int) {
	err := x.OpenedFile.SetRowHeight(x.CurrentSheet, x.CurrentRow, float64(rowHeight))

	if err != nil {
		fmt.Println(err)
	}

}

func (x ExcelizeGenerator) GetCoords() (string,error) {
	return excelize.CoordinatesToCellName(x.CurrentCol, x.CurrentRow)
}

/**
Sets cospan style for column.
Current column indexes are taken from CurrentRow & CurrentCol of  ExcelizeGenerator
instance
 */
func (x *ExcelizeGenerator) SetColspan(endColumnNumber int) {
	endColumnIndex := x.CurrentCol + endColumnNumber-1
	endColumnName,err := excelize.CoordinatesToCellName(endColumnIndex, x.CurrentRow)

	if err != nil {
		fmt.Println(err)
	}

	currentCellName,err := excelize.CoordinatesToCellName(x.CurrentCol, x.CurrentRow)
	if err != nil {
		fmt.Println(err)
	}

	err = x.OpenedFile.MergeCell(x.CurrentSheet, currentCellName, endColumnName)
	if err != nil {
		fmt.Println(err)
	}
}

/**
Returns style for current cell
 */
func (x *ExcelizeGenerator) GetCellStyle() types.HtmlStyle {
	index,err := x.OpenedFile.GetCellStyle(x.CurrentSheet, x.GetCell())

	if err != nil {
		fmt.Println(err)
	}

	xlsxXf := x.OpenedFile.Styles.CellXfs.Xf[index]
	fontID := xlsxXf.FontID // font id in inner style table of workbook
	borderID := xlsxXf.BorderID
	font := x.OpenedFile.Styles.Fonts.Font[*fontID]
	isBold := font.B // bold
	fontSize := font.Sz.Val
	horizontalAlignment := xlsxXf.Alignment.Horizontal // horizontal cell alignment
	wrapWords := xlsxXf.Alignment.WrapText

	cellName,err := excelize.CoordinatesToCellName(x.CurrentCol, x.CurrentRow)

	if err != nil {
		fmt.Println(err)
	}

	width,err := x.OpenedFile.GetColWidth(x.CurrentSheet, cellName)

	if err != nil {
		fmt.Println(err)
	}

	rowHeight,err := x.OpenedFile.GetRowHeight(x.CurrentSheet, x.CurrentRow)

	if err != nil {
		fmt.Println(err)
	}

	isBordered := false
	border := x.OpenedFile.Styles.Borders.Border[*borderID]

	if border.Outline || border.DiagonalDown || border.DiagonalUp || border.Bottom.Style != "" || border.Top.Style != "" ||
		border.Left.Style != "" || border.Right.Style != "" {
		isBordered = true
	}

	//TODO: parse colspan from merged cells
	//mergedCells := x.OpenedFile.GetMergeCells()


	htmlStyle := types.HtmlStyle{
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
func (x *ExcelizeGenerator) getCellAlignment(style *types.HtmlStyle) string{
	return AlignmentToExcelizeString(style)
}

/**
Returns aligment json string
*/
func (x *ExcelizeGenerator) getCellFont(style *types.HtmlStyle) string{
	return FontToExcelizeString(style)
}

/**
Returns border json string
*/
func (x *ExcelizeGenerator) getCellBorders(style *types.HtmlStyle) string{
	return BordersToExcelizeString(style)
}


func (x *ExcelizeGenerator) ApplyCellStyle(style *types.HtmlStyle) {
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


	newStyle, err := x.OpenedFile.NewStyle(styleJson)

	if err != nil {
		fmt.Println(err)
	}

	cell,err := excelize.CoordinatesToCellName(x.CurrentCol, x.CurrentRow)

	if err != nil {
		fmt.Println(err)
	}

	err = x.OpenedFile.SetCellStyle(x.CurrentSheet, cell, cell, newStyle)

	if err != nil {
		fmt.Println(err)
	}

	if style.Colspan > 1 {
		x.SetColspan(style.Colspan)
	}

}

func (x *ExcelizeGenerator) ApplyRowStyle(style *types.HtmlStyle) {
	if style.Height > 0 {
		x.SetRowHeight(int(style.Height))
	} else {
		x.SetRowHeight(15) // default
	}
}

func (x *ExcelizeGenerator) ApplyColumnStyle(style *types.HtmlStyle) {
	styleJson := fmt.Sprintf(`{"alignment": %s}`,
		x.getCellAlignment(style),
	)

	newStyle, err := x.OpenedFile.NewStyle(styleJson)
	if err != nil {
		fmt.Println(err)
	}

	colName,err := excelize.ColumnNumberToName(x.CurrentCol)

	if style.Width > 0 {
		err := x.OpenedFile.SetColWidth(x.CurrentSheet, colName, colName, style.Width)
		if err != nil {
			fmt.Println(err)
		}

	}

	err = x.OpenedFile.SetColStyle(x.CurrentSheet, colName, newStyle)
	if err != nil {
		fmt.Println(err)
	}
}

func (x *ExcelizeGenerator) SetCellValue(value interface{}) {
	cellName,_ := excelize.CoordinatesToCellName(x.CurrentCol,x.CurrentRow)
	err := x.OpenedFile.SetCellValue(x.CurrentSheet, cellName, value)
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
func FontToExcelizeString(style *types.HtmlStyle) string  {
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
func AlignmentToExcelizeString(style *types.HtmlStyle) string {
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
func BordersToExcelizeString(style *types.HtmlStyle) string {
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