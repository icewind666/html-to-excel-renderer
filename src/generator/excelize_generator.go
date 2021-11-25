package generator

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/icewind666/html-to-excel-renderer/src/types"
	log "github.com/sirupsen/logrus"
)

// ExcelizeGenerator struct for handling state of excel generation processing
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
		log.WithError(err).Fatalln("Cant create excel file!")
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
		log.WithError(err).Fatalln("Cant save excel file!")
	}
}

// AddSheet Add new sheet with given name to workbook. New sheet is set as current
func (x *ExcelizeGenerator) AddSheet(sheetName string) {
	index := x.OpenedFile.NewSheet(sheetName)
	x.OpenedFile.SetActiveSheet(index)
	x.CurrentSheet = sheetName
}

// SetSheetName Renames sheet
func (x *ExcelizeGenerator) SetSheetName(oldSheetName string, sheetName string) {
	x.OpenedFile.SetSheetName(oldSheetName, sheetName)
	x.CurrentSheet = sheetName
}

// AddRow Move pointer to next row
func (x *ExcelizeGenerator) AddRow() {
	x.CurrentRow += 1
}

func (x *ExcelizeGenerator) GetSheetAtIndex(index int) string {
	return x.OpenedFile.GetSheetName(index)
}

func (x *ExcelizeGenerator) GetCurrentRow() int {
	return x.CurrentRow
}

// SetCellFont Set cell font from given style
func (x *ExcelizeGenerator) SetCellFont(style *types.HtmlStyle)  {
	newStyle, err := x.OpenedFile.NewStyle(FontToExcelizeString(style))

	if err != nil {
		log.WithError(err).Error("Cant set cell font")
	}

	cell := x.GetCell()
	err = x.OpenedFile.SetCellStyle(x.CurrentSheet, cell, cell, newStyle)

	if err != nil {
		log.WithError(err).Error("Cant set cell style")
	}
}

func (x *ExcelizeGenerator) GetCell() string {
	return fmt.Sprintf("%d%d", x.CurrentCol, x.CurrentRow)
}

func (x *ExcelizeGenerator) SetRowHeight(rowHeight int) {
	err := x.OpenedFile.SetRowHeight(x.CurrentSheet, x.CurrentRow, float64(rowHeight))

	if err != nil {
		log.WithError(err).Error("Cant set row height")
	}

}

func (x ExcelizeGenerator) GetCoords() (string,error) {
	return excelize.CoordinatesToCellName(x.CurrentCol, x.CurrentRow)
}

// SetColspan Sets cospan style for column. Current column indexes are taken from
// CurrentRow & CurrentCol of  ExcelizeGenerator instance
func (x *ExcelizeGenerator) SetColspan(endColumnNumber int) {
	endColumnIndex := x.CurrentCol + endColumnNumber-1
	endColumnName,err := excelize.CoordinatesToCellName(endColumnIndex, x.CurrentRow)

	if err != nil {
		log.WithError(err).Error("Cant get current cell coordinates")
	}

	currentCellName,err := excelize.CoordinatesToCellName(x.CurrentCol, x.CurrentRow)
	if err != nil {
		log.WithError(err).Error("Cant get current cell coordinates")
	}

	err = x.OpenedFile.MergeCell(x.CurrentSheet, currentCellName, endColumnName)
	if err != nil {
		log.WithError(err).Error("Cant merge cells")
	}
}


/**
Returns alignment json string
 */
func (x *ExcelizeGenerator) getCellAlignment(style *types.HtmlStyle) string{
	return AlignmentToExcelizeString(style)
}

/**
Returns alignment json string
*/
func (x *ExcelizeGenerator) getCellColor(style *types.HtmlStyle) string{
	return ColorToExcelizeString(style)
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

func (x *ExcelizeGenerator) ApplyBordersRange(style *types.HtmlStyle) {
	styleJson := fmt.Sprintf(`
				{
					"font": %s,
				
					"border": %s,				
				
					"alignment": %s,

					"fill": %s
				}`,
		x.getCellFont(style),
		x.getCellBorders(style),
		x.getCellAlignment(style),
		x.getCellColor(style),
	)

	newStyle, err := x.OpenedFile.NewStyle(styleJson)

	if err != nil {
		log.WithError(err).Fatalln("Cant create new style in Excel sheet")
	}

	cellFrom,err := excelize.CoordinatesToCellName(x.CurrentCol, x.CurrentRow)
	cellTo,err := excelize.CoordinatesToCellName(x.CurrentCol + style.Colspan -1, x.CurrentRow)

	log.Infoln("cell from ", x.CurrentCol)
	log.Infoln("cell to ", x.CurrentCol + style.Colspan)

	err = x.OpenedFile.SetCellStyle(x.CurrentSheet, cellFrom, cellTo, newStyle)

	if err != nil {
		log.WithError(err).Error("Cant set style")
	}

	if style.Colspan > 1 {
		x.SetColspan(style.Colspan)
	}


}

func (x *ExcelizeGenerator) ApplyCellStyle(style *types.HtmlStyle) {
	styleJson := fmt.Sprintf(`
				{
					"font": %s,
				
					"border": %s,				
				
					"alignment": %s,

					"fill": %s
				}`,
				x.getCellFont(style),
				x.getCellBorders(style),
				x.getCellAlignment(style),
				x.getCellColor(style),
				)

	newStyle, err := x.OpenedFile.NewStyle(styleJson)

	if err != nil {
		log.WithError(err).Fatalln("Cant create new style in Excel sheet")
	}

	cell,err := excelize.CoordinatesToCellName(x.CurrentCol, x.CurrentRow)

	if err != nil {
		log.WithError(err).Error("Cant get current cell coordinates")
	}

	err = x.OpenedFile.SetCellStyle(x.CurrentSheet, cell, cell, newStyle)

	if err != nil {
		log.WithError(err).Error("Cant set style")
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
		log.WithError(err).Error("Cant create new style")
	}

	colName,err := excelize.ColumnNumberToName(x.CurrentCol)

	if style.Width > 0 {
		err := x.OpenedFile.SetColWidth(x.CurrentSheet, colName, colName, style.Width)
		if err != nil {
			log.WithError(err).Error("Cant set column width")
		}

	}

	err = x.OpenedFile.SetColStyle(x.CurrentSheet, colName, newStyle)
	if err != nil {
		log.WithError(err).Error("Cant set column style")
	}
}

func (x *ExcelizeGenerator) SetCellValue(value string) {
	cellName,_ := excelize.CoordinatesToCellName(x.CurrentCol,x.CurrentRow)
	err := x.OpenedFile.SetCellStr(x.CurrentSheet, cellName, value)
	if err != nil {
		log.Fatalln(err.Error())
	}
}

func (x *ExcelizeGenerator) SetCellFloatValue(value float64) {
	cellName,_ := excelize.CoordinatesToCellName(x.CurrentCol,x.CurrentRow)
	err := x.OpenedFile.SetCellFloat(x.CurrentSheet, cellName, value, 3, 64)
	if err != nil {
		log.Fatalln(err.Error())
	}
}

func (x *ExcelizeGenerator) SetCellBoolValue(value bool) {
	cellName,_ := excelize.CoordinatesToCellName(x.CurrentCol,x.CurrentRow)
	err := x.OpenedFile.SetCellBool(x.CurrentSheet, cellName, value)
	if err != nil {
		log.WithError(err).Fatalln("Cant set bool value to cell")
	}
}

func (x *ExcelizeGenerator) SetCellIntValue(value int) {
	cellName,_ := excelize.CoordinatesToCellName(x.CurrentCol,x.CurrentRow)
	err := x.OpenedFile.SetCellInt(x.CurrentSheet, cellName, value)
	if err != nil {
		log.Fatalln(err.Error())
	}
}

func FontToExcelizeString(style *types.HtmlStyle) string  {
	// making font style json for excel
	isBold := "false"
	if style.IsBold {
		isBold = "true"
	}
	return fmt.Sprintf(`{"bold": %s,"size":%f}`, isBold, style.FontSize)
}

func AlignmentToExcelizeString(style *types.HtmlStyle) string {
	isWrapped := "false"
	if style.WordWrap {
		isWrapped = "true"
	}
	return fmt.Sprintf(`{"horizontal": "%s","wrap_text": %s, "vertical": "%s"}`, style.TextAlign, isWrapped, style.VerticalAlign)
}

func ColorToExcelizeString(style *types.HtmlStyle) string {
	if style.BackgroundColor != "" {
		return fmt.Sprintf(`{"type": "pattern","color":["%s"],"pattern":1}`, style.BackgroundColor)
	}
	return fmt.Sprintf(`{"type": "pattern","color":["%s"],"pattern":1}`, "#ffffff")
}

func BordersToExcelizeString(style *types.HtmlStyle) string {
	if style.BorderStyle {
		return `[
			{
				"type": "left",
				"style": 1,
				"color":"#000000"
			},
			{
				"type": "top",
				"style": 1,
				"color":"#000000"
			},
			{
				"type": "bottom",
				"style": 1,
				"color":"#000000"
			},
			{
				"type": "right",
				"style": 1,
				"color":"#000000"
			}]`
	}
	return "[]"
}