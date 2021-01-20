package main

import (
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/tealeg/xlsx/v3"
)

type TealegGenerator struct {
	openedFile *excelize.File
	filename string
	currentSheet string
	currentCol int
	currentRow int
}


// Merge cells horizontally to apply colspan
func setColspan(cell *xlsx.Cell, style *HtmlStyle) {
	if style.Colspan > 1 {
		cell.Merge(style.Colspan-1, 0)
	}
}

// Setup font options for cell
func setFont(currentCellStyle *xlsx.Style, style *HtmlStyle) {
	if style.FontSize > 0 {
		currentCellStyle.Font.Size = float64(style.FontSize)
	}
	currentCellStyle.Font.Bold = style.IsBold
	currentCellStyle.ApplyFont = true
}


// Setup cell alignment options from style
func setAlignment(cellStyle *xlsx.Style, style *HtmlStyle) {
	if style.WordWrap {
		cellStyle.Alignment.WrapText = style.WordWrap
	}

	cellStyle.Alignment.Vertical = TextVerticalAlignStyleAttrValue
	cellStyle.ApplyAlignment = true

	if style.TextAlign != "" {
		cellStyle.Alignment.Horizontal = style.TextAlign
	}
}

// Setup cell border options from html style
// Border is always set to "thin" value
func setBorder(cellStyle *xlsx.Style, style *HtmlStyle) {
	if style.BorderStyle {
		cellStyle.Border.Top = ExcelBorderTypeValue
		cellStyle.Border.Bottom = ExcelBorderTypeValue
		cellStyle.Border.Left = ExcelBorderTypeValue
		cellStyle.Border.Right = ExcelBorderTypeValue
		cellStyle.ApplyBorder = true
	}
}

// Apply style from html to excel column
//func ApplyColumnStyle(cell *xlsx.Cell, sheet *xlsx.Sheet, colIndex int, style *HtmlStyle) {
//	if cell == nil {
//		return
//	}
//	currentCellStyle := cell.GetStyle()
//	setColspan(cell, style)
//	column := sheet.Col(colIndex-1)
//
//	if column == nil {
//		c := xlsx.NewColForRange(colIndex,colIndex)
//		sheet.SetColParameters(c)
//		columnAgain := sheet.Col(colIndex-1)
//
//		if columnAgain == nil {
//			panic("Error setting column styles")
//		}
//		column = columnAgain
//		column.SetStyle(xlsx.NewStyle())
//	}
//
//	column.GetStyle().Alignment.WrapText = style.WordWrap
//	column.GetStyle().ApplyAlignment = true
//
//	if style.Width > 0 {
//		sheet.SetColWidth(colIndex, colIndex, style.Width)
//	}
//
//	setFont(currentCellStyle, style)
//	setBorder(currentCellStyle, style)
//}


// Apply style from html Td to excel Cell
//func ApplyCellStyle(cell *xlsx.Cell, style *HtmlStyle) {
//	if cell == nil {
//		return
//	}
//	currentCellStyle := cell.GetStyle()
//	setFont(currentCellStyle, style)
//	setAlignment(currentCellStyle, style)
//	setBorder(currentCellStyle, style)
//	setColspan(cell, style)
//}

// Apply html style to Excel row.
// Alignment and Border are applied to each cell of the row
func ApplyRowStyle(row *xlsx.Row, style *HtmlStyle) {
	if row == nil {
		return
	}

	if style.Height > 0 {
		row.SetHeight(style.Height)
	} else {
		row.SetHeight(15) // default
	}


	row.ForEachCell(func(c *xlsx.Cell) error {
		setAlignment(c.GetStyle(), style)
		setBorder(c.GetStyle(), style)
		return nil
	})
}