package main

import (
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"testing"
)

func TestStyles(t *testing.T) {
	ex := ExcelizeGenerator{
		filename:     "",
		currentSheet: "",
		currentRow: 1,
		currentCol: "A",
	}
	//ex.Open("test.xlsx")
	ex.Create()
	ex.AddSheet("first")
	ex.openedFile.SetCellValue("first", "A1", "this is a test")


	//style_str := fmt.Sprintf(`{"alignment":{"horizontal": "center","wrap_text": false}}`)
	//newStyle, err := ex.openedFile.NewStyle(al)

	titleStyle, err := ex.openedFile.NewStyle(&excelize.Style{
		Font:      &excelize.Font{Color: "1f7f3b", Bold: true, Family: "Arial"},
		Fill:      excelize.Fill{Type: "pattern", Color: []string{"E6F4EA"}, Pattern: 1},
		Alignment: &excelize.Alignment{Vertical: "center", Horizontal: "right"},
		Border:    []excelize.Border{{Type: "top", Style: 2, Color: "1f7f3b"}},
	})
	if err != nil {
		fmt.Println(err)
	}

	cell := "A1"
	err = ex.openedFile.SetCellStyle(ex.currentSheet, cell, cell, titleStyle)
	if err != nil {
		fmt.Println(err)
	}
	index, err := ex.openedFile.GetCellStyle("first", "A1")
	err = ex.openedFile.SetCellStyle(ex.currentSheet, cell, cell, titleStyle)
	fillID := *ex.openedFile.Styles.CellXfs.Xf[index].Alignment
	fillID2 := *ex.openedFile.Styles.CellXfs.Xf[index].FillID
	fmt.Println("style id = ", fillID.Horizontal)
	fmt.Println("style id = ", fillID2)
	//ex.SetCellAlignment(s)

	//ex.SetRowHeight(40)
	ex.Save("test.xlsx")

	//t.Error("rrrrrr")
}
