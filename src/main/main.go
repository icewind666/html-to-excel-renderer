package main

import (
	"encoding/json"
	"fmt"
	"github.com/aymerick/raymond"
	"github.com/jbowtie/gokogiri"
	"github.com/jbowtie/gokogiri/xml"
	"github.com/jbowtie/gokogiri/xpath"
	"github.com/tealeg/xlsx/v3"
	"gopkg.in/natefinch/lumberjack.v2"
	"io/ioutil"
	"log"
	"os"
	"runtime"
	"strconv"
	"strings"
	"time"
)

// App version
const Version = 5

// Json root attribute name to get data from
const JsonDataRootAttributeName = "data"

// Multipliers for converting html values to excel
const PixelsToExcelWidthCoeff = 0.15
const PixelsToExcelHeightCoeff = 0.10

// Html style constants
const StyleAttrName = "style"
const ColspanAttrName = "colspan"
const TextAlignStyleAttr = "text-align"
const WordWrapStyleAttr = "word-wrap"
const BreakWordWrapStyleAttrValue = "break-word"
const FontSizeStyleAttr = "font-size"
const FontWeightStyleAttr = "font-weight"
const BorderStyleAttr = "border-style"
const BorderStyleAttrValue = "solid"
const BorderInheritanceStyleAttr = "border-inheritance-type"
const BorderInheritanceStyleAttrValue = "solid"
const WidthStyleAttr = "width"
const MinWidthStyleAttr = "min-width"
const MaxWidthStyleAttr = "max-width"
const HeightStyleAttr = "height"
const MinHeightStyleAttr = "min-height"
const MaxHeightStyleAttr = "max-height"
const TextVerticalAlignStyleAttrValue = "center"
const ExcelBorderTypeValue = "thin"
const DefaultHorizontalAlignment = "left"


var XpathTable = xpath.Compile(".//table")
var XpathThead = xpath.Compile(".//thead/tr")
var XpathTh = xpath.Compile(".//th")
var XpathTr = xpath.Compile("./tr")
var XpathTd = xpath.Compile(".//td")


// Parsed style from html element
type HtmlStyle struct {
	TextAlign         string
	WordWrap          bool
	Width             float64
	Height            float64
	BorderInheritance bool
	BorderStyle       bool
	FontSize          int
	IsBold            bool
	Colspan           int
}

func NewHtmlStyle() *HtmlStyle {
	return &HtmlStyle {
		TextAlign:         "",
		WordWrap:          false,
		Width:             0,
		Height:            0,
		BorderInheritance: false,
		BorderStyle:       false,
		FontSize:          0,
		IsBold:            false,
		Colspan:           0,
	}
}


// Reads and unmarshalls json from file
func ReadJsonFile(jsonFilename string) map[string]interface{} {
	byteValue, _ := ioutil.ReadFile(jsonFilename)

	if byteValue == nil {
		LogMsg(fmt.Sprintf("File is empty? %s", jsonFilename))
	}

	var result map[string]interface{}
	json.Unmarshal(byteValue, &result)
	return result
}

func ReadHbsFile(filename string) string {
	byteValue, _ := ioutil.ReadFile(filename)

	if byteValue == nil {
		LogMsg(fmt.Sprintf("File is empty? %s", filename))
	}

	return string(byteValue)
}


// PrintMemUsage outputs the current, total and OS memory being used. As well as the number
// of garage collection cycles completed.
func PrintMemUsage() {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)
	// For info on each, see: https://golang.org/pkg/runtime/#MemStats
	log.Printf("\tAlloc = %v MiB", bToMb(m.Alloc))
	log.Printf("\tHeapAlloc = %v MiB", bToMb(m.HeapAlloc))
	log.Printf("\tSys = %v MiB", bToMb(m.Sys))
	log.Printf("\tFor info on each, see: https://golang.org/pkg/runtime/#MemStats\n")

}


// Converts bytes to human readable file size
func bToMb(b uint64) uint64 {
	return b / 1024 / 1024
}

// Log + console print
func LogMsg(s string) {
	log.Println(s)
	fmt.Println(s)
	//TODO: extend logging not only to file
}

// Log, print, die
func LogFatal(s string) {
	log.Println(s)
	fmt.Println(s)
	os.Exit(1)
}

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


// Apply style from html Td to excel Cell
func ApplyCellStyle(cell *xlsx.Cell, style *HtmlStyle) {
	if cell == nil {
		return
	}
	currentCellStyle := cell.GetStyle()
	setFont(currentCellStyle, style)
	setAlignment(currentCellStyle, style)
	setBorder(currentCellStyle, style)
	setColspan(cell, style)
}


// Apply style from html to excel column
func ApplyColumnStyle(cell *xlsx.Cell, sheet *xlsx.Sheet, colIndex int, style *HtmlStyle) {
	if cell == nil {
		return
	}
	currentCellStyle := cell.GetStyle()
	setColspan(cell, style)
	column := sheet.Col(colIndex-1)

	if column == nil {
		c := xlsx.NewColForRange(colIndex,colIndex)
		sheet.SetColParameters(c)
		columnAgain := sheet.Col(colIndex-1)

		if columnAgain == nil {
			panic("Error setting column styles")
		}
		column = columnAgain
		column.SetStyle(xlsx.NewStyle())
	}

	column.GetStyle().Alignment.WrapText = style.WordWrap
	column.GetStyle().ApplyAlignment = true

	if style.Width > 0 {
		sheet.SetColWidth(colIndex, colIndex, style.Width)
	}

	setFont(currentCellStyle, style)
	setBorder(currentCellStyle, style)
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

// Returns parsed style struct
func ExtractStyles(node *xml.AttributeNode) *HtmlStyle {
	if node == nil {
		return NewHtmlStyle()
	}
	styleStr := node.Content()
	entries := strings.Split(styleStr, ";")
	resultStyle := NewHtmlStyle()

	for _, e := range entries {
		if e != "" {
			parts := strings.Split(e, ":")
			value := strings.Trim(parts[1], " ")
			attr := strings.Trim(parts[0], " ")

			switch attr {
			case ColspanAttrName:
				resultStyle.Colspan, _ = strconv.Atoi(value)
			case TextAlignStyleAttr:
				resultStyle.TextAlign = value
			case WordWrapStyleAttr:
				resultStyle.WordWrap = value == BreakWordWrapStyleAttrValue
			case WidthStyleAttr:
				widthEntry := strings.Trim(value, " px")
				widthInt, _ := strconv.Atoi(widthEntry)
				translatedWidth := float64(widthInt) * PixelsToExcelWidthCoeff
				resultStyle.Width = translatedWidth

			case MinWidthStyleAttr:
				if resultStyle.Width <= 0 {
					widthEntry := strings.Trim(value, " px")
					widthInt, _ := strconv.Atoi(widthEntry)
					translatedWidth := float64(widthInt) * PixelsToExcelWidthCoeff
					resultStyle.Width = translatedWidth
				}
			case MaxWidthStyleAttr:
				if resultStyle.Width <= 0 {
					widthEntry := strings.Trim(value, " px")
					widthInt, _ := strconv.Atoi(widthEntry)
					translatedWidth := float64(widthInt) * PixelsToExcelWidthCoeff
					resultStyle.Width = translatedWidth
				}

			case HeightStyleAttr:
				heightEntry := strings.Trim(value, " px")
				heightInt, _ := strconv.Atoi(heightEntry)
				translatedHeight := float64(heightInt) * PixelsToExcelHeightCoeff
				resultStyle.Height = translatedHeight
			case MinHeightStyleAttr:
				if resultStyle.Height <= 0 {
					heightEntry := strings.Trim(value, " px")
					heightInt, _ := strconv.Atoi(heightEntry)
					translatedHeight := float64(heightInt) * PixelsToExcelHeightCoeff
					resultStyle.Height = translatedHeight
				}
			case MaxHeightStyleAttr:
				if resultStyle.Height <= 0 {
					heightEntry := strings.Trim(value, " px")
					heightInt, _ := strconv.Atoi(heightEntry)
					translatedHeight := float64(heightInt) * PixelsToExcelHeightCoeff
					resultStyle.Height = translatedHeight
				}
			case BorderStyleAttr:
				resultStyle.BorderStyle = value == BorderStyleAttrValue
			case BorderInheritanceStyleAttr:
				resultStyle.BorderStyle = value == BorderInheritanceStyleAttrValue
			case FontSizeStyleAttr:
				widthEntry := strings.Trim(value, " px")
				resultStyle.FontSize,_ = strconv.Atoi(widthEntry)
			case FontWeightStyleAttr:
				resultStyle.IsBold = strings.Contains(value, "bold")
			}
		}
	}
	return resultStyle
}

func main() {
	log.SetOutput(&lumberjack.Logger{
		Filename:   "html-to-xslx.log",
		MaxSize:    10, // megabytes
		MaxBackups: 3,
		MaxAge:     48, //days
		Compress:   false, // disabled by default
	})

	LogMsg("***********************************\n")
	LogMsg(fmt.Sprintf("HTML-TO-XSLX converter v.%d\n", Version))

	if len(os.Args) < 5 {
		fmt.Println("Usage:", os.Args[0], "hbs_template",  "data_json",
			"output_excel_file", "batch_size")
		LogFatal("Invalid command line args")
	}

	filename := os.Args[1]
	dataFilename := os.Args[2]
	outputFilename := os.Args[3]
	batchSize,_ := strconv.Atoi(os.Args[4])

	start := time.Now()
	renderedHtml := applyHandlebarsTemplate(filename, dataFilename)

	end := time.Now()

	timeResultStr := fmt.Sprintf("Elapsed time (Render Handlebars): %f s\n", end.Sub(start).Seconds())
	LogMsg(timeResultStr)
	LogMsg("Memory usage after rendering Handlebars.js template")
	PrintMemUsage()

	// Force GC to clear up, should see a memory drop
	//runtime.GC()
	//ioutil.WriteFile("rendered.html", []byte(renderedHtml), 0777)
	generateXlsxFile(renderedHtml, outputFilename, batchSize)

	end = time.Now()
	timeResultStr = fmt.Sprintf("Total elapsed time: %f s\n", end.Sub(start).Seconds())
	LogMsg(timeResultStr)
	LogMsg("Memory usage after all work done")
	PrintMemUsage()
}


// Parses given html and generated xslt file.
// File is generated by adding batches of batchSize to in on every iteration.
// On every iteration xlsx file is freshly opened and then saved with added data
func generateXlsxFile(html string, outputFilename string, batchSize int) string {
	start := time.Now()
	doc, err := gokogiri.ParseHtml([]byte(html))

	if err != nil {
		panic(err)
	}

	end := time.Now()
	LogMsg(fmt.Sprintf("Elapsed time (gokogiri html parsing): %f s\n", end.Sub(start).Seconds()))
	LogMsg("Html parsing: done. Starting xpath tables search")

	tables, _ := doc.Root().Search(XpathTable)
	defer doc.Free()
	xlsxFile := xlsx.NewFile()

	if xlsxFile == nil {
		panic("cant create excel file")
	}

	excelFilename := fmt.Sprintf("%s", outputFilename)
	start = time.Now()
	totalRows := 0

	// Main cycle through all tables in file
	for i, table := range tables {
		// Create new sheet for each table. Name it with data-name from html attribute
		sheetName := table.Attr("data-name")

		if sheetName == "" {
			sheetName = fmt.Sprintf("DataSheet %d", i)
			LogMsg(fmt.Sprintf("Warning! No data-name in for table found. Used %s as sheet name", sheetName))
		}

		if xlsxFile == nil {
			xlsxFile,err = xlsx.OpenFile(excelFilename)
		}


		sheet, err := xlsxFile.AddSheet(sheetName)

		if err != nil {
			panic(err)
		}

		if sheet == nil {
			LogFatal(fmt.Sprintf("Error: Cant add new sheet to xlsx file. %s", err))
		}

		// Get thead for table and create header in xlsx
		theadTrs, _ := table.Search(XpathThead)
		processHtmlTheadTag(theadTrs, sheet)

		// Get all rows in html table
		rows, _ := table.Search(XpathTr)
		rowsProceeded := 0
		packSize := batchSize

		for rowsProceeded < len(rows) {
			if xlsxFile == nil {
				xlsxFile,err = xlsx.OpenFile(excelFilename)
				sheet = xlsxFile.Sheet[sheetName]

				if sheet == nil {
					sheet, err = xlsxFile.AddSheet(sheetName)
				}
			}

			// processing one batch of rows
			// len(theadTrs) is an offset to skip (table headers)
			processTableRows(rows, sheet, len(theadTrs), rowsProceeded, packSize)

			err = xlsxFile.Save(excelFilename)
			xlsxFile = nil // help gc - prevent memory leak :)
			sheet = nil // help gc - prevent memory leak :)

			if err != nil {
				panic(err)
			}

			rowsProceeded += packSize
			runtime.GC() // prevent memory leak :)
		}

		totalRows += len(rows) // stored only for log output
		rows = nil // help gc - prevent memory leak :)
	}

	if err != nil {
		LogMsg(fmt.Sprintf("Error: can't save to %s", excelFilename))
		panic(err)
	}

	end = time.Now()
	LogMsg(fmt.Sprintf("Total elapsed time (main cycle): %f s\n", end.Sub(start).Seconds()))
	LogMsg(fmt.Sprintf("Total rows done: %d \n", totalRows))
	return excelFilename
}


// Process all html table rows
func processTableRows(rows []xml.Node, sheet *xlsx.Sheet, headerOffset int,  offset int, rowsNumber int) {
	if offset >= len(rows) {
		return
	}

	if len(rows) < rowsNumber {
		rowsNumber = len(rows) // when less than one page
	}

	for i := offset; i <= (offset + rowsNumber - 1); i++ {
		if i >= len(rows) {
			break
		}

		tr := rows[i]
		_ = sheet.AddRow()

		xlsxRow, e := sheet.AddRowAtIndex(headerOffset + i)

		if e != nil {
			LogMsg("Cant add row to excel sheet")
			panic(e)
		}

		// HEADERS
		theadTrs, _ := tr.Search(XpathTh)
		colIndex := 1

		for _, theadTh := range theadTrs {
			thStyle := theadTh.Attribute(StyleAttrName)
			xlsxCell := xlsxRow.AddCell()

			if thStyle != nil {
				style := ExtractStyles(thStyle)
				thColspan := theadTh.Attribute(ColspanAttrName)

				if thColspan != nil {
					style.Colspan,_ = strconv.Atoi(thColspan.Value())
				}

				ApplyColumnStyle(xlsxCell, sheet, colIndex, style)
				ApplyCellStyle(xlsxCell, style)
			} else {
				xlsxCell.SetStyle(xlsx.NewStyle())
			}

			xlsxCell.Value = theadTh.Content()
			colIndex++
		}

		cells, _ := tr.Search(XpathTd)

		for _, td := range cells {
			tdStyle := td.Attribute("style")
			xlsxCell := xlsxRow.AddCell()

			if tdStyle != nil {
				cellStyle := ExtractStyles(tdStyle)
				tdColspan := tdStyle.Attribute(ColspanAttrName)

				if tdColspan != nil {
					cellStyle.Colspan,_ = strconv.Atoi(tdColspan.Value())
				}

				ApplyCellStyle(xlsxCell, cellStyle)
			} else {
				xlsxCell.SetStyle(xlsx.NewStyle())
			}

			xlsxCell.Value = td.Content()
			xlsxCell = nil
		}

		trStyle := tr.Attribute("style")

		if trStyle != nil {
			styleExtracted := ExtractStyles(trStyle)
			ApplyRowStyle(xlsxRow, styleExtracted)
		} else {
			xlsxRow.SetHeight(15) // default
		}

		xlsxRow = nil
	}
}

// Process thead tag (thead->tr + thead->tr->th). Apply column styles. Apply cell styles
func processHtmlTheadTag(theadTrs []xml.Node, sheet *xlsx.Sheet) {
	for _, theadTr := range theadTrs {
		xlsxTheadRow := sheet.AddRow()
		theadTrThs, _ := theadTr.Search(XpathTh) // search for <th>
		colIndex := 1

		for _, theadTh := range theadTrThs { // for each <th> in <tr>
			thStyle := theadTh.Attribute(StyleAttrName)
			xlsxCell := xlsxTheadRow.AddCell()

			style := ExtractStyles(thStyle)
			if style != nil {
				ApplyColumnStyle(xlsxCell, sheet, colIndex, style)
				ApplyCellStyle(xlsxCell, style)
			} else {
				xlsxCell.SetStyle(xlsx.NewStyle())
			}

			thColspan := theadTh.Attribute(ColspanAttrName)

			if thColspan != nil {
				style.Colspan, _ = strconv.Atoi(thColspan.Value())
			}

			xlsxCell.Value = theadTh.Content()
			colIndex++
		}

		thStyle := theadTr.Attribute(StyleAttrName)

		if thStyle != nil {
			rowStyle := ExtractStyles(thStyle)
			thColspan := thStyle.Attribute(ColspanAttrName)

			if thColspan != nil {
				rowStyle.Colspan, _ = strconv.Atoi(thColspan.Value())
			}
			if rowStyle != nil {
				ApplyRowStyle(xlsxTheadRow, rowStyle)
			} else {
				xlsxTheadRow.SetHeight(15) // default
			}
		}
	}
}


// Apply Handlebars template to json data in dataFilename file.
// Note: takes content by "data" key
func applyHandlebarsTemplate(templateFilename string, dataFilename string) string {
	jsonCtx := ReadJsonFile(dataFilename)
	tpl, err := raymond.ParseFile(templateFilename)

	if err != nil {
		LogMsg(fmt.Sprintf("Error while parsing template %s ", templateFilename))
		panic(err)
	}

	// register helpers
	//TODO: we need more helpers
	tpl.RegisterHelper("math", func(x int, op string,  y int) string {
		if op == "+" {
			result := x + y
			return fmt.Sprintf("%d", result)
		}
		if op == "-" {
			result := x - y
			return fmt.Sprintf("%d", result)
		}
		if op == "*" {
			result := x * y
			return fmt.Sprintf("%d", result)
		}
		if op == "/" && y != 0 {
			result := x / y
			return fmt.Sprintf("%d", result)
		}
		return ""
	})

	data := jsonCtx
	result, err := tpl.Exec(data)

	if err != nil {
		LogMsg(fmt.Sprintf("Error while appying template %s to json file %s", templateFilename, dataFilename))
		panic(err)
	}

	return result
}


