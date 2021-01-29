package main

import (
	"encoding/json"
	"fmt"
	"github.com/aymerick/raymond"
	"github.com/jbowtie/gokogiri"
	"github.com/jbowtie/gokogiri/xml"
	"github.com/jbowtie/gokogiri/xpath"
	_ "image"
	_ "image/jpeg"
	_ "image/png"
	"io/ioutil"
	"os"
	"runtime"
	"strconv"
	"strings"
	"time"
)


// App version
const Version = 8

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


var XpathTable = xpath.Compile(".//table")
var XpathThead = xpath.Compile(".//thead/tr")
var XpathTh = xpath.Compile(".//th")
var XpathTr = xpath.Compile("./tr")
var XpathTd = xpath.Compile(".//td")
var XpathImg = xpath.Compile(".//img")


// Mapped style from html element
type HtmlStyle struct {
	TextAlign         string
	WordWrap          bool
	Width             float64
	Height            float64
	BorderInheritance bool
	BorderStyle       bool
	FontSize          float64
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


// PrintMemUsage outputs the current, total and OS memory being used. As well as the number
// of garage collection cycles completed.
func PrintMemUsage() {
	var m runtime.MemStats
	runtime.ReadMemStats(&m)
	// For info on each, see: https://golang.org/pkg/runtime/#MemStats
	LogMsg(fmt.Sprintf("\tAlloc = %v MiB", bToMb(m.Alloc)))
	LogMsg(fmt.Sprintf("\tHeapAlloc = %v MiB", bToMb(m.HeapAlloc)))
	LogMsg(fmt.Sprintf("\tSys = %v MiB", bToMb(m.Sys)))
	LogMsg(fmt.Sprintf("\tFor info on each, see: https://golang.org/pkg/runtime/#MemStats\n"))
}


// Converts bytes to human readable file size
func bToMb(b uint64) uint64 {
	return b / 1024 / 1024
}

func LogMsg(s string) {
	//TODO: extended logging?
	fmt.Println(s)
}

// Log, print, die
func LogFatal(s string) {
	//TODO: extended logging?
	fmt.Println(s)
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
				sz,_ := strconv.Atoi(widthEntry)
				resultStyle.FontSize = float64(sz)
			case FontWeightStyleAttr:
				resultStyle.IsBold = strings.Contains(value, "bold")
			}
		}
	}
	return resultStyle
}

func main() {
	if len(os.Args) < 5 {
		fmt.Println("Usage:", os.Args[0], "hbs_template",  "data_json",
			"output_excel_file", "batch_size", "debug(0|1)")
		LogFatal("Invalid command line args")
	}

	LogMsg(fmt.Sprintf("HTML-TO-XSLX converter v.%d started", Version))
	debugOn := false // true turns on writing rendered html to file along with result excel file
	filename := os.Args[1]
	dataFilename := os.Args[2]
	outputFilename := os.Args[3]
	batchSize,_ := strconv.Atoi(os.Args[4])

	if len(os.Args) == 6 {
		debugOnArg,_ := strconv.Atoi(os.Args[5])
		if debugOnArg == 1 {
			debugOn = true
			LogMsg("debug mode on")
		}
	}

	start := time.Now()
	renderedHtml := applyHandlebarsTemplate(filename, dataFilename)
	end := time.Now()

	timeResultStr := fmt.Sprintf("Elapsed time (Render Handlebars): %f s\n", end.Sub(start).Seconds())
	LogMsg(timeResultStr)
	LogMsg("Memory usage after rendering Handlebars.js template")
	PrintMemUsage()

	if debugOn {
		ioutil.WriteFile("rendered.html", []byte(renderedHtml), 0777)
	}

	generateXlsxFile(renderedHtml, outputFilename, batchSize)

	end = time.Now()
	timeResultStr = fmt.Sprintf("Total elapsed time: %f s\n", end.Sub(start).Seconds())
	LogMsg(timeResultStr)
	LogMsg("Memory usage after all work done")
	PrintMemUsage()
}

func getExcelizeGenerator() *ExcelizeGenerator {
	return &ExcelizeGenerator{
		openedFile:   nil,
		filename:     "",
		currentSheet: "",
		currentCol:   0,
		currentRow:   0,
	}
}

// Parses given html and generated xslt file.
// File is generated by adding batches of batchSize to in on every iteration.
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

	// creating excel generator
	generator := getExcelizeGenerator()
	excelFilename := fmt.Sprintf("%s", outputFilename)
	generator.filename = excelFilename
	generator.currentCol = 1
	generator.currentRow = 1
	generator.Create()

	start = time.Now()
	totalRows := 0
	currentSheetIndex := 0

	// Main cycle through all tables in file
	for i, table := range tables {
		// Create new sheet for each table. Name it with data-name from html attribute
		sheetName := table.Attr("data-name")

		if sheetName == "" {
			sheetName = fmt.Sprintf("DataSheet %d", i)
			LogMsg(fmt.Sprintf("Warning! No data-name in for table found. Used %s as sheet name", sheetName))
		}

		if currentSheetIndex == 0 {
			generator.SetSheetName("Sheet1", sheetName)
		} else {
			generator.AddSheet(sheetName)
		}

		generator.currentCol = 1
		generator.currentRow = 0

		// Get thead for table and create header in xlsx
		theadTrs, _ := table.Search(XpathThead)
		processHtmlTheadTag(theadTrs, generator)

		// Get all rows in html table
		rows, _ := table.Search(XpathTr)
		rowsProceeded := 0
		packSize := batchSize

		for rowsProceeded < len(rows) {
			processTableRows(rows, generator, rowsProceeded, packSize)
			rowsProceeded += packSize
			runtime.GC() // prevent memory leak :)
		}

		totalRows += len(rows) // stored only for log output
		rows = nil // help gc - prevent memory leak :)
		currentSheetIndex += 1

	}

	generator.Save(generator.filename)

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
func processTableRows(rows []xml.Node, generator *ExcelizeGenerator, offset int, rowsNumber int) {
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
		generator.AddRow()

		// HEADERS
		theadTrs, _ := tr.Search(XpathTh)
		generator.currentCol = 1

		for _, theadTh := range theadTrs {
			thStyle := theadTh.Attribute(StyleAttrName)
			cellValue := theadTh.Content()

			if thStyle != nil {
				style := ExtractStyles(thStyle)
				thColspan := theadTh.Attribute(ColspanAttrName)

				if thColspan != nil {
					style.Colspan,_ = strconv.Atoi(thColspan.Value())
				}
				generator.ApplyColumnStyle(style)
				generator.ApplyCellStyle(style)
			}

			imgs, _ := theadTh.Search(XpathImg)

			if len(imgs) > 0 {
				for _, img := range imgs {
					imgSrc := img.Attribute("src")
					imgAlt := img.Attribute("alt")

					if _, err := os.Stat(imgSrc.Value()); os.IsNotExist(err) {
						if err != nil {
							fmt.Println(err)
						}
						generator.SetCellValue(imgAlt.Value())
					}

					currentCellCoords,errCoords := generator.GetCoords()

					if errCoords != nil {
						fmt.Println(errCoords)
					}

					errAdd := generator.openedFile.AddPicture(generator.currentSheet,
						currentCellCoords,
						imgSrc.Value(),
						`{"autofit":true, "positioning": "oneCell"}`)
					if errAdd != nil {
						fmt.Println(errAdd)
					}
				}
			} else {
				if cellValue != "" {
					generator.SetCellValue(cellValue)
				}
			}

			generator.currentCol += 1
		}

		cells, _ := tr.Search(XpathTd)
		generator.currentCol = 1

		// Table cells
		for _, td := range cells {
			tdStyle := td.Attribute("style")
			cellValue := td.Content()

			if tdStyle != nil {
				cellStyle := ExtractStyles(tdStyle)
				tdColspan := td.Attribute(ColspanAttrName)

				if tdColspan != nil {
					cellStyle.Colspan,_ = strconv.Atoi(tdColspan.Value())
				}

				generator.ApplyCellStyle(cellStyle)
			}

			imgs, _ := td.Search(XpathImg)

			if len(imgs) > 0 {
				for _, img := range imgs {
					imgSrc := img.Attribute("src")
					imgAlt := img.Attribute("alt")

					if _, err := os.Stat(imgSrc.Value()); os.IsNotExist(err) {
						if err != nil {
							fmt.Println(err)
						}
						generator.SetCellValue(imgAlt.Value())
					}

					currentCellCoords,errCoords := generator.GetCoords()

					if errCoords != nil {
						fmt.Println(errCoords)
					}

					errAdd := generator.openedFile.AddPicture(generator.currentSheet,
						currentCellCoords,
						imgSrc.Value(),
						`{"autofit":true, "lock_aspect_ratio": false, "locked": false, "positioning": "oneCell"}`)
					if errAdd != nil {
						fmt.Println(errAdd)
					}

				}
			} else {
				if cellValue != "" {
					generator.SetCellValue(cellValue)
				}
			}

			generator.currentCol += 1
		}

		trStyle := tr.Attribute("style")

		// Apply row style if present
		if trStyle != nil {
			styleExtracted := ExtractStyles(trStyle)
			generator.ApplyRowStyle(styleExtracted)
		}
	}
}

// Process thead tag (thead->tr + thead->tr->th). Apply column styles. Apply cell styles
func processHtmlTheadTag(theadTrs []xml.Node, generator *ExcelizeGenerator) {
	for _, theadTr := range theadTrs {
		generator.AddRow()
		theadTrThs, _ := theadTr.Search(XpathTh) // search for <th>
		colIndex := 1

		for _, theadTh := range theadTrThs { // for each <th> in <tr>
			thStyle := theadTh.Attribute(StyleAttrName)

			style := ExtractStyles(thStyle)
			thColspan := theadTh.Attribute(ColspanAttrName)

			if thColspan != nil {
				style.Colspan, _ = strconv.Atoi(thColspan.Value())
			}

			content := theadTh.Content()

			if content != "" {
				generator.SetCellValue(content)
			}

			if style != nil {
				generator.ApplyColumnStyle(style)
				generator.ApplyCellStyle(style)

			}

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
				generator.ApplyRowStyle(rowStyle)
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
	registerAllHelpers(tpl)
	data := jsonCtx
	result, err := tpl.Exec(data)

	if err != nil {
		LogMsg(fmt.Sprintf("Error while appying template %s to json file %s", templateFilename, dataFilename))
		panic(err)
	}

	return result
}

func registerAllHelpers(template *raymond.Template)  {
	// math
	template.RegisterHelper("math", func(x int, op string,  y int) string {
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

	// key
	template.RegisterHelper("key", func(x map[string]interface{}, key string) interface{} {
		return x[key]
	})

	// zeroIntHelper
	template.RegisterHelper("zeroIntHelper", func(x string) string {
		if x == "" {
			return "00"
		}

		num,_ := strconv.Atoi(x)

		if num < 10 {
			return "0" + x
		}

		return x
	})

	template.RegisterHelper("percentHelper", func(x string) string {
		num, err := strconv.ParseFloat(x, 32)
		if err != nil {
			return ""
		}
		return fmt.Sprintf("%.2f%%", num * 100)
	})

	template.RegisterHelper("inspectionTimeHelper", func(min string, sec string) string {
		if min == "" {
			min = "00"
		}
		if sec == "" {
			sec = "00"
		}
		return fmt.Sprintf("%s:%s", min, sec);
	})

	template.RegisterHelper("dashHelper", func(x string) string {
		if x == "0" || x == "" {
			return "-"
		}
		return x
	})

	template.RegisterHelper("pressureHelper", func(pressure string, upper string) string {
		if pressure != "" && upper != "" && upper != "0" {
			return pressure
		}
		return "-"
	})

	template.RegisterHelper("allowHelper", func(allow string) string {
		if allow == "Допущен" {
			return "Прошел"
		}
		return "Не прошел"
	})

	template.RegisterHelper("upper", func(str string) string {
		return strings.ToUpper(str)
	})

	template.RegisterHelper("ifnull", func(obj interface{}, ifNull interface{}) interface{} {
		if obj == nil {
			return ifNull
		} else {
			return obj
		}
	})

	template.RegisterHelper("isAfterBeforeSheet", func(sheetName string) bool {
		return strings.Contains(sheetName, "Предрейсовый") || strings.Contains(sheetName, "Послерейсовый") || strings.Contains(sheetName, "Линейный")
	})

	template.RegisterHelper("summarize", func(obj map[string]interface{}) interface{} {
		sheetNameInterface := obj["sheetName"]
		sheetName := fmt.Sprintf("%v", sheetNameInterface)

		if strings.Contains(sheetName, "Предрейсовый") {
			return beforeSumRows(obj)
		}

		if strings.Contains(sheetName, "Послерейсовый") {
			return afterSumRows(obj)
		} else {
			return ""
		}
	})

	template.RegisterHelper("lineSumRows", func(obj map[string]interface{}) interface{} {
		inspectionsList := obj["inspections"].([]interface{})
		allCount := len(inspectionsList)
		goodLine := 0
		badLine := 0

		for _,inspection := range inspectionsList {
			inspAllow := inspection.(map[string] interface{})["allow"].(string)

			if inspAllow == "Допущен" {
				goodLine++
			} else {
				badLine++
			}
		}

		return fmt.Sprintf(`<tr style="height: 300px; border-style: solid"><td colspan="4" style="text-align: left">Итого осмотрено: </td><td></td><td></td><td></td><td>%d</td></tr>
    			<tr style="height: 300px; border-style: solid"><td colspan="4" style="text-align: left">Итого прошло линейный контроль: </td><td></td><td></td><td></td><td>%d</td></tr>
    			<tr style="height: 300px; border-style: solid"><td colspan="4" style="text-align: left">Итого отстраненных от трудовых обязанностей: </td><td></td><td></td><td></td><td>%d</td></tr>`,
				allCount, goodLine, badLine)
	})

	template.RegisterHelper("faceIdNotFoundName", func(name string, surname string, patronymic string) string {
		if name != "" {
			return fmt.Sprintf("%s %s %s", surname, name, patronymic)
		}
		return "нет соответствия"
	})


}

func beforeSumRows(obj map[string]interface{}) interface{} {
	inspectionsList := obj["inspections"].([]interface{})
	allCount := len(inspectionsList)
	goodBefore := 0
	badBefore := 0
	goodLine := 0
	badLine := 0

	for _,inspection := range inspectionsList {
		inspType := inspection.(map[string] interface{})["type"].(string)
		inspAllow := inspection.(map[string] interface{})["allow"].(string)

		if inspType == "Предрейсовый" || inspType == "Предсменный" {
			if inspAllow == "Допущен" {
				goodBefore++
			} else {
				badBefore++
			}
		} else {
			if inspAllow == "Допущен" {
				goodLine++
			} else {
				badLine++
			}
		}
	}

	return fmt.Sprintf(`<tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого осмотрено: </td><td></td><td></td><td></td><td>%d</td></tr>
		<tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого допущено к исполнению трудовых обязанностей: </td><td></td><td></td><td></td><td>%d</td></tr>
		<tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого не допущено к исполнению трудовых обязанностей: </td><td></td><td></td><td></td><td>%d</td></tr>`,
		allCount, goodBefore, badBefore)
}

func afterSumRows(obj map[string]interface{}) interface{} {
	inspectionsList := obj["inspections"].([]interface{})
	allCount := len(inspectionsList)
	goodAfter := 0
	badAfter := 0
	goodAftershift := 0
	badAftershift := 0

	for _,inspection := range inspectionsList {
		inspType := inspection.(map[string] interface{})["type"].(string)
		inspAllow := inspection.(map[string] interface{})["allow"].(string)

		if inspType == "Послерейсовый" {
			if inspAllow == "Допущен" {
				goodAfter++
			} else {
				badAfter++
			}
		} else {
			if inspAllow == "Допущен" {
				goodAftershift++
			} else {
				badAftershift++
			}
		}
	}

	return fmt.Sprintf(`<tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого осмотрено: </td><td></td><td></td><td></td><td>%d</td></tr>
        <tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого прошло послерейсовый: </td><td></td><td></td><td></td><td>%d</td></tr>
        <tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого прошло послесменный: </td><td></td><td></td><td></td><td>%d</td></tr>
        <tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого не прошло послерейсовый: </td><td></td><td></td><td></td><td>%d</td></tr>
        <tr style="height: 300px; border-style: solid;"><td colspan="4" style="text-align: left">Итого не прошло послесменный: </td><td></td><td></td><td></td><td>%d</td></tr>`,
		allCount, goodAfter, goodAftershift, badAfter, badAftershift)
}

