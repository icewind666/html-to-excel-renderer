package main

import (
	"encoding/json"
	"fmt"
	"github.com/aymerick/raymond"
	"github.com/icewind666/html-to-excel-render/src/generator"
	"github.com/icewind666/html-to-excel-render/src/helpers"
	"github.com/icewind666/html-to-excel-render/src/types"
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


var (
	version = "1.1.3"
	commit  = "-"
	date    = "03-2021"
	builtBy = "v.korennoj"
)

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
const TextVerticalAlignStyleAttr = "vertical-align" // values can be: top | middle | bottom | baseline


var XpathTable = xpath.Compile(".//table")
var XpathThead = xpath.Compile(".//thead/tr")
var XpathTh = xpath.Compile(".//th")
var XpathTr = xpath.Compile("./tr")
var XpathTd = xpath.Compile(".//td")
var XpathImg = xpath.Compile(".//img")


func main() {
	if len(os.Args) < 5 {
		fmt.Println("Usage:", os.Args[0], "hbs_template",  "data_json",
			"output_excel_file", "batch_size", "debug(0|1)")
		LogFatal("Invalid command line args")
	}

	LogMsg(fmt.Sprintf("html-to-excel-renderer v%s, commit %s, built at %s by %s", version, commit, date, builtBy))
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
		err := ioutil.WriteFile("rendered.html", []byte(renderedHtml), 0777)
		if err != nil {
			LogMsg("Cant write debug log - rendered html file!")
		}
	}

	generateXlsxFile(renderedHtml, outputFilename, batchSize)

	end = time.Now()
	timeResultStr = fmt.Sprintf("Total elapsed time: %f s\n", end.Sub(start).Seconds())
	LogMsg(timeResultStr)
	LogMsg("Memory usage after all work done")
	PrintMemUsage()
}

func NewHtmlStyle() *types.HtmlStyle {
	return &types.HtmlStyle {
		TextAlign:         "",
		WordWrap:          false,
		Width:             0,
		Height:            0,
		BorderInheritance: false,
		BorderStyle:       false,
		FontSize:          0,
		IsBold:            false,
		Colspan:           0,
		VerticalAlign:     "",
	}
}


// Reads and unmarshalls json from file
func ReadJsonFile(jsonFilename string) map[string]interface{} {
	byteValue, _ := ioutil.ReadFile(jsonFilename)

	if byteValue == nil {
		LogMsg(fmt.Sprintf("File is empty? %s", jsonFilename))
	}

	var result map[string]interface{}

	err := json.Unmarshal(byteValue, &result)
	if err != nil {
		LogMsg("Error reading data json file! Can't deserialize json!")
	}

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

// Logging string.
// Extension point for additional logging. For now its stdout
func LogMsg(s string) {
	//TODO: extended logging?
	fmt.Println(s)
}

// Log, print, die. Stdout
func LogFatal(s string) {
	//TODO: extended logging?
	fmt.Println(s)
	os.Exit(-1)
}




func getExcelizeGenerator() *generator.ExcelizeGenerator {
	return &generator.ExcelizeGenerator{
		OpenedFile:   nil,
		Filename:     "",
		CurrentSheet: "",
		CurrentCol:   0,
		CurrentRow:   0,
	}
}

// Parses given html and generates xslt file.
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

	// creating excel excelizeGenerator
	excelizeGenerator := getExcelizeGenerator()
	excelFilename := fmt.Sprintf("%s", outputFilename)
	excelizeGenerator.Filename = excelFilename
	excelizeGenerator.CurrentCol = 1
	excelizeGenerator.CurrentRow = 1
	excelizeGenerator.Create()

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
			excelizeGenerator.SetSheetName("Sheet1", sheetName)
		} else {
			excelizeGenerator.AddSheet(sheetName)
		}

		excelizeGenerator.CurrentCol = 1
		excelizeGenerator.CurrentRow = 0

		// Get thead for table and create header in xlsx
		theadTrs, _ := table.Search(XpathThead)
		processHtmlTheadTag(theadTrs, excelizeGenerator)

		// Get all rows in html table
		rows, _ := table.Search(XpathTr)
		rowsProceeded := 0
		packSize := batchSize

		for rowsProceeded < len(rows) {
			processTableRows(rows, excelizeGenerator, rowsProceeded, packSize)
			rowsProceeded += packSize
			runtime.GC() // prevent memory leak :)
		}

		totalRows += len(rows) // stored only for log output
		rows = nil // help gc - prevent memory leak :)
		currentSheetIndex += 1
	}

	excelizeGenerator.Save(excelizeGenerator.Filename)

	end = time.Now()
	LogMsg(fmt.Sprintf("Total elapsed time (main cycle): %f s\n", end.Sub(start).Seconds()))
	LogMsg(fmt.Sprintf("Total rows done: %d \n", totalRows))
	return excelFilename
}


// Process all html table rows
func processTableRows(rows []xml.Node, generator *generator.ExcelizeGenerator, offset int, rowsNumber int) {
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
		generator.CurrentCol = 1

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

					errAdd := generator.OpenedFile.AddPicture(generator.CurrentSheet,
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

			generator.CurrentCol += 1
		}

		cells, _ := tr.Search(XpathTd)
		generator.CurrentCol = 1

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

					errAdd := generator.OpenedFile.AddPicture(generator.CurrentSheet,
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

			generator.CurrentCol += 1
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
func processHtmlTheadTag(theadTrs []xml.Node, generator *generator.ExcelizeGenerator) {
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
	template.RegisterHelper("math", helpers.MathHelper)
	template.RegisterHelper("key", helpers.KeyHelper)
	template.RegisterHelper("zeroIntHelper", helpers.ZeroIntHelper)
	template.RegisterHelper("percentHelper", helpers.PercentHelper)
	template.RegisterHelper("inspectionTimeHelper", helpers.InspectionTimeHelper)
	template.RegisterHelper("dashHelper", helpers.DashHelper)
	template.RegisterHelper("pressureHelper", helpers.PressureHelper)
	template.RegisterHelper("allowHelper", helpers.AllowHelper)
	template.RegisterHelper("upper", helpers.UpperHelper)
	template.RegisterHelper("ifnull", helpers.IfNullHelper)
	template.RegisterHelper("isAfterBeforeSheet", helpers.IsAfterBeforeSheetHelper)
	template.RegisterHelper("summarize", helpers.SummarizeHelper)
	template.RegisterHelper("lineSumRows", helpers.LineSumRowsHelper)
	template.RegisterHelper("faceIdNotFoundName", helpers.FaceIdNotFoundNameHelpder)
}



// Returns parsed style struct
func ExtractStyles(node *xml.AttributeNode) *types.HtmlStyle {
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

			case TextVerticalAlignStyleAttr:
				resultStyle.VerticalAlign = value
			}

		}
	}
	return resultStyle
}
