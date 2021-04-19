package helpers

import (
	"fmt"
	log "github.com/sirupsen/logrus"
	"strconv"
	"strings"
	"time"
)

func MathHelper(x int, op string,  y int) string {
	result := 0.0
	if op == "+" {
		result = float64(x + y)
	}
	if op == "-" {
		result = float64(x - y)
	}
	if op == "*" {
		result = float64(x * y)
	}
	if op == "/" && y != 0 {
		result = float64(x / y)
	}
	return fmt.Sprintf("%f", result)
}

func KeyHelper(x map[string]interface{}, key string) interface{} {
	return x[key]
}

func ZeroIntHelper(x string) string {
	if x == "" {
		return "00"
	}
	num,_ := strconv.Atoi(x)
	if num < 10 {
		return "0" + x
	}
	return x
}

func PercentHelper(x string) string {
	num, err := strconv.ParseFloat(x, 32)
	if err != nil {
		return ""
	}
	return fmt.Sprintf("%.2f%%", num * 100)
}

func InspectionTimeHelper(min string, sec string) string {
	if min == "" {
		min = "00"
	}
	if sec == "" {
		sec = "00"
	}
	return fmt.Sprintf("%s:%s", min, sec)
}

func DashHelper(x string) string {
	if x == "0" || x == "" {
		return "-"
	}
	return x
}

func PressureHelper(pressure string, upper string) string {
	if pressure != "" && upper != "" && upper != "0" {
		return pressure
	}
	return "-"
}

func AllowHelper(allow string) string {
	if allow == "Допущен" {
		return "Прошел"
	}
	return "Не прошел"
}

func UpperHelper(str string) string {
	return strings.ToUpper(str)
}

func IfNullHelper(obj interface{}, ifNull interface{}) interface{} {
	if obj == nil {
		return ifNull
	} else {
		return obj
	}
}

func IsAfterBeforeSheetHelper(sheetName string) bool {
	return strings.Contains(sheetName, "Предрейсовый") || strings.Contains(sheetName, "Послерейсовый") || strings.Contains(sheetName, "Линейный")
}

func SummarizeHelper(obj map[string]interface{}) interface{} {
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
}

func LineSumRowsHelper(obj map[string]interface{}) interface{} {
	inspectionsList := obj["inspections"].([]interface{})
	allCount := len(inspectionsList)
	goodLine := 0
	badLine := 0

	for _, inspection := range inspectionsList {
		inspAllow := inspection.(map[string]interface{})["allow"].(string)

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
}

func FaceIdNotFoundNameHelper(name string, surname string, patronymic string) string {
	if name != "" {
		return fmt.Sprintf("%s %s %s", surname, name, patronymic)
	}
	return "нет соответствия"
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

func FormatDate(dateStr string) string {
	if dateStr == "" {
		return ""
	}
	parsedDate, err := time.Parse("2006-01-02", dateStr)

	if err != nil {
		log.WithError(err).Error("Cant parse date from string")
	}

	formattedDate := parsedDate.Format("02-01-2006")
	return formattedDate
}

func FormatDateOfBirth(dateStr string) string {
	return FormatDate(dateStr)
}

func FormatGender (gender string) string {
	if gender == "MALE" {
		return "М"
	}
	return "Ж"
}

func FormatDateTime (dateStr string) string {
	if dateStr == "" {
		return ""
	}
	parsedDate, err := time.Parse("2006-01-02 15:04", dateStr)

	if err != nil {
		log.WithError(err).Error("Cant parse date from string")
	}

	formattedDate := parsedDate.Format("02-01-2006\n15:04")
	return formattedDate
}

func FormatOrganization (obj map[string]interface{}) interface{} {
	return obj["name"].(string)
}

func FormatType (inspectionType string) string {
	result := "Неизвестный"

	switch inspectionType {
		case "BEFORE": {
			result = "Предрейсовый"
			break
		}
		case "BEFORE_SHIFT": {
			result = "Предсменный"
			break
		}
		case "LINE": {
			result = "Линейный"
			break
		}
		case "AFTER": {
			result = "Послерейсовый"
			break
		}
		case "AFTER_SHIFT": {
			result = "Послесменный"
			break
		}
		case "ALCO": {
			result = "Алкотестирование"
			break
		}
		case "PIRO": {
			result = "Контроль температуры"
			break
		}
		case "PREVENTION": {
			result = "Профилактический"
			break
		}
	}

	return result
}

func FormatResult(result bool) string {
	if result {
		return "Допуск"
	}
	return "Не допуск"
}

func FormatComplains(complains string) string {
	if complains == "" {
		return "-"
	}
	if complains == "true" {
		return "Есть"
	}

	return "Нет"
}

func DashOrData (data interface{}) interface{}{
	if data == nil {
		return "-"
	}
	return data
}

func FormatPressure (obj map[string]interface{}) interface{} {
	if obj == nil {
		return "- / -"
	}
	systolic := DashOrData(obj["systolicPressure"])
	diastolic := DashOrData(obj["diastolicPressure"])
	result := fmt.Sprintf("%d / %d", systolic, diastolic)
	return result
}

func Sleep(sleep string) string {
	if sleep == "" {
		return "-"
	}
	if sleep == "true" {
		return "более 8 часов"
	}

	return "менее 8 часов"
}

