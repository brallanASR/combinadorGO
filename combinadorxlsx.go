package main

import (
	"fmt"
	"io/ioutil"
	"log"
	"strings"
	"time"

	excelize "github.com/xuri/excelize/v2"
)

func main() {
	files, err := ioutil.ReadDir(".")
	if err != nil {
		log.Fatalf("Error al leer el directorio: %v", err)
	}

	categories := []string{"fcc", "fex", "fc"}
	workbooks := make(map[string]*excelize.File)

	for _, category := range categories {
		workbooks[category] = excelize.NewFile()
		firstFile := true

		for _, file := range files {
			filename := file.Name()
			baseName := strings.TrimSuffix(filename, ".xlsx")

			if strings.Contains(baseName, category) && !strings.Contains(baseName, category+"c") {
				log.Printf("Añadiendo archivo %s a la categoría '%s'\n", filename, category)

				if firstFile {
					newSheetName := baseName
					sheetIndex, err := workbooks[category].GetSheetIndex("Sheet1")
					if err != nil {
						log.Fatalf("Error al obtener el índice de la hoja 'Sheet1': %v", err)
					}
					workbooks[category].SetActiveSheet(sheetIndex)
					workbooks[category].SetSheetName("Sheet1", newSheetName)
					processFileToSheet(filename, workbooks[category], newSheetName)
					firstFile = false
				} else {
					processFile(filename, workbooks[category])
				}
			}
		}

		if !firstFile {
			outputFilename := fmt.Sprintf("documento combinado %s %s.xlsx", category, time.Now().Format("2006-01-02_15-04-05"))
			err := workbooks[category].SaveAs(outputFilename)
			if err != nil {
				log.Printf("Error al guardar el archivo combinado %s: %v\n", outputFilename, err)
			} else {
				log.Printf("Archivo combinado creado: %s\n", outputFilename)
			}
		}
	}
}

func processFile(filename string, combinedFile *excelize.File) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Fatalf("Error al abrir el archivo %s: %v", filename, err)
	}

	sheetName := f.GetSheetName(0)
	newSheetName := strings.TrimSuffix(filename, ".xlsx")
	rows, err := f.GetRows(sheetName)
	if err != nil {
		log.Fatalf("Error al obtener filas del archivo %s: %v", filename, err)
	}

	combinedFile.NewSheet(newSheetName)
	rowIndex := 1
	for _, row := range rows {
		combinedFile.SetSheetRow(newSheetName, indexToColumnLetter(len(row))+fmt.Sprintf("%d", rowIndex), &row)
		rowIndex++
	}
}

func processFileToSheet(filename string, combinedFile *excelize.File, sheetName string) {
	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Fatalf("Error al abrir el archivo %s: %v", filename, err)
	}

	sourceSheetName := f.GetSheetName(0)
	rows, err := f.GetRows(sourceSheetName)
	if err != nil {
		log.Fatalf("Error al obtener filas del archivo %s: %v", filename, err)
	}

	rowIndex := 1
	for _, row := range rows {
		combinedFile.SetSheetRow(sheetName, indexToColumnLetter(len(row))+fmt.Sprintf("%d", rowIndex), &row)
		rowIndex++
	}
}

func indexToColumnLetter(index int) string {
	column := ""
	for index > 0 {
		index--
		column = string('A'+index%26) + column
		index /= 26
	}
	return column
}
