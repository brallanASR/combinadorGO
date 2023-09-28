package main

import (
	"fmt"
	"math/rand"
	"time"

	excelize "github.com/xuri/excelize/v2"
)

const numProducts = 25

func main() {
	rand.Seed(time.Now().UnixNano())

	createExcelFile("fcc_producto1.xlsx")
	createExcelFile("fex_producto1.xlsx")
	createExcelFile("fex_producto2.xlsx")
	createExcelFile("fc_producto1.xlsx")
	createExcelFile("fc_producto2.xlsx")
}

func createExcelFile(filename string) {
	f := excelize.NewFile()

	// Set column headers
	headers := []string{"ID de Producto", "Nombre del Producto", "Marca", "Cantidad", "Unidad de Medida", "Área de Venta", "Fecha de Vencimiento"}
	for i, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue("Sheet1", cell, h)
	}

	for rowIndex := 2; rowIndex <= numProducts+1; rowIndex++ {
		f.SetCellValue("Sheet1", "A"+fmt.Sprint(rowIndex), rowIndex-1)
		f.SetCellValue("Sheet1", "B"+fmt.Sprint(rowIndex), "Producto "+randString(5))
		f.SetCellValue("Sheet1", "C"+fmt.Sprint(rowIndex), "Marca "+randString(3))
		f.SetCellValue("Sheet1", "D"+fmt.Sprint(rowIndex), rand.Intn(100))
		f.SetCellValue("Sheet1", "E"+fmt.Sprint(rowIndex), randomUnit())
		f.SetCellValue("Sheet1", "F"+fmt.Sprint(rowIndex), "Área "+fmt.Sprint(rand.Intn(5)))
		f.SetCellValue("Sheet1", "G"+fmt.Sprint(rowIndex), randomDate().Format("2006-01-02"))
	}

	// Save the file
	err := f.SaveAs(filename)
	if err != nil {
		panic(err)
	}
}

func randString(n int) string {
	const alphabet = "abcdefghijklmnopqrstuvwxyz"
	res := ""
	for i := 0; i < n; i++ {
		res += string(alphabet[rand.Intn(len(alphabet))])
	}
	return res
}

func randomUnit() string {
	units := []string{"Kg", "L", "Unidad", "Caja", "Paquete"}
	return units[rand.Intn(len(units))]
}

func randomDate() time.Time {
	return time.Now().AddDate(0, 0, rand.Intn(365))
}
