package main

import (
	"fmt"
	"os"
	"path/filepath"
	"runtime"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

type HeaderExcel map[string]string

var (
	_, b, _, _ = runtime.Caller(0)
	basepath   = filepath.Dir(b)
)

type ProductExcel struct {
	Name     string
	Category string
	Price    string
	Dose     string
	Factory  string
	Usage    string
	Desc     string
}

func main() {

	// createXlxs()
	readTheFile()

	// time.Sleep(4 * time.Second)
	// removeFile(patname)
}

func removeFile(filePath string) {
	os.Remove(filePath)
}

func readTheFile() {
	// Read file
	// sheet name would be `product`

	pathFile := filepath.Join(basepath, "excel/workbook.xlsx")

	file, err := excelize.OpenFile(pathFile)
	if err != nil {
		fmt.Println("Error: ", err.Error())
		return
	}

	rows, err := file.GetRows("product")
	if err != nil {
		fmt.Println("Error: ", err.Error())
		return
	}

	products := []ProductExcel{}
	for idx, row := range rows {
		if idx == 0 {
			continue
		}
		if len(row[0]) <= 0 {
			break
		}

		valRow := &ProductExcel{}
		for i, colCell := range row {
			switch i {
			case 0:
				valRow.Name = colCell
			case 1:
				valRow.Category = colCell
			case 2:
				valRow.Price = colCell
			case 3:
				valRow.Dose = colCell
			case 4:
				valRow.Factory = colCell
			case 5:
				valRow.Usage = colCell
			case 6:
				valRow.Desc = colCell
			}
		}

		products = append(products, *valRow)
	}

	fmt.Println(products)

}

func createXlxs() string {
	fx := excelize.NewFile()
	fmt.Println("Creating file...")
	var (
		headerStyle int
		sheetName   = "product"
		err         error
	)

	// renaming shee name
	// default value from Sheet1
	fx.SetSheetName("Sheet1", sheetName)

	// Header Value
	headerValue := HeaderExcel{
		"A1": "Nama",
		"B1": "Kategori",
		"C1": "Harga Jual",
		"D1": "Dosis",
		"E1": "Pabrik",
		"F1": "Kegunaan",
		"G1": "Deskripsi",
	}

	for key, val := range headerValue {
		fx.SetCellValue(sheetName, key, val)
	}

	fileSaveName := filepath.Join(basepath, "excel/workbook.xlsx")

	// define the border style
	border := []excelize.Border{
		{Type: "top", Style: 1, Color: "cccccc"},
		{Type: "left", Style: 1, Color: "cccccc"},
		{Type: "right", Style: 1, Color: "cccccc"},
		{Type: "bottom", Style: 1, Color: "cccccc"},
	}

	// header style
	// define the style of the header row
	if headerStyle, err = fx.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true},
		Fill: excelize.Fill{
			Type: "pattern", Color: []string{"dae9f3"}, Pattern: 1},
		Border: border},
	); err != nil {
		fmt.Println("Error: ", err.Error())
		return "Error 1"
	}

	if err != nil {
		fmt.Println(err.Error())
		return "Error 2"
	}

	if err = fx.SetCellStyle(sheetName, "A1", "Z1", headerStyle); err != nil {
		fmt.Println(err.Error())
		return "Error 3"
	}

	// Number Format
	numFormat, err := fx.NewStyle(&excelize.Style{
		NumFmt: 359,
	})

	if err != nil {
		fmt.Println(err.Error())
		return "Error"
	}

	if err = fx.SetCellStyle(sheetName, "C2", "C100", numFormat); err != nil {
		fmt.Println(err.Error())
		return "Error "
	}

	// Validation Category
	dvCat := excelize.NewDataValidation(true)
	dvCat.Sqref = "B2:B100"
	dvCat.SetDropList([]string{"Obat Khusus", "Obat Puyeng"})
	dvCat.ShowErrorMessage = true

	fx.AddDataValidation(sheetName, dvCat)

	if err := fx.SaveAs(fileSaveName); err != nil {
		fmt.Println("Error: ", err.Error())
		return "Error 2"
	}

	fmt.Println(fileSaveName)
	return fileSaveName
}
