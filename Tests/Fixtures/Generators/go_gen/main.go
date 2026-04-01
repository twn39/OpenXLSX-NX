package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
)

func main() {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	f.SetSheetName("Sheet1", "TestSheet")
	f.SetCellValue("TestSheet", "A1", "Hello from excelize")
	f.SetCellValue("TestSheet", "A2", 42)
	f.SetCellValue("TestSheet", "A3", 3.1415)

	f.SetCellValue("TestSheet", "B1", "Styled")
	style, err := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Color: "FF0000"},
		Fill: excelize.Fill{Type: "pattern", Color: []string{"00FF00"}, Pattern: 1},
	})
	if err == nil {
		f.SetCellStyle("TestSheet", "B1", "B1", style)
	}

	if err := f.SaveAs("../../excelize_generated.xlsx"); err != nil {
		fmt.Println(err)
	}
}
