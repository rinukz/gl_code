package main

import (
	"github.com/xuri/excelize/v2"
)

const excelPass = "1234"

func OpenMyExcel() (*excelize.File, error) {
	return excelize.OpenFile("RinukzGL.xlsx", excelize.Options{Password: excelPass})
}

func SaveMyExcel(f *excelize.File) error {
	return f.SaveAs("RinukzGL.xlsx", excelize.Options{Password: excelPass})
}
