package main

import (
	"bytes"
	"embed"

	"github.com/xuri/excelize/v2"
)

//go:embed 7.5.1.xlsx
//go:embed 7.5.2.xlsx
//go:embed 7.6.2.xlsx
//go:embed 7.6.3.xlsx
//go:embed 7.6.4.xlsx
//go:embed 7.6.5.1.xlsx
//go:embed 7.6.5.2.xlsx
//go:embed 7.6.5.3.xlsx
//go:embed 7.6.5.4.xlsx
//go:embed 7.6.7.xlsx
//go:embed 7.6.9.xlsx
//go:embed 7.6.10.xlsx
//go:embed 7.6.11.xlsx
//go:embed 7.6.12.xlsx
//go:embed 7.6.13.xlsx
//go:embed 7.6.14.xlsx
//go:embed 7.6.15.xlsx
//go:embed 7.6.17.xlsx
//go:embed 7.6.18.xlsx
//go:embed 7.6.19.xlsx
var templates embed.FS

func openTpl(filename string) (*excelize.File, error) {
	data, err := templates.ReadFile(filename)
	if err != nil {
		panic(err)
	}
	return excelize.OpenReader(bytes.NewReader(data))
}
