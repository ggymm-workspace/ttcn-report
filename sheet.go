package main

import (
	"fmt"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func copySheet(src, dst *excelize.File, srcSheet, dstSheet string) error {
	dim, err := src.GetSheetDimension(srcSheet)
	if err != nil {
		return fmt.Errorf("get sheet dimension: %w", err)
	}
	if dim != "" {
		startCell, endCell := dim, dim
		if strings.Contains(dim, ":") {
			parts := strings.SplitN(dim, ":", 2)
			startCell = parts[0]
			endCell = parts[1]
		}
		startCol, startRow, err := excelize.CellNameToCoordinates(startCell)
		if err != nil {
			return fmt.Errorf("parse start cell %s: %w", startCell, err)
		}
		endCol, endRow, err := excelize.CellNameToCoordinates(endCell)
		if err != nil {
			return fmt.Errorf("parse end cell %s: %w", endCell, err)
		}

		styleMap := make(map[int]int)
		resolveStyle := func(styleID int) (int, error) {
			if styleID <= 0 {
				return 0, nil
			}
			if mapped, ok := styleMap[styleID]; ok {
				return mapped, nil
			}
			style, err := src.GetStyle(styleID)
			if err != nil {
				return 0, err
			}
			newID, err := dst.NewStyle(style)
			if err != nil {
				return 0, err
			}
			styleMap[styleID] = newID
			return newID, nil
		}

		for row := startRow; row <= endRow; row++ {
			for col := startCol; col <= endCol; col++ {
				cell, err := excelize.CoordinatesToCellName(col, row)
				if err != nil {
					continue
				}

				if formula, err := src.GetCellFormula(srcSheet, cell); err == nil && formula != "" {
					_ = dst.SetCellFormula(dstSheet, cell, formula)
				} else {
					raw, err := src.GetCellValue(srcSheet, cell, excelize.Options{RawCellValue: true})
					if err == nil && raw != "" {
						cellType, _ := src.GetCellType(srcSheet, cell)
						switch cellType {
						case excelize.CellTypeBool:
							val := raw == "1" || strings.EqualFold(raw, "true")
							_ = dst.SetCellValue(dstSheet, cell, val)
						case excelize.CellTypeNumber, excelize.CellTypeDate:
							if num, err := strconv.ParseFloat(raw, 64); err == nil {
								_ = dst.SetCellValue(dstSheet, cell, num)
							} else {
								_ = dst.SetCellValue(dstSheet, cell, raw)
							}
						default:
							_ = dst.SetCellValue(dstSheet, cell, raw)
						}
					}
				}

				styleID, err := src.GetCellStyle(srcSheet, cell)
				if err == nil && styleID > 0 {
					if newStyleID, err := resolveStyle(styleID); err == nil && newStyleID > 0 {
						_ = dst.SetCellStyle(dstSheet, cell, cell, newStyleID)
					}
				}
			}
		}

		for col := startCol; col <= endCol; col++ {
			colName, err := excelize.ColumnNumberToName(col)
			if err != nil {
				continue
			}
			if width, err := src.GetColWidth(srcSheet, colName); err == nil && width > 0 {
				_ = dst.SetColWidth(dstSheet, colName, colName, width)
			}
		}

		for row := startRow; row <= endRow; row++ {
			if height, err := src.GetRowHeight(srcSheet, row); err == nil && height > 0 {
				_ = dst.SetRowHeight(dstSheet, row, height)
			}
		}
	}

	if mergeCells, err := src.GetMergeCells(srcSheet); err == nil {
		for _, mc := range mergeCells {
			_ = dst.MergeCell(dstSheet, mc.GetStartAxis(), mc.GetEndAxis())
		}
	}

	if picCells, err := src.GetPictureCells(srcSheet); err == nil {
		for _, cell := range picCells {
			pics, err := src.GetPictures(srcSheet, cell)
			if err != nil {
				continue
			}
			for _, pic := range pics {
				pic.InsertType = excelize.PictureInsertTypePlaceOverCells
				_ = dst.AddPictureFromBytes(dstSheet, cell, &pic)
			}
		}
	}

	return nil
}
