package main

import (
	"fmt"
	"slices"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func copySheet(src, dst *excelize.File, srcSheet, dstSheet string) error {
	// 防止 sheet 名称重复
	i := 1
	for {
		if !slices.Contains(dst.GetSheetList(), dstSheet+fmt.Sprintf("(%d)", i)) {
			idx, err := dst.NewSheet(dstSheet)
			if err != nil {
				return err
			}
			dst.SetActiveSheet(idx)
			break
		}
		i++
	}

	// 复制 sheet 到目标文件
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
		startCol, startRow, err1 := excelize.CellNameToCoordinates(startCell)
		if err1 != nil {
			return fmt.Errorf("parse start cell %s: %w", startCell, err1)
		}
		endCol, endRow, err1 := excelize.CellNameToCoordinates(endCell)
		if err1 != nil {
			return fmt.Errorf("parse end cell %s: %w", endCell, err1)
		}

		styleMap := make(map[int]int)
		resolveStyle := func(styleID int) (int, error) {
			if styleID <= 0 {
				return 0, nil
			}
			if mapped, ok := styleMap[styleID]; ok {
				return mapped, nil
			}
			style, err2 := src.GetStyle(styleID)
			if err2 != nil {
				return 0, err2
			}
			newID, err2 := dst.NewStyle(style)
			if err2 != nil {
				return 0, err2
			}
			styleMap[styleID] = newID
			return newID, nil
		}

		for row := startRow; row <= endRow; row++ {
			for col := startCol; col <= endCol; col++ {
				cell, err2 := excelize.CoordinatesToCellName(col, row)
				if err2 != nil {
					continue
				}

				if formula, err3 := src.GetCellFormula(srcSheet, cell); err3 == nil && formula != "" {
					_ = dst.SetCellFormula(dstSheet, cell, formula)
				} else {
					raw, err4 := src.GetCellValue(srcSheet, cell, excelize.Options{RawCellValue: true})
					if err4 == nil && raw != "" {
						cellType, _ := src.GetCellType(srcSheet, cell)
						switch cellType {
						case excelize.CellTypeBool:
							val := raw == "1" || strings.EqualFold(raw, "true")
							_ = dst.SetCellValue(dstSheet, cell, val)
						case excelize.CellTypeNumber, excelize.CellTypeDate:
							if num, err5 := strconv.ParseFloat(raw, 64); err5 == nil {
								_ = dst.SetCellValue(dstSheet, cell, num)
							} else {
								_ = dst.SetCellValue(dstSheet, cell, raw)
							}
						default:
							_ = dst.SetCellValue(dstSheet, cell, raw)
						}
					}
				}

				styleID, err2 := src.GetCellStyle(srcSheet, cell)
				if err2 == nil && styleID > 0 {
					if newStyleID, err3 := resolveStyle(styleID); err3 == nil && newStyleID > 0 {
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
