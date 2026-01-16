package main

import (
	_ "image/png"

	"image"
	"strings"

	"github.com/xuri/excelize/v2"
)

func setCell(tpl *excelize.File, params map[string]string) error {
	rows, err := tpl.GetRows(sheet)
	if err != nil {
		return err
	}

	for i := range rows {
		for j := range rows[i] {
			cell, _ := excelize.CoordinatesToCellName(j+1, i+1)
			origin, _ := tpl.GetCellValue(sheet, cell)

			for label, value := range params {
				if !strings.Contains(origin, label) {
					continue
				}
				// 处理图片
				if strings.HasSuffix(label, "（图片）") {
					if len(value) == 0 || !exists(value) {
						// 如果图片地址为空
						// 如果图片文件不存在，则不执行后续的逻辑
						continue
					}
					err = tpl.SetCellValue(sheet, cell, "") // 清空文字
					if err != nil {
						return err
					}

					// 获取图片尺寸
					width, height := imageCellSize(value)

					// 更新单元格尺寸
					err = updateImageCellSize(tpl, j+1, i+1, width, height)
					if err != nil {
						return err
					}

					// 添加图片
					err = tpl.AddPicture(sheet, cell, value, &excelize.GraphicOptions{
						AutoFit: true,
					})
					if err != nil {
						return err
					}
				} else {
					err = tpl.SetCellValue(sheet, cell, value)
					if err != nil {
						return err
					}
				}
			}
		}
	}
	return nil
}

func imageCellSize(p string) (float64, float64) {
	f := open(p)
	defer func() {
		_ = f.Close()
	}()
	c, _, err := image.DecodeConfig(f)
	if err != nil {
		return 100, 100
	}

	w := float64(c.Width) * 1.5
	h := float64(c.Height) * 1.5

	cellWidth := (w - 0.5) / 8.0
	if cellWidth < 0 {
		cellWidth = 0
	}
	cellHeight := h * 3.4 / 4.0

	// 返回单元格实际宽和高
	return min(cellWidth, 255), min(cellHeight, 409)
}

func updateImageCellSize(tpl *excelize.File, col, row int, width, height float64) error {
	ranges := make([]int, 0)
	mergeCells, _ := tpl.GetMergeCells(sheet)
	for _, cell := range mergeCells {
		startCol, startRow, err := excelize.CellNameToCoordinates(cell.GetStartAxis())
		if err != nil {
			continue
		}
		endCol, endRow, err := excelize.CellNameToCoordinates(cell.GetEndAxis())
		if err != nil {
			continue
		}
		if col >= startCol && col <= endCol && row >= startRow && row <= endRow {
			ranges = []int{startCol, startRow, endCol, endRow}
		}
	}

	startCol, startRow, endCol, endRow := col, row, col, row
	if len(ranges) == 4 {
		startCol, startRow, endCol, endRow = ranges[0], ranges[1], ranges[2], ranges[3]
	}

	// 设置宽度
	colCount := endCol - startCol + 1
	if colCount > 1 {
		width = width / float64(colCount)
		if width < 0 {
			width = 0
		}
		for c := startCol; c <= endCol; c++ {
			name, err := excelize.ColumnNumberToName(c)
			if err != nil {
				return err
			}
			if err = tpl.SetColWidth(sheet, name, name, width); err != nil {
				return err
			}
		}
	} else {
		colName, err := excelize.ColumnNumberToName(col)
		if err != nil {
			return err
		}
		if err = tpl.SetColWidth(sheet, colName, colName, width); err != nil {
			return err
		}
	}

	// 设置高度
	rowCount := endRow - startRow + 1
	if rowCount > 1 {
		height = height / float64(rowCount)
		if height < 0 {
			height = 0
		}
		for r := startRow; r <= endRow; r++ {
			if err := tpl.SetRowHeight(sheet, r, height); err != nil {
				return err
			}
		}
	} else {
		if err := tpl.SetRowHeight(sheet, row, height); err != nil {
			return err
		}
	}
	return nil
}
