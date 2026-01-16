package main

import "github.com/xuri/excelize/v2"

func render769(log []string, dst *excelize.File) error {
	t := "7.6.9"

	// 打开 tpl 文件
	tpl, err := excelize.OpenFile(t + ".xlsx")
	if err != nil {
		return err
	}
	defer func() {
		_ = tpl.Close()
	}()

	keys := []string{
		"V_I_EV_TARGET",
		"V_V_EV_TARGET",
		"V_I_EVSE_MEASURE",
		"V_V_EVSE_MEASURE",
		"V_I_EVSE_SIDEB_DC",
		"V_V_EVSE_SIDEB_DC",
		"V_I_DEV_ABS",
		"V_I_DEV_ABS_LIMIT",
		"V_I_DEV_ABS_MEASURE",
		"V_I_DEV_ABS_MEASURE_LIMIT",
		"V_V_DEV_ABS_MEASURE",
		"V_V_DEV_ABS_MEASURE_LIMIT",
		"V_I_RIP_LOW",
		"V_I_RIP_LOW_LIMIT",
		"V_I_RIP_MID",
		"V_I_RIP_MID_LIMIT",
		"V_I_RIP_HIGH",
		"V_I_RIP_HIGH_LIMIT",
	}

	for _, s := range log {
		v := make([]map[string]string, 0, 450)
		unmarshal([]byte(s), &v)

		// 赋值
		startCol := 3
		endCol := 20
		startRow := 7
		endRow := 447
		row := startRow
		for i := 0; i < len(v) && row <= endRow; i++ {
			col := startCol
			for _, key := range keys {
				if col > endCol {
					break
				}
				cell, err := excelize.CoordinatesToCellName(col, row)
				if err == nil {
					_ = tpl.SetCellValue(sheet, cell, v[i][key])
				}
				col++
			}
			row++
		}
	}

	// 复制 sheet 页面
	err = copySheet(tpl, dst, sheet, t)
	if err != nil {
		return err
	}
	return nil
}
