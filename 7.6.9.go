package main

import (
	"github.com/xuri/excelize/v2"
)

func render769(log []string, dst *excelize.File) error {
	t := "7.6.9"

	// 模板
	tpl, err := openTpl(t + ".xlsx")
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
		v := make([]map[string]string, 0, 441)
		unmarshal([]byte(s), &v)

		// 赋值
		row := 7
		for i := 0; i < len(v) && row <= 447; i++ {
			col := 3
			for _, key := range keys {
				if col > 20 {
					break
				}
				cell, _ := excelize.CoordinatesToCellName(col, row)

				// 设置 cell 的值
				_ = tpl.SetCellValue(sheet, cell, v[i][key])

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
