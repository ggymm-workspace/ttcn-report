package main

import (
	"github.com/xuri/excelize/v2"
)

func render7610(log []string, dst *excelize.File) error {
	t := "7.6.10"

	// 打开 tpl 文件
	tpl, err := excelize.OpenFile(t + ".xlsx")
	if err != nil {
		return err
	}
	defer func() {
		_ = tpl.Close()
	}()

	keys := []string{
		"V_I_EVSE_MEASURE_INITIAL",
		"V_V_EVSE_MEASURE_INITIAL",
		"V_I_EVSE_MEASURE",
		"V_V_EVSE_MEASURE",
		"V_TRANSITION_MAX",
		"V_TRANSITION_MAX_LIMIT",
		"V_I_EVSE_SIDEB_DC_STEP1",
		"V_V_EVSE_SIDEB_DC_STEP1",
		"V_I_DEV_ABS",
		"V_I_DEV_ABS_LIMIT",
		"V_I_DEV_ABS_MEASURE",
		"V_I_DEV_ABS_MEASURE_LIMIT",
		"V_V_DEV_ABS_MEASURE",
		"V_V_DEV_ABS_MEASURE_LIMIT",
		"V_I_EVSE_SIDEB_DC_STEP2",
		"V_V_EVSE_SIDEB_DC_STEP2",
	}
	cols := []int{3, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19}

	for _, s := range log {
		v := make([]map[string]string, 0, 42)
		unmarshal([]byte(s), &v)

		// 赋值
		row := 9
		for i := 0; i < len(v) && row <= 50; i++ {
			for j, key := range keys {
				if j >= len(cols) {
					break
				}
				cell, _ := excelize.CoordinatesToCellName(cols[j], row)

				// 设置 cell 的值
				_ = tpl.SetCellValue(sheet, cell, v[i][key])
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
