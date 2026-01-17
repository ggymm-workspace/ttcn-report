package main

import (
	"slices"

	"github.com/xuri/excelize/v2"
)

func render7611(log []string, dst *excelize.File) error {
	t := "7.6.11"

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
		"V_V_DEV_REL_1",
		"V_V_DEV_REL_2",
		"V_V_DEV_ABS_MEASURE",
		"V_V_RIP",
		"V_V_RIP_LIMIT",
	}
	cols := []int{5, 6, 7, 8, 9, 11, 12, 14, 16, 18, 19}

	for _, s := range log {
		id := caseId(s)

		v := make([]map[string]string, 0, 63)
		unmarshal([]byte(s), &v)

		startRow := 7
		endRow := 27
		switch id {
		case "TC_EVSE_DC_VTB_DIN_CVM_7_6_11_Part_I":
			startRow = 7
			endRow = 27
		case "TC_EVSE_DC_VTB_DIN_CVM_7_6_11_Part_II":
			startRow = 28
			endRow = 48
		case "TC_EVSE_DC_VTB_DIN_CVM_7_6_11_Part_III":
			startRow = 49
			endRow = 69

			slices.Reverse(v) // 反转列表
		}

		// 赋值
		row := startRow
		for i := 0; i < len(v) && row <= endRow; i++ {
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
