package main

import (
	"github.com/xuri/excelize/v2"
)

func render7612(log []string, dst *excelize.File) error {
	t := "7.6.12"

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
	nums := []int{-451, -451, -451, -449}

	for _, s := range log {
		v := make([]map[string]string, 0, 441*4)
		unmarshal([]byte(s), &v)

		startRow := 9
		endRow := 449
		id := caseId(s)
		switch id {
		case "TC_EVSE_DC_VTB_DIN_CCM_7_6_12_1":
			nums[0] = -nums[0]
			startRow = 9
			endRow = 449
		case "TC_EVSE_DC_VTB_DIN_CCM_7_6_12_2":
			nums[1] = -nums[1]
			startRow = 460
			endRow = 900
		case "TC_EVSE_DC_VTB_DIN_CCM_7_6_12_3":
			nums[2] = -nums[2]
			startRow = 911
			endRow = 1351
		case "TC_EVSE_DC_VTB_DIN_CCM_7_6_12_4":
			nums[3] = -nums[3]
			startRow = 1362
			endRow = 1802
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

	// 精简 sheet 页面
	start := 2 // 有内容的作为第一行
	for _, n := range nums {
		if n > 0 {
			// 存在，则增加行
			start += n
		} else {
			// 不存在，则删除行
			for i := n; i < 0; i++ {
				// 删除同一行
				_ = tpl.RemoveRow(sheet, start)
			}
		}
	}

	// 复制 sheet 页面
	err = copySheet(tpl, dst, sheet, t)
	if err != nil {
		return err
	}
	return nil
}
