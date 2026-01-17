package main

import (
	"github.com/ggymm/gopkg/conv"
	"github.com/xuri/excelize/v2"
)

func render762(log []string, dst *excelize.File) error {
	t := "7.6.2"

	// 打开 tpl 文件
	tpl, err := excelize.OpenFile(t + ".xlsx")
	if err != nil {
		return err
	}
	defer func() {
		_ = tpl.Close()
	}()

	nums := []int{-6, -5, -7, -7, -7, -7}
	params := make(map[string]string)
	for _, s := range log {
		id := caseId(s)

		// 序号
		i := id[len(id)-1:]
		idx := conv.ParseInt(i) - 1
		nums[idx] = -nums[idx] // 对应序号图表有值

		switch id {
		case "TC_EVSE_DC_VTB_DIN_7_6_2_1", "TC_EVSE_DC_VTB_DIN_7_6_2_2":
			// 解析
			v := make(map[string]string)
			unmarshal([]byte(s), &v)

			// 赋值
			params[i+"_参数_1"] = v["V_CP_OFF_MS"]
			params[i+"_参数_2"] = v["V_I_DROP_MS"]
			params[i+"_参数_3"] = v["V_V_DROP_MS"]
			params[i+"_参数_4_图片"] = v["V_SCREEN_SHOT"]
		case "TC_EVSE_DC_VTB_DIN_7_6_2_3", "TC_EVSE_DC_VTB_DIN_7_6_2_4",
			"TC_EVSE_DC_VTB_DIN_7_6_2_5", "TC_EVSE_DC_VTB_DIN_7_6_2_6":
			// 解析
			v := make(map[string]string)
			unmarshal([]byte(s), &v)

			// 赋值
			params[i+"_参数_1"] = v["V_I_EVSE"]
			params[i+"_参数_2"] = v["V_V_EVSE"]
			params[i+"_参数_3"] = v["V_P_EVSE_MAX"]
			params[i+"_参数_4"] = v["V_CP_OFF_MS"]
			params[i+"_参数_5"] = v["V_I_DROP_MS"]
			params[i+"_参数_6"] = v["V_V_DROP_MS"]
			params[i+"_参数_7_图片"] = v["V_SCREEN_SHOT"]
		}
	}

	// 设置 cell 的值
	err = setCell(tpl, params)
	if err != nil {
		return err
	}

	// 精简 sheet 页面
	start := 4 // 有内容的作为第一行
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
