package main

import (
	"github.com/xuri/excelize/v2"
)

func render7619(log []string, dst *excelize.File) error {
	t := "7.6.19"

	// 模板
	tpl, err := openTpl(t + ".xlsx")
	if err != nil {
		return err
	}
	defer func() {
		_ = tpl.Close()
	}()

	params := make(map[string]string)
	for _, s := range log {
		// 解析
		v := make(map[string]string)
		unmarshal([]byte(s), &v)

		// 赋值
		params["参数_1"] = v["V_AT1_AFTER"]
		params["参数_2"] = v["V_TIME1_S"]
		params["参数_3"] = v["V_TIME2_S"]
		params["参数_4"] = v["V_EVSE_IMD_DISABLED"]
		params["参数_5"] = v["V_SCREEN_SHOT1"]
		params["参数_6"] = v["V_SCREEN_SHOT2"]
		params["参数_7"] = v["V_TIME3_S"]
		params["参数_8"] = v["V_SCREEN_SHOT3"]
		params["参数_9"] = v["V_TIME4_S"]
		params["参数_10"] = v["V_CP_TURNED_OFF"]
		params["参数_11"] = v["V_SCREEN_SHOT4"]
	}

	// 设置 cell 的值
	err = setCell(tpl, params)
	if err != nil {
		return err
	}

	// 复制 sheet 页面
	err = copySheet(tpl, dst, sheet, t)
	if err != nil {
		return err
	}
	return nil
}
