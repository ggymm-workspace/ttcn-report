package main

import (
	"github.com/xuri/excelize/v2"
)

func render7652(log []string, dst *excelize.File) error {
	t := "7.5.6.2"

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
		params["参数_1"] = v["V_I_EVSE"]
		params["参数_2"] = v["V_V_EVSE"]
		params["参数_3"] = v["V_P_EVSE_MAX"]
		params["参数_4"] = v["V_TIME1_MS"]
		params["参数_5"] = v["V_TIME2_MS"]
		params["参数_6_图片"] = v["V_SCREEN_SHOT"]
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
