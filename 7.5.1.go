package main

import (
	"github.com/xuri/excelize/v2"
)

func render751(log []string, dst *excelize.File) error {
	t := "7.5.1"

	// 打开 tpl 文件
	tpl, err := excelize.OpenFile(t + ".xlsx")
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
		unmarshal([]byte(s), v)

		// 赋值
		params["参数_1"] = v["V_EVSE_MIN"]
		params["参数_2"] = v["V_EV_MAX_VOLT_LIMIT"]
		params["参数_3"] = v["V_IS_EVSE_STOP_NEXT_MSG"]
		params["参数_4"] = v["V_BODY"]
		params["参数_5"] = v["V_RCV_FAILED_WRONG_CHARGE_PARAMETER"]
		params["参数_6"] = v["V_T_8"]
		params["参数_7"] = v["V_LIMIT"]
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
