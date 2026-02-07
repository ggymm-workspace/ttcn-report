package main

import (
	"github.com/xuri/excelize/v2"
)

func render752(log []string, dst *excelize.File) error {
	t := "7.5.2"

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
		params["参数_1"] = v["V_PP_RESISTOR"]
		params["参数_2"] = v["V_CP_DUTY_CYCLE"]
		params["参数_3_图片"] = v["V_SCREEN_SHOT1"]
		params["参数_4"] = v["V_PREPARATION_PHASE_BODY"]
		params["参数_5"] = v["V_V_EVSE_CABLE_CHECK"]
		params["参数_6_图片"] = v["V_SCREEN_SHOT2"]
		params["参数_7"] = v["V_BDC_CABLE_CHECK_PHASE_BODY"]
		params["参数_8"] = v["V_AC_CABLE_CHECK_PHASE_BODY"]
		params["参数_9"] = v["V_PRESENT_CURRENT_SIDEB"]
		params["参数_10"] = v["V_PRESENT_VOLTAGE_SIDEB"]
		params["参数_11"] = v["V_TARGET_VOLTAGE"]
		params["参数_12_图片"] = v["V_SCREEN_SHOT3"]
		params["参数_13"] = v["V_PRECHARGE_PHASE_BODY"]
		params["参数_14"] = v["V_AFTER_PRECHARGE_PHASE_BODY"]
		params["参数_15"] = v["V_ENERGY_TRANSFER_PHASE_BODY"]
		params["参数_16_图片"] = v["V_SCREEN_SHOT4"]
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
