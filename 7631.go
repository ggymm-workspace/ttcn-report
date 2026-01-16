package main

import (
	"github.com/xuri/excelize/v2"
)

type case7631 struct {
	CASEID         string `json:"CASE_ID"`
	VIEVSEINTENDED string `json:"V_I_EVSE_INTENDED"`
	VVEVSEINTENDED string `json:"V_V_EVSE_INTENDED"`
	VPEVSEINTENDED string `json:"V_P_EVSE_INTENDED"`
	VCPOFFMS       string `json:"V_CP_OFF_MS"`
	VIDROPMS       string `json:"V_I_DROP_MS"`
	VVDROPMS       string `json:"V_V_DROP_MS"`
	VSCREENSHOT    string `json:"V_SCREEN_SHOT"`
}

func render7631(id string, val *case7631, dst *excelize.File) error {
	tpl, err := excelize.OpenFile("template/" + id + ".xlsx")
	if err != nil {
		return err
	}
	defer func() {
		_ = tpl.Close()
	}()

	// 设置 cell 的值
	err = setCell(tpl, map[string]string{
		"参数1":     val.VIEVSEINTENDED,
		"参数2":     val.VVEVSEINTENDED,
		"参数3":     val.VPEVSEINTENDED,
		"参数4":     val.VCPOFFMS,
		"参数5":     val.VIDROPMS,
		"参数6":     val.VVDROPMS,
		"参数7（图片）": val.VSCREENSHOT,
	})
	if err != nil {
		return err
	}

	// 新建 sheet 页面
	idx, err := dst.NewSheet(id)
	if err != nil {
		return err
	}
	dst.SetActiveSheet(idx)

	// 复制 sheet 页面
	err = copySheet(tpl, dst, sheet, id)
	if err != nil {
		return err
	}
	return nil
}
