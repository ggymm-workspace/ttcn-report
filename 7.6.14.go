package main

import (
	"regexp"
	"slices"

	"github.com/xuri/excelize/v2"
)

var (
	regex7614 = regexp.MustCompile(
		`(?:Symmetric|Asymmetric|Disturbance)_(\d+)(?:_(\d+)(?:st|nd|rd|th))?$`,
	)
)

func render7614(log []string, dst *excelize.File) error {
	t := "7.6.14"

	// 模板
	tpl, err := openTpl(t + ".xlsx")
	if err != nil {
		return err
	}
	defer func() {
		_ = tpl.Close()
	}()

	nums := []int{
		-10,                          // 1
		-11, -10, -10, -10, -10, -10, // 2
		-9,                           // 3
		-9,                           // 4
		-11, -10, -10, -10, -10, -10, // 5
		-11, -10, -10, -10, -10, -10, // 6
		-11, // 7
		-9,  // 8
		-11, // 9
		-11, // 10
		-11, // 11
	}
	parts := []string{
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_1",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_2_1st",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_2_2nd",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_2_3rd",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_2_4th",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_2_5th",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_2_6th",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Asymmetric_3",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Asymmetric_4",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_5_1st",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_5_2nd",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_5_3rd",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_5_4th",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_5_5th",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_5_6th",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_6_1st",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_6_2nd",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_6_3rd",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_6_4th",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_6_5th",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Symmetric_6_6th",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Disturbance_7",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Disturbance_8",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Disturbance_9",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Disturbance_10",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_14_Disturbance_11",
	}
	params := make(map[string]string)
	for _, s := range log {
		id := caseId(s)

		i := ""
		matches := regex7614.FindStringSubmatch(id)
		if len(matches) < 2 {
			i = matches[1]
			if len(matches) >= 3 && matches[2] != "" {
				i = i + "_" + matches[2]
			}
		}

		// 序号
		idx := slices.Index(parts, id)
		if idx < 0 {
			continue
		}
		nums[idx] = -nums[idx] // 对应序号图表有值

		// 解析
		v := make(map[string]string)
		unmarshal([]byte(s), &v)

		// 赋值
		params[i+"_参数_1"] = v["V_I_EVSE"]
		params[i+"_参数_2"] = v["V_V_EVSE"]
		params[i+"_参数_3"] = v["V_P_EVSE_MAX"]
		params[i+"_参数_4"] = v["V_KEEP_TRANSMISSION"]
		params[i+"_参数_5"] = v["V_TIME1_MS"]
		params[i+"_参数_6"] = v["V_TIME2_MS"]
		params[i+"_参数_7"] = v["V_TIME3_MS"]
		params[i+"_参数_8"] = v["V_VS1"]
		params[i+"_参数_9_图片"] = v["V_SCREEN_SHOT"]
	}

	// 设置 cell 的值
	err = setCell(tpl, params)
	if err != nil {
		return err
	}

	// 精简 sheet 页面
	start := 3 // 有内容的作为第一行
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
