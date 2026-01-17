package main

import (
	"regexp"
	"slices"

	"github.com/xuri/excelize/v2"
)

var (
	regex7613 = regexp.MustCompile(
		`(?:Symmetric|Asymmetric|Disturbance)_(\d+)(?:_(\d+)(?:st|nd|rd|th))?$`,
	)
)

func render7613(log []string, dst *excelize.File) error {
	t := "7.6.13"

	// 模板
	tpl, err := openTpl(t + ".xlsx")
	if err != nil {
		return err
	}
	defer func() {
		_ = tpl.Close()
	}()

	nums := []int{
		-18,                          // 1
		-19, -17, -17, -17, -17, -17, // 2
		-16,                          // 3
		-16,                          // 4
		-19, -17, -17, -17, -17, -17, // 5
		-19, -17, -17, -17, -17, -17, // 6
		-19, // 7
		-16, // 8
		-20, // 9
		-20, // 10
		-21, // 11
	}
	parts := []string{
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Symmetric_1",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Symmetric_2_1st",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Symmetric_2_2nd",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Symmetric_2_3rd",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Symmetric_2_4th",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Symmetric_2_5th",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Symmetric_2_6th",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Asymmetric_3",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Asymmetric_4",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Asymmetric_5_1st",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Asymmetric_5_2nd",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Asymmetric_5_3rd",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Asymmetric_5_4th",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Asymmetric_5_5th",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Asymmetric_5_6th",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Asymmetric_6_1st",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Asymmetric_6_2nd",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Asymmetric_6_3rd",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Asymmetric_6_4th",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Asymmetric_6_5th",
		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Asymmetric_6_6th",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Disturbance_7",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Disturbance_8",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Disturbance_9",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Disturbance_10",

		"TC_EVSE_DC_VTB_DIN_CCM_7_6_13_Disturbance_11",
	}
	params := make(map[string]string)
	for _, s := range log {
		id := caseId(s)

		i := ""
		matches := regex7613.FindStringSubmatch(id)
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
		params[i+"_参数_1"] = v["V_EVSE_VT8"]
		params[i+"_参数_2"] = v["V_VD"]
		params[i+"_参数_3"] = v["V_TIME1_MS"]
		params[i+"_参数_4"] = v["V_TIME2_MS"]
		params[i+"_参数_5"] = v["V_UCC"]
		params[i+"_参数_6_图片"] = v["V_SCREEN_SHOT"]
	}

	// 设置 cell 的值
	err = setCell(tpl, params)
	if err != nil {
		return err
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
