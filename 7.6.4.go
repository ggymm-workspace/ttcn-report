package main

import (
	"regexp"

	"github.com/ggymm/gopkg/conv"
	"github.com/xuri/excelize/v2"
)

var (
	regex764 = regexp.MustCompile(`TP(\d+)_Part([A-Za-z]+)`)
)

func render764(log []string, dst *excelize.File) error {
	t := "7.6.4"

	// 打开 tpl 文件
	tpl, err := excelize.OpenFile(t + ".xlsx")
	if err != nil {
		return err
	}
	defer func() {
		_ = tpl.Close()
	}()

	nums := []int{
		-1, -1, -1, -1, -1, -1, // A
		-1, -1, -1, -1, -1, -1, // B
	}
	params := make(map[string]string)
	for _, s := range log {
		id := caseId(s)

		// 通过正则获取
		matches := regex764.FindStringSubmatch(id)
		if len(matches) != 3 {
			continue
		}
		i := matches[1] // 1 - 6
		p := matches[2] // A / B

		idx := conv.ParseInt(i) - 1
		if p == "B" {
			idx += 6
		}
		nums[idx] = -nums[idx] // 对应序号图表有值

		v := make(map[string]string)
		unmarshal([]byte(s), &v)

		// 赋值
		params[i+"_"+p+"_参数_1"] = v["V_I_EV_TARGET"]
		params[i+"_"+p+"_参数_2"] = v["V_V_EV_TARGET"]
		params[i+"_"+p+"_参数_3"] = v["V_I_EVSE_MEASURE"]
		params[i+"_"+p+"_参数_4"] = v["V_V_EVSE_MEASURE"]
		params[i+"_"+p+"_参数_5"] = v["V_I_EVSE_SIDEB_DC"]
		params[i+"_"+p+"_参数_6"] = v["V_V_EVSE_SIDEB_DC"]
		params[i+"_"+p+"_参数_7"] = v["V_I_DEV_ABS"]
		params[i+"_"+p+"_参数_8"] = v["V_I_DEV_ABS_LIMIT"]
		params[i+"_"+p+"_参数_9"] = v["V_I_DEV_ABS_MEASURE"]
		params[i+"_"+p+"_参数_10"] = v["V_I_DEV_ABS_MEASURE_LIMIT"]
		params[i+"_"+p+"_参数_11"] = v["V_V_DEV_ABS_MEASURE"]
		params[i+"_"+p+"_参数_12"] = v["V_V_DEV_ABS_MEASURE_LIMIT"]
		params[i+"_"+p+"_参数_13"] = v["V_I_RIP_LOW"]
		params[i+"_"+p+"_参数_14"] = v["V_I_RIP_LOW_LIMIT"]
		params[i+"_"+p+"_参数_15"] = v["V_I_RIP_MID"]
		params[i+"_"+p+"_参数_16"] = v["V_I_RIP_MID_LIMIT"]
		params[i+"_"+p+"_参数_17"] = v["V_I_RIP_HIGH"]
		params[i+"_"+p+"_参数_18"] = v["V_I_RIP_HIGH_LIMIT"]
	}

	// 设置 cell 的值
	err = setCell(tpl, params)
	if err != nil {
		return err
	}

	// 精简 sheet 页面
	start := 7
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
