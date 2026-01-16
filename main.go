package main

import (
	"bufio"
	"fmt"
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

var (
	// file = "docs/case报告/示例日志/ATS_Test_2026-01-15_09-47-26-MTC.log"
	file = "docs/case报告/示例日志/ATS_Test_2026-01-15_09-50-12-MTC.log"
	// file = "docs/case报告/示例日志/ATS_Test_2026-01-15_09-58-31-MTC.log"
)

var (
	sheet = "Sheet1"
)

func init() {
	err := os.Chdir("/Volumes/Data/Code/workspace/ttcn-report")
	if err != nil {
		panic(err)
	}
}

func main() {
	src := open(file)
	dst := excelize.NewFile()
	defer func() {
		_ = src.Close()
		_ = dst.Close()
	}()

	// 按行解析
	scanner := bufio.NewScanner(src)
	for scanner.Scan() {
		text := scanner.Text()
		if len(text) <= 16 {
			continue
		}

		text = text[16:]
		if len(text) <= 9 || text[:9] != `"[LOG_ID:` {
			continue
		}
		text = text[9:]

		// 提取出 LOG_ID 用于确定模板 excel 内容
		var (
			i = 0
			l = len(text)
		)
		for i < l {
			if text[i] == ']' {
				break
			}
			i++
		}
		if i == -1 || text[l-1:] != `"` {
			continue
		}

		id := text[:i]
		str := text[i+1 : l-1]
		str = strings.Replace(str, `\`, "", -1) // 移除额外的转义字符

		if len(str) == 0 {
			continue
		}

		// 输出到文件
		var err error
		switch id {
		case "TC_EVSE_DC_VTB_DIN_7_6_2_1":
			val := &case7621{}
			unmarshal([]byte(str), val)

			err = render7621(id, val, dst) // 解析后输出到文件
		case "TC_EVSE_DC_VTB_DIN_7_6_3_TP1":
			val := &case7631{}
			unmarshal([]byte(str), val)

			err = render7631(id, val, dst) // 解析后输出到文件
		}
		if err != nil {
			fmt.Printf("[error] %s generate failed: %s\n", id, err)
			continue
		}
	}
	err := scanner.Err()
	if err != nil {
		panic(err)
	}

	// 保存文件
	_ = os.Remove("output/report.xlsx")
	err = dst.SaveAs("output/report.xlsx")
	if err != nil {
		panic(err)
	}
}
