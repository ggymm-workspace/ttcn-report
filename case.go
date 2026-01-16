package main

import (
	"strings"

	"github.com/bytedance/sonic"
)

func caseId(str string) string {
	root, err := sonic.Get([]byte(str))
	if err != nil {
		panic(err)
	}
	id, err := root.Get("CASE_ID").String()
	if err != nil {
		panic(err)
	}
	return id
}

func caseTpl(str string) string {
	if strings.Contains(str, "7_5_1") {
		return "7.5.1"
	} else if strings.Contains(str, "7_5_2") {
		return "7.5.2"
	} else if strings.Contains(str, "7_6_2_") {
		return "7.6.2"
	} else if strings.Contains(str, "7_6_3_") {
		return "7.6.3"
	} else if strings.Contains(str, "7_6_4_") {
		return "7.6.4"
	} else if strings.Contains(str, "7_6_5_1_") {
		return "7.6.5.1"
	} else if strings.Contains(str, "7_6_5_2") {
		return "7.6.5.2"
	} else if strings.Contains(str, "7_6_5_3") {
		return "7.6.5.3"
	} else if strings.Contains(str, "7_6_5_4") {
		return "7.6.5.4"
	} else if strings.Contains(str, "7_6_7_") {
		return "7.6.7"
	} else if strings.Contains(str, "7_6_9") {
		return "7.6.9"
	}
	return ""
}
