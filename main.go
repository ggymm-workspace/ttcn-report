package main

import (
	"bufio"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"

	"github.com/xuri/excelize/v2"
)

var (
	body  = map[string]string{}
	logs  = map[string][]string{}
	sheet = "Sheet1"
)

func init() {
	// 获取工作目录
	exe, err := os.Executable()
	if err != nil {
		panic(err)
	}
	dir := filepath.Dir(exe)
	base := filepath.Base(exe)
	if strings.HasPrefix(exe, os.TempDir()) ||
		strings.HasPrefix(base, "___") {
		_, filename, _, ok := runtime.Caller(0)
		if ok {
			dir = filepath.Dir(filename)
		}
	}

	// 设置工作目录
	err = os.Chdir(dir)
	if err != nil {
		panic(err)
	}
}

func main() {
	var in, out string
	flag.StringVar(&in, "i", "", "日志文件路径")
	flag.StringVar(&out, "o", "", "日志文件路径")
	flag.Parse()

	if len(in) == 0 {
		os.Exit(-1)
	}
	if len(out) == 0 {
		os.Exit(-1)
	}
	_ = os.Remove(out)

	// 打开文件
	src := open(in)
	dst := excelize.NewFile()
	defer func() {
		_ = src.Close()

		// 保存并关闭文件
		_ = dst.SaveAs(out)
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
		if len(text) <= 9 {
			continue
		}
		if strings.HasPrefix(text, `"[LOG_ID:`) {
			text = text[9:]

			// 获取 LOG_ID 和 JSON 字符串
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
			if i == 0 || text[l-1:] != `"` {
				continue
			}

			// 获取对应模板
			id := text[:i]
			tpl := caseTpl(id)

			// 获取日志内容
			str := text[i+1 : l-1]
			str = strings.Replace(str, `\`, "", -1) // 移除额外的转义字符

			if len(str) == 0 || len(tpl) == 0 {
				continue
			}
			logs[tpl] = append(logs[tpl], str)
		} else if strings.HasPrefix(text, `"[BODY_ID:`) {
			// 获取 BODY_ID 和 内容 字符串
			var (
				i = 0
				l = len(text)
			)
			for i < l {
				if text[i] == '"' {
					break
				}
				i++
			}

			// 获取 ID
			id := text[:i]

			// 获取 内容
			str := text[i+1:]

			if len(id) == 0 || len(str) == 0 {
				continue
			}
			body[id] = str
		}
	}
	err := scanner.Err()
	if err != nil {
		panic(err)
	}

	// 遍历统计结果
	// 生成 excel 文件
	for k, log := range logs {
		switch k {
		case "7.5.1":
			err = render751(log, dst)
		case "7.5.2":
			err = render752(log, dst)
		case "7.6.2":
			err = render762(log, dst)
		case "7.6.3":
			err = render763(log, dst)
		case "7.6.4":
			err = render764(log, dst)
		case "7.6.5.1":
			err = render7651(log, dst)
		case "7.6.5.2":
			err = render7652(log, dst)
		case "7.6.5.3":
			err = render7653(log, dst)
		case "7.6.5.4":
			err = render7654(log, dst)
		case "7.6.7":
			err = render767(log, dst)
		case "7.6.8":
			err = render768(log, dst) // TODO: 未完成
		case "7.6.9":
			err = render769(log, dst)
		case "7.6.10":
			err = render7610(log, dst)
		case "7.6.11":
			err = render7611(log, dst)
		case "7.6.12":
			err = render7612(log, dst)
		case "7.6.13":
			err = render7613(log, dst)
		case "7.6.14":
			err = render7614(log, dst)
		case "7.6.15":
			err = render7615(log, dst)
		case "7.6.16":
			err = render7616(log, dst) // TODO: 未完成
		case "7.6.17":
			err = render7617(log, dst)
		case "7.6.18":
			err = render7618(log, dst)
		case "7.6.19":
			err = render7619(log, dst)
		case "7.6.20":
			err = render7620(log, dst) // TODO: 未完成
		case "7.6.21":
			err = render7621(log, dst) // TODO: 未完成
		}
		if err != nil {
			fmt.Printf("[error] render %s error: %s\n", k, err)
		}
	}
}
