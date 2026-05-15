package main

import (
	"os"

	"github.com/ggymm/gopkg/conv"
)

func open(p string) *os.File {
	f, err := os.OpenFile(p, os.O_RDONLY, os.ModePerm)
	if err != nil {
		panic(err)
	}
	return f
}

func exists(p string) bool {
	_, err := os.Stat(p)
	return err == nil
}

func formatFloat(v any) float64 {
	switch v.(type) {
	case string:
		return conv.ParseFloat64(v.(string))
	case float32:
		return float64(v.(float32))
	case float64:
		return v.(float64)
	}
	return 0.0
}
