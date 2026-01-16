package main

import (
	"github.com/bytedance/sonic"
)

func marshal(val any) []byte {
	buf, _ := sonic.Marshal(val)
	return buf
}

func unmarshal(buf []byte, val any) {
	_ = sonic.Unmarshal(buf, val)
}
