package main

import (
	"os"
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
