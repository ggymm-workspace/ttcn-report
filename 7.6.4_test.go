package main

import (
	"testing"
)

func Test_Regex764(t *testing.T) {
	str := "TC_EVSE_DC_VTB_DIN_7_6_4_TP6_PartA"

	t.Logf("%+v", regex764.FindStringSubmatch(str))
}
