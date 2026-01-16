package main

import (
	"testing"

	"github.com/bytedance/sonic"
)

func Test_CaseID(t *testing.T) {
	str := `{"CASE_ID":"TC_EVSE_DC_VTB_DIN_7_6_2_1","V_CP_OFF_MS":"","V_I_DROP_MS":"","V_V_DROP_MS":"","V_SCREEN_SHOT":""}`

	id, err := sonic.Get([]byte(str))
	if err != nil {
		t.Fatal(err)
	}
	i, err := id.Get("CASE_ID").String()
	if err != nil {
		t.Fatal(err)
	}
	t.Logf("id:%v", i)
}
