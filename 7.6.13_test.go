package main

import (
	"testing"
)

func Test_Regex7613(t *testing.T) {
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

	for _, p := range parts {
		t.Logf("%+v", regex7613.FindStringSubmatch(p))
	}
}
