\ Model START
\ LP format - for model browsing. Use MPS format to capture full model detail.
Minimize
  1581 xstand_0 + 953 xstand_1 + 696 xstand_2 + 592 xstand_3 + 592 xstand_4
   + 1541 xdisc_0 + 913 xdisc_1 + 656 xdisc_2 + 552 xdisc_3 + 552 xdisc_4
   - 40 xdisc_5 + 845 ystand_0 + 973 ystand_1 + 973 ystand_2
   + 1072 ystand_3 + 615 ystand_4 + 885 ystand_5 + 805 ydisc_0
   + 933 ydisc_1 + 933 ydisc_2 + 1032 ydisc_3 + 575 ydisc_4 + 845 ydisc_5
Subject To
 Profiles_0_0: beta_0_0_0 + beta_4_0_0 + beta_8_0_0 + beta_9_0_0 >= 2
 Profiles_0_1: beta_0_0_1 + beta_4_0_1 + beta_8_0_1 + beta_9_0_1 >= 2
 Profiles_0_2: beta_0_0_2 + beta_4_0_2 + beta_8_0_2 + beta_9_0_2 >= 2
 Profiles_0_3: beta_0_0_3 + beta_4_0_3 + beta_8_0_3 + beta_9_0_3 >= 2
 Profiles_0_4: beta_0_0_4 + beta_4_0_4 + beta_8_0_4 + beta_9_0_4 >= 2
 Profiles_1_0: beta_1_1_0 + beta_2_1_0 + beta_3_1_0 + beta_14_1_0 >= 1
 Profiles_1_1: beta_1_1_1 + beta_2_1_1 + beta_3_1_1 + beta_14_1_1 >= 1
 Profiles_1_2: beta_1_1_2 + beta_2_1_2 + beta_3_1_2 + beta_14_1_2 >= 1
 Profiles_1_3: beta_1_1_3 + beta_2_1_3 + beta_3_1_3 + beta_14_1_3 >= 1
 Profiles_1_4: beta_1_1_4 + beta_2_1_4 + beta_3_1_4 + beta_14_1_4 >= 1
 Profiles_2_0: beta_1_2_0 + beta_2_2_0 + beta_13_2_0 >= 1
 Profiles_2_1: beta_1_2_1 + beta_2_2_1 + beta_13_2_1 >= 1
 Profiles_2_2: beta_1_2_2 + beta_2_2_2 + beta_13_2_2 >= 1
 Profiles_2_3: beta_1_2_3 + beta_2_2_3 + beta_13_2_3 >= 1
 Profiles_2_4: beta_1_2_4 + beta_2_2_4 + beta_13_2_4 >= 1
 Profiles_3_0: beta_5_3_0 + beta_7_3_0 + beta_16_3_0 >= 1
 Profiles_3_1: beta_5_3_1 + beta_7_3_1 + beta_16_3_1 >= 1
 Profiles_3_2: beta_5_3_2 + beta_7_3_2 + beta_16_3_2 >= 1
 Profiles_3_3: beta_5_3_3 + beta_7_3_3 + beta_16_3_3 >= 1
 Profiles_3_4: beta_5_3_4 + beta_7_3_4 + beta_16_3_4 >= 1
 Profiles_4_0: beta_6_4_0 + beta_10_4_0 + beta_11_4_0 + beta_12_4_0
   + beta_17_4_0 >= 2
 Profiles_4_1: beta_6_4_1 + beta_10_4_1 + beta_11_4_1 + beta_12_4_1
   + beta_17_4_1 >= 2
 Profiles_4_2: beta_6_4_2 + beta_10_4_2 + beta_11_4_2 + beta_12_4_2
   + beta_17_4_2 >= 2
 Profiles_4_3: beta_6_4_3 + beta_10_4_3 + beta_11_4_3 + beta_12_4_3
   + beta_17_4_3 >= 2
 Profiles_4_4: beta_6_4_4 + beta_10_4_4 + beta_11_4_4 + beta_12_4_4
   + beta_17_4_4 >= 2
 Profiles_5_0: beta_1_5_0 + beta_2_5_0 + beta_5_5_0 + beta_7_5_0
   + beta_10_5_0 + beta_11_5_0 + beta_12_5_0 + beta_13_5_0 + beta_15_5_0
   + beta_16_5_0 >= 2
 Profiles_5_1: beta_1_5_1 + beta_2_5_1 + beta_5_5_1 + beta_7_5_1
   + beta_10_5_1 + beta_11_5_1 + beta_12_5_1 + beta_13_5_1 + beta_15_5_1
   + beta_16_5_1 >= 2
 Profiles_5_2: beta_1_5_2 + beta_2_5_2 + beta_5_5_2 + beta_7_5_2
   + beta_10_5_2 + beta_11_5_2 + beta_12_5_2 + beta_13_5_2 + beta_15_5_2
   + beta_16_5_2 >= 2
 Profiles_5_3: beta_1_5_3 + beta_2_5_3 + beta_5_5_3 + beta_7_5_3
   + beta_10_5_3 + beta_11_5_3 + beta_12_5_3 + beta_13_5_3 + beta_15_5_3
   + beta_16_5_3 >= 2
 Profiles_5_4: beta_1_5_4 + beta_2_5_4 + beta_5_5_4 + beta_7_5_4
   + beta_10_5_4 + beta_11_5_4 + beta_12_5_4 + beta_13_5_4 + beta_15_5_4
   + beta_16_5_4 >= 2
 Maxperiods_0: - alphaout_0_1 - 2 alphaout_0_2 - 3 alphaout_0_3
   - 4 alphaout_0_4 - 5 alphaout_0_5 + alpharet_0_1 + 2 alpharet_0_2
   + 3 alpharet_0_3 + 4 alpharet_0_4 + 5 alpharet_0_5 <= 3
 Maxperiods_1: - alphaout_1_1 - 2 alphaout_1_2 - 3 alphaout_1_3
   - 4 alphaout_1_4 - 5 alphaout_1_5 + alpharet_1_1 + 2 alpharet_1_2
   + 3 alpharet_1_3 + 4 alpharet_1_4 + 5 alpharet_1_5 <= 3
 Maxperiods_2: - alphaout_2_1 - 2 alphaout_2_2 - 3 alphaout_2_3
   - 4 alphaout_2_4 - 5 alphaout_2_5 + alpharet_2_1 + 2 alpharet_2_2
   + 3 alpharet_2_3 + 4 alpharet_2_4 + 5 alpharet_2_5 <= 3
 Maxperiods_3: - alphaout_3_1 - 2 alphaout_3_2 - 3 alphaout_3_3
   - 4 alphaout_3_4 - 5 alphaout_3_5 + alpharet_3_1 + 2 alpharet_3_2
   + 3 alpharet_3_3 + 4 alpharet_3_4 + 5 alpharet_3_5 <= 3
 Maxperiods_4: - alphaout_4_1 - 2 alphaout_4_2 - 3 alphaout_4_3
   - 4 alphaout_4_4 - 5 alphaout_4_5 + alpharet_4_1 + 2 alpharet_4_2
   + 3 alpharet_4_3 + 4 alpharet_4_4 + 5 alpharet_4_5 <= 3
 Maxperiods_5: - alphaout_5_1 - 2 alphaout_5_2 - 3 alphaout_5_3
   - 4 alphaout_5_4 - 5 alphaout_5_5 + alpharet_5_1 + 2 alpharet_5_2
   + 3 alpharet_5_3 + 4 alpharet_5_4 + 5 alpharet_5_5 <= 3
 Maxperiods_6: - alphaout_6_1 - 2 alphaout_6_2 - 3 alphaout_6_3
   - 4 alphaout_6_4 - 5 alphaout_6_5 + alpharet_6_1 + 2 alpharet_6_2
   + 3 alpharet_6_3 + 4 alpharet_6_4 + 5 alpharet_6_5 <= 3
 Maxperiods_7: - alphaout_7_1 - 2 alphaout_7_2 - 3 alphaout_7_3
   - 4 alphaout_7_4 - 5 alphaout_7_5 + alpharet_7_1 + 2 alpharet_7_2
   + 3 alpharet_7_3 + 4 alpharet_7_4 + 5 alpharet_7_5 <= 3
 Maxperiods_8: - alphaout_8_1 - 2 alphaout_8_2 - 3 alphaout_8_3
   - 4 alphaout_8_4 - 5 alphaout_8_5 + alpharet_8_1 + 2 alpharet_8_2
   + 3 alpharet_8_3 + 4 alpharet_8_4 + 5 alpharet_8_5 <= 3
 Maxperiods_9: - alphaout_9_1 - 2 alphaout_9_2 - 3 alphaout_9_3
   - 4 alphaout_9_4 - 5 alphaout_9_5 + alpharet_9_1 + 2 alpharet_9_2
   + 3 alpharet_9_3 + 4 alpharet_9_4 + 5 alpharet_9_5 <= 3
 Maxperiods_10: - alphaout_10_1 - 2 alphaout_10_2 - 3 alphaout_10_3
   - 4 alphaout_10_4 - 5 alphaout_10_5 + alpharet_10_1 + 2 alpharet_10_2
   + 3 alpharet_10_3 + 4 alpharet_10_4 + 5 alpharet_10_5 <= 3
 Maxperiods_11: - alphaout_11_1 - 2 alphaout_11_2 - 3 alphaout_11_3
   - 4 alphaout_11_4 - 5 alphaout_11_5 + alpharet_11_1 + 2 alpharet_11_2
   + 3 alpharet_11_3 + 4 alpharet_11_4 + 5 alpharet_11_5 <= 3
 Maxperiods_12: - alphaout_12_1 - 2 alphaout_12_2 - 3 alphaout_12_3
   - 4 alphaout_12_4 - 5 alphaout_12_5 + alpharet_12_1 + 2 alpharet_12_2
   + 3 alpharet_12_3 + 4 alpharet_12_4 + 5 alpharet_12_5 <= 3
 Maxperiods_13: - alphaout_13_1 - 2 alphaout_13_2 - 3 alphaout_13_3
   - 4 alphaout_13_4 - 5 alphaout_13_5 + alpharet_13_1 + 2 alpharet_13_2
   + 3 alpharet_13_3 + 4 alpharet_13_4 + 5 alpharet_13_5 <= 3
 Maxperiods_14: - alphaout_14_1 - 2 alphaout_14_2 - 3 alphaout_14_3
   - 4 alphaout_14_4 - 5 alphaout_14_5 + alpharet_14_1 + 2 alpharet_14_2
   + 3 alpharet_14_3 + 4 alpharet_14_4 + 5 alpharet_14_5 <= 3
 Maxperiods_15: - alphaout_15_1 - 2 alphaout_15_2 - 3 alphaout_15_3
   - 4 alphaout_15_4 - 5 alphaout_15_5 + alpharet_15_1 + 2 alpharet_15_2
   + 3 alpharet_15_3 + 4 alpharet_15_4 + 5 alpharet_15_5 <= 3
 Maxperiods_16: - alphaout_16_1 - 2 alphaout_16_2 - 3 alphaout_16_3
   - 4 alphaout_16_4 - 5 alphaout_16_5 + alpharet_16_1 + 2 alpharet_16_2
   + 3 alpharet_16_3 + 4 alpharet_16_4 + 5 alpharet_16_5 <= 3
 Maxperiods_17: - alphaout_17_1 - 2 alphaout_17_2 - 3 alphaout_17_3
   - 4 alphaout_17_4 - 5 alphaout_17_5 + alpharet_17_1 + 2 alpharet_17_2
   + 3 alpharet_17_3 + 4 alpharet_17_4 + 5 alpharet_17_5 <= 3
 Vuelos_0_0: - alphaout_0_0 + alpharet_0_0 = 0
 Vuelos_0_1: - alphaout_0_1 + alpharet_0_1 = 0
 Vuelos_0_2: - alphaout_0_2 + alpharet_0_2 = 0
 Vuelos_0_3: - alphaout_0_3 + alpharet_0_3 = 0
 Vuelos_0_4: - alphaout_0_4 + alpharet_0_4 = 0
 Vuelos_0_5: - alphaout_0_5 + alpharet_0_5 = 0
 Vuelos_1_0: - alphaout_1_0 + alpharet_1_0 = 0
 Vuelos_1_1: - alphaout_1_1 + alpharet_1_1 = 0
 Vuelos_1_2: - alphaout_1_2 + alpharet_1_2 = 0
 Vuelos_1_3: - alphaout_1_3 + alpharet_1_3 = 0
 Vuelos_1_4: - alphaout_1_4 + alpharet_1_4 = 0
 Vuelos_1_5: - alphaout_1_5 + alpharet_1_5 = 0
 Vuelos_2_0: - alphaout_2_0 + alpharet_2_0 = 0
 Vuelos_2_1: - alphaout_2_1 + alpharet_2_1 = 0
 Vuelos_2_2: - alphaout_2_2 + alpharet_2_2 = 0
 Vuelos_2_3: - alphaout_2_3 + alpharet_2_3 = 0
 Vuelos_2_4: - alphaout_2_4 + alpharet_2_4 = 0
 Vuelos_2_5: - alphaout_2_5 + alpharet_2_5 = 0
 Vuelos_3_0: - alphaout_3_0 + alpharet_3_0 = 0
 Vuelos_3_1: - alphaout_3_1 + alpharet_3_1 = 0
 Vuelos_3_2: - alphaout_3_2 + alpharet_3_2 = 0
 Vuelos_3_3: - alphaout_3_3 + alpharet_3_3 = 0
 Vuelos_3_4: - alphaout_3_4 + alpharet_3_4 = 0
 Vuelos_3_5: - alphaout_3_5 + alpharet_3_5 = 0
 Vuelos_4_0: - alphaout_4_0 + alpharet_4_0 = 0
 Vuelos_4_1: - alphaout_4_1 + alpharet_4_1 = 0
 Vuelos_4_2: - alphaout_4_2 + alpharet_4_2 = 0
 Vuelos_4_3: - alphaout_4_3 + alpharet_4_3 = 0
 Vuelos_4_4: - alphaout_4_4 + alpharet_4_4 = 0
 Vuelos_4_5: - alphaout_4_5 + alpharet_4_5 = 0
 Vuelos_5_0: - alphaout_5_0 + alpharet_5_0 = 0
 Vuelos_5_1: - alphaout_5_1 + alpharet_5_1 = 0
 Vuelos_5_2: - alphaout_5_2 + alpharet_5_2 = 0
 Vuelos_5_3: - alphaout_5_3 + alpharet_5_3 = 0
 Vuelos_5_4: - alphaout_5_4 + alpharet_5_4 = 0
 Vuelos_5_5: - alphaout_5_5 + alpharet_5_5 = 0
 Vuelos_6_0: - alphaout_6_0 + alpharet_6_0 = 0
 Vuelos_6_1: - alphaout_6_1 + alpharet_6_1 = 0
 Vuelos_6_2: - alphaout_6_2 + alpharet_6_2 = 0
 Vuelos_6_3: - alphaout_6_3 + alpharet_6_3 = 0
 Vuelos_6_4: - alphaout_6_4 + alpharet_6_4 = 0
 Vuelos_6_5: - alphaout_6_5 + alpharet_6_5 = 0
 Vuelos_7_0: - alphaout_7_0 + alpharet_7_0 = 0
 Vuelos_7_1: - alphaout_7_1 + alpharet_7_1 = 0
 Vuelos_7_2: - alphaout_7_2 + alpharet_7_2 = 0
 Vuelos_7_3: - alphaout_7_3 + alpharet_7_3 = 0
 Vuelos_7_4: - alphaout_7_4 + alpharet_7_4 = 0
 Vuelos_7_5: - alphaout_7_5 + alpharet_7_5 = 0
 Vuelos_8_0: - alphaout_8_0 + alpharet_8_0 = 0
 Vuelos_8_1: - alphaout_8_1 + alpharet_8_1 = 0
 Vuelos_8_2: - alphaout_8_2 + alpharet_8_2 = 0
 Vuelos_8_3: - alphaout_8_3 + alpharet_8_3 = 0
 Vuelos_8_4: - alphaout_8_4 + alpharet_8_4 = 0
 Vuelos_8_5: - alphaout_8_5 + alpharet_8_5 = 0
 Vuelos_9_0: - alphaout_9_0 + alpharet_9_0 = 0
 Vuelos_9_1: - alphaout_9_1 + alpharet_9_1 = 0
 Vuelos_9_2: - alphaout_9_2 + alpharet_9_2 = 0
 Vuelos_9_3: - alphaout_9_3 + alpharet_9_3 = 0
 Vuelos_9_4: - alphaout_9_4 + alpharet_9_4 = 0
 Vuelos_9_5: - alphaout_9_5 + alpharet_9_5 = 0
 Vuelos_10_0: - alphaout_10_0 + alpharet_10_0 = 0
 Vuelos_10_1: - alphaout_10_1 + alpharet_10_1 = 0
 Vuelos_10_2: - alphaout_10_2 + alpharet_10_2 = 0
 Vuelos_10_3: - alphaout_10_3 + alpharet_10_3 = 0
 Vuelos_10_4: - alphaout_10_4 + alpharet_10_4 = 0
 Vuelos_10_5: - alphaout_10_5 + alpharet_10_5 = 0
 Vuelos_11_0: - alphaout_11_0 + alpharet_11_0 = 0
 Vuelos_11_1: - alphaout_11_1 + alpharet_11_1 = 0
 Vuelos_11_2: - alphaout_11_2 + alpharet_11_2 = 0
 Vuelos_11_3: - alphaout_11_3 + alpharet_11_3 = 0
 Vuelos_11_4: - alphaout_11_4 + alpharet_11_4 = 0
 Vuelos_11_5: - alphaout_11_5 + alpharet_11_5 = 0
 Vuelos_12_0: - alphaout_12_0 + alpharet_12_0 = 0
 Vuelos_12_1: - alphaout_12_1 + alpharet_12_1 = 0
 Vuelos_12_2: - alphaout_12_2 + alpharet_12_2 = 0
 Vuelos_12_3: - alphaout_12_3 + alpharet_12_3 = 0
 Vuelos_12_4: - alphaout_12_4 + alpharet_12_4 = 0
 Vuelos_12_5: - alphaout_12_5 + alpharet_12_5 = 0
 Vuelos_13_0: - alphaout_13_0 + alpharet_13_0 = 0
 Vuelos_13_1: - alphaout_13_1 + alpharet_13_1 = 0
 Vuelos_13_2: - alphaout_13_2 + alpharet_13_2 = 0
 Vuelos_13_3: - alphaout_13_3 + alpharet_13_3 = 0
 Vuelos_13_4: - alphaout_13_4 + alpharet_13_4 = 0
 Vuelos_13_5: - alphaout_13_5 + alpharet_13_5 = 0
 Vuelos_14_0: - alphaout_14_0 + alpharet_14_0 = 0
 Vuelos_14_1: - alphaout_14_1 + alpharet_14_1 = 0
 Vuelos_14_2: - alphaout_14_2 + alpharet_14_2 = 0
 Vuelos_14_3: - alphaout_14_3 + alpharet_14_3 = 0
 Vuelos_14_4: - alphaout_14_4 + alpharet_14_4 = 0
 Vuelos_14_5: - alphaout_14_5 + alpharet_14_5 = 0
 Vuelos_15_0: - alphaout_15_0 + alpharet_15_0 = 0
 Vuelos_15_1: - alphaout_15_1 + alpharet_15_1 = 0
 Vuelos_15_2: - alphaout_15_2 + alpharet_15_2 = 0
 Vuelos_15_3: - alphaout_15_3 + alpharet_15_3 = 0
 Vuelos_15_4: - alphaout_15_4 + alpharet_15_4 = 0
 Vuelos_15_5: - alphaout_15_5 + alpharet_15_5 = 0
 Vuelos_16_0: - alphaout_16_0 + alpharet_16_0 = 0
 Vuelos_16_1: - alphaout_16_1 + alpharet_16_1 = 0
 Vuelos_16_2: - alphaout_16_2 + alpharet_16_2 = 0
 Vuelos_16_3: - alphaout_16_3 + alpharet_16_3 = 0
 Vuelos_16_4: - alphaout_16_4 + alpharet_16_4 = 0
 Vuelos_16_5: - alphaout_16_5 + alpharet_16_5 = 0
 Vuelos_17_0: - alphaout_17_0 + alpharet_17_0 = 0
 Vuelos_17_1: - alphaout_17_1 + alpharet_17_1 = 0
 Vuelos_17_2: - alphaout_17_2 + alpharet_17_2 = 0
 Vuelos_17_3: - alphaout_17_3 + alpharet_17_3 = 0
 Vuelos_17_4: - alphaout_17_4 + alpharet_17_4 = 0
 Vuelos_17_5: - alphaout_17_5 + alpharet_17_5 = 0
 Oneprofile_0_0: beta_0_0_0 + beta_0_1_0 + beta_0_2_0 + beta_0_3_0
   + beta_0_4_0 + beta_0_5_0 <= 1
 Oneprofile_0_1: beta_0_0_1 + beta_0_1_1 + beta_0_2_1 + beta_0_3_1
   + beta_0_4_1 + beta_0_5_1 <= 1
 Oneprofile_0_2: beta_0_0_2 + beta_0_1_2 + beta_0_2_2 + beta_0_3_2
   + beta_0_4_2 + beta_0_5_2 <= 1
 Oneprofile_0_3: beta_0_0_3 + beta_0_1_3 + beta_0_2_3 + beta_0_3_3
   + beta_0_4_3 + beta_0_5_3 <= 1
 Oneprofile_0_4: beta_0_0_4 + beta_0_1_4 + beta_0_2_4 + beta_0_3_4
   + beta_0_4_4 + beta_0_5_4 <= 1
 Oneprofile_0_5: beta_0_0_5 + beta_0_1_5 + beta_0_2_5 + beta_0_3_5
   + beta_0_4_5 + beta_0_5_5 <= 1
 Oneprofile_1_0: beta_1_0_0 + beta_1_1_0 + beta_1_2_0 + beta_1_3_0
   + beta_1_4_0 + beta_1_5_0 <= 1
 Oneprofile_1_1: beta_1_0_1 + beta_1_1_1 + beta_1_2_1 + beta_1_3_1
   + beta_1_4_1 + beta_1_5_1 <= 1
 Oneprofile_1_2: beta_1_0_2 + beta_1_1_2 + beta_1_2_2 + beta_1_3_2
   + beta_1_4_2 + beta_1_5_2 <= 1
 Oneprofile_1_3: beta_1_0_3 + beta_1_1_3 + beta_1_2_3 + beta_1_3_3
   + beta_1_4_3 + beta_1_5_3 <= 1
 Oneprofile_1_4: beta_1_0_4 + beta_1_1_4 + beta_1_2_4 + beta_1_3_4
   + beta_1_4_4 + beta_1_5_4 <= 1
 Oneprofile_1_5: beta_1_0_5 + beta_1_1_5 + beta_1_2_5 + beta_1_3_5
   + beta_1_4_5 + beta_1_5_5 <= 1
 Oneprofile_2_0: beta_2_0_0 + beta_2_1_0 + beta_2_2_0 + beta_2_3_0
   + beta_2_4_0 + beta_2_5_0 <= 1
 Oneprofile_2_1: beta_2_0_1 + beta_2_1_1 + beta_2_2_1 + beta_2_3_1
   + beta_2_4_1 + beta_2_5_1 <= 1
 Oneprofile_2_2: beta_2_0_2 + beta_2_1_2 + beta_2_2_2 + beta_2_3_2
   + beta_2_4_2 + beta_2_5_2 <= 1
 Oneprofile_2_3: beta_2_0_3 + beta_2_1_3 + beta_2_2_3 + beta_2_3_3
   + beta_2_4_3 + beta_2_5_3 <= 1
 Oneprofile_2_4: beta_2_0_4 + beta_2_1_4 + beta_2_2_4 + beta_2_3_4
   + beta_2_4_4 + beta_2_5_4 <= 1
 Oneprofile_2_5: beta_2_0_5 + beta_2_1_5 + beta_2_2_5 + beta_2_3_5
   + beta_2_4_5 + beta_2_5_5 <= 1
 Oneprofile_3_0: beta_3_0_0 + beta_3_1_0 + beta_3_2_0 + beta_3_3_0
   + beta_3_4_0 + beta_3_5_0 <= 1
 Oneprofile_3_1: beta_3_0_1 + beta_3_1_1 + beta_3_2_1 + beta_3_3_1
   + beta_3_4_1 + beta_3_5_1 <= 1
 Oneprofile_3_2: beta_3_0_2 + beta_3_1_2 + beta_3_2_2 + beta_3_3_2
   + beta_3_4_2 + beta_3_5_2 <= 1
 Oneprofile_3_3: beta_3_0_3 + beta_3_1_3 + beta_3_2_3 + beta_3_3_3
   + beta_3_4_3 + beta_3_5_3 <= 1
 Oneprofile_3_4: beta_3_0_4 + beta_3_1_4 + beta_3_2_4 + beta_3_3_4
   + beta_3_4_4 + beta_3_5_4 <= 1
 Oneprofile_3_5: beta_3_0_5 + beta_3_1_5 + beta_3_2_5 + beta_3_3_5
   + beta_3_4_5 + beta_3_5_5 <= 1
 Oneprofile_4_0: beta_4_0_0 + beta_4_1_0 + beta_4_2_0 + beta_4_3_0
   + beta_4_4_0 + beta_4_5_0 <= 1
 Oneprofile_4_1: beta_4_0_1 + beta_4_1_1 + beta_4_2_1 + beta_4_3_1
   + beta_4_4_1 + beta_4_5_1 <= 1
 Oneprofile_4_2: beta_4_0_2 + beta_4_1_2 + beta_4_2_2 + beta_4_3_2
   + beta_4_4_2 + beta_4_5_2 <= 1
 Oneprofile_4_3: beta_4_0_3 + beta_4_1_3 + beta_4_2_3 + beta_4_3_3
   + beta_4_4_3 + beta_4_5_3 <= 1
 Oneprofile_4_4: beta_4_0_4 + beta_4_1_4 + beta_4_2_4 + beta_4_3_4
   + beta_4_4_4 + beta_4_5_4 <= 1
 Oneprofile_4_5: beta_4_0_5 + beta_4_1_5 + beta_4_2_5 + beta_4_3_5
   + beta_4_4_5 + beta_4_5_5 <= 1
 Oneprofile_5_0: beta_5_0_0 + beta_5_1_0 + beta_5_2_0 + beta_5_3_0
   + beta_5_4_0 + beta_5_5_0 <= 1
 Oneprofile_5_1: beta_5_0_1 + beta_5_1_1 + beta_5_2_1 + beta_5_3_1
   + beta_5_4_1 + beta_5_5_1 <= 1
 Oneprofile_5_2: beta_5_0_2 + beta_5_1_2 + beta_5_2_2 + beta_5_3_2
   + beta_5_4_2 + beta_5_5_2 <= 1
 Oneprofile_5_3: beta_5_0_3 + beta_5_1_3 + beta_5_2_3 + beta_5_3_3
   + beta_5_4_3 + beta_5_5_3 <= 1
 Oneprofile_5_4: beta_5_0_4 + beta_5_1_4 + beta_5_2_4 + beta_5_3_4
   + beta_5_4_4 + beta_5_5_4 <= 1
 Oneprofile_5_5: beta_5_0_5 + beta_5_1_5 + beta_5_2_5 + beta_5_3_5
   + beta_5_4_5 + beta_5_5_5 <= 1
 Oneprofile_6_0: beta_6_0_0 + beta_6_1_0 + beta_6_2_0 + beta_6_3_0
   + beta_6_4_0 + beta_6_5_0 <= 1
 Oneprofile_6_1: beta_6_0_1 + beta_6_1_1 + beta_6_2_1 + beta_6_3_1
   + beta_6_4_1 + beta_6_5_1 <= 1
 Oneprofile_6_2: beta_6_0_2 + beta_6_1_2 + beta_6_2_2 + beta_6_3_2
   + beta_6_4_2 + beta_6_5_2 <= 1
 Oneprofile_6_3: beta_6_0_3 + beta_6_1_3 + beta_6_2_3 + beta_6_3_3
   + beta_6_4_3 + beta_6_5_3 <= 1
 Oneprofile_6_4: beta_6_0_4 + beta_6_1_4 + beta_6_2_4 + beta_6_3_4
   + beta_6_4_4 + beta_6_5_4 <= 1
 Oneprofile_6_5: beta_6_0_5 + beta_6_1_5 + beta_6_2_5 + beta_6_3_5
   + beta_6_4_5 + beta_6_5_5 <= 1
 Oneprofile_7_0: beta_7_0_0 + beta_7_1_0 + beta_7_2_0 + beta_7_3_0
   + beta_7_4_0 + beta_7_5_0 <= 1
 Oneprofile_7_1: beta_7_0_1 + beta_7_1_1 + beta_7_2_1 + beta_7_3_1
   + beta_7_4_1 + beta_7_5_1 <= 1
 Oneprofile_7_2: beta_7_0_2 + beta_7_1_2 + beta_7_2_2 + beta_7_3_2
   + beta_7_4_2 + beta_7_5_2 <= 1
 Oneprofile_7_3: beta_7_0_3 + beta_7_1_3 + beta_7_2_3 + beta_7_3_3
   + beta_7_4_3 + beta_7_5_3 <= 1
 Oneprofile_7_4: beta_7_0_4 + beta_7_1_4 + beta_7_2_4 + beta_7_3_4
   + beta_7_4_4 + beta_7_5_4 <= 1
 Oneprofile_7_5: beta_7_0_5 + beta_7_1_5 + beta_7_2_5 + beta_7_3_5
   + beta_7_4_5 + beta_7_5_5 <= 1
 Oneprofile_8_0: beta_8_0_0 + beta_8_1_0 + beta_8_2_0 + beta_8_3_0
   + beta_8_4_0 + beta_8_5_0 <= 1
 Oneprofile_8_1: beta_8_0_1 + beta_8_1_1 + beta_8_2_1 + beta_8_3_1
   + beta_8_4_1 + beta_8_5_1 <= 1
 Oneprofile_8_2: beta_8_0_2 + beta_8_1_2 + beta_8_2_2 + beta_8_3_2
   + beta_8_4_2 + beta_8_5_2 <= 1
 Oneprofile_8_3: beta_8_0_3 + beta_8_1_3 + beta_8_2_3 + beta_8_3_3
   + beta_8_4_3 + beta_8_5_3 <= 1
 Oneprofile_8_4: beta_8_0_4 + beta_8_1_4 + beta_8_2_4 + beta_8_3_4
   + beta_8_4_4 + beta_8_5_4 <= 1
 Oneprofile_8_5: beta_8_0_5 + beta_8_1_5 + beta_8_2_5 + beta_8_3_5
   + beta_8_4_5 + beta_8_5_5 <= 1
 Oneprofile_9_0: beta_9_0_0 + beta_9_1_0 + beta_9_2_0 + beta_9_3_0
   + beta_9_4_0 + beta_9_5_0 <= 1
 Oneprofile_9_1: beta_9_0_1 + beta_9_1_1 + beta_9_2_1 + beta_9_3_1
   + beta_9_4_1 + beta_9_5_1 <= 1
 Oneprofile_9_2: beta_9_0_2 + beta_9_1_2 + beta_9_2_2 + beta_9_3_2
   + beta_9_4_2 + beta_9_5_2 <= 1
 Oneprofile_9_3: beta_9_0_3 + beta_9_1_3 + beta_9_2_3 + beta_9_3_3
   + beta_9_4_3 + beta_9_5_3 <= 1
 Oneprofile_9_4: beta_9_0_4 + beta_9_1_4 + beta_9_2_4 + beta_9_3_4
   + beta_9_4_4 + beta_9_5_4 <= 1
 Oneprofile_9_5: beta_9_0_5 + beta_9_1_5 + beta_9_2_5 + beta_9_3_5
   + beta_9_4_5 + beta_9_5_5 <= 1
 Oneprofile_10_0: beta_10_0_0 + beta_10_1_0 + beta_10_2_0 + beta_10_3_0
   + beta_10_4_0 + beta_10_5_0 <= 1
 Oneprofile_10_1: beta_10_0_1 + beta_10_1_1 + beta_10_2_1 + beta_10_3_1
   + beta_10_4_1 + beta_10_5_1 <= 1
 Oneprofile_10_2: beta_10_0_2 + beta_10_1_2 + beta_10_2_2 + beta_10_3_2
   + beta_10_4_2 + beta_10_5_2 <= 1
 Oneprofile_10_3: beta_10_0_3 + beta_10_1_3 + beta_10_2_3 + beta_10_3_3
   + beta_10_4_3 + beta_10_5_3 <= 1
 Oneprofile_10_4: beta_10_0_4 + beta_10_1_4 + beta_10_2_4 + beta_10_3_4
   + beta_10_4_4 + beta_10_5_4 <= 1
 Oneprofile_10_5: beta_10_0_5 + beta_10_1_5 + beta_10_2_5 + beta_10_3_5
   + beta_10_4_5 + beta_10_5_5 <= 1
 Oneprofile_11_0: beta_11_0_0 + beta_11_1_0 + beta_11_2_0 + beta_11_3_0
   + beta_11_4_0 + beta_11_5_0 <= 1
 Oneprofile_11_1: beta_11_0_1 + beta_11_1_1 + beta_11_2_1 + beta_11_3_1
   + beta_11_4_1 + beta_11_5_1 <= 1
 Oneprofile_11_2: beta_11_0_2 + beta_11_1_2 + beta_11_2_2 + beta_11_3_2
   + beta_11_4_2 + beta_11_5_2 <= 1
 Oneprofile_11_3: beta_11_0_3 + beta_11_1_3 + beta_11_2_3 + beta_11_3_3
   + beta_11_4_3 + beta_11_5_3 <= 1
 Oneprofile_11_4: beta_11_0_4 + beta_11_1_4 + beta_11_2_4 + beta_11_3_4
   + beta_11_4_4 + beta_11_5_4 <= 1
 Oneprofile_11_5: beta_11_0_5 + beta_11_1_5 + beta_11_2_5 + beta_11_3_5
   + beta_11_4_5 + beta_11_5_5 <= 1
 Oneprofile_12_0: beta_12_0_0 + beta_12_1_0 + beta_12_2_0 + beta_12_3_0
   + beta_12_4_0 + beta_12_5_0 <= 1
 Oneprofile_12_1: beta_12_0_1 + beta_12_1_1 + beta_12_2_1 + beta_12_3_1
   + beta_12_4_1 + beta_12_5_1 <= 1
 Oneprofile_12_2: beta_12_0_2 + beta_12_1_2 + beta_12_2_2 + beta_12_3_2
   + beta_12_4_2 + beta_12_5_2 <= 1
 Oneprofile_12_3: beta_12_0_3 + beta_12_1_3 + beta_12_2_3 + beta_12_3_3
   + beta_12_4_3 + beta_12_5_3 <= 1
 Oneprofile_12_4: beta_12_0_4 + beta_12_1_4 + beta_12_2_4 + beta_12_3_4
   + beta_12_4_4 + beta_12_5_4 <= 1
 Oneprofile_12_5: beta_12_0_5 + beta_12_1_5 + beta_12_2_5 + beta_12_3_5
   + beta_12_4_5 + beta_12_5_5 <= 1
 Oneprofile_13_0: beta_13_0_0 + beta_13_1_0 + beta_13_2_0 + beta_13_3_0
   + beta_13_4_0 + beta_13_5_0 <= 1
 Oneprofile_13_1: beta_13_0_1 + beta_13_1_1 + beta_13_2_1 + beta_13_3_1
   + beta_13_4_1 + beta_13_5_1 <= 1
 Oneprofile_13_2: beta_13_0_2 + beta_13_1_2 + beta_13_2_2 + beta_13_3_2
   + beta_13_4_2 + beta_13_5_2 <= 1
 Oneprofile_13_3: beta_13_0_3 + beta_13_1_3 + beta_13_2_3 + beta_13_3_3
   + beta_13_4_3 + beta_13_5_3 <= 1
 Oneprofile_13_4: beta_13_0_4 + beta_13_1_4 + beta_13_2_4 + beta_13_3_4
   + beta_13_4_4 + beta_13_5_4 <= 1
 Oneprofile_13_5: beta_13_0_5 + beta_13_1_5 + beta_13_2_5 + beta_13_3_5
   + beta_13_4_5 + beta_13_5_5 <= 1
 Oneprofile_14_0: beta_14_0_0 + beta_14_1_0 + beta_14_2_0 + beta_14_3_0
   + beta_14_4_0 + beta_14_5_0 <= 1
 Oneprofile_14_1: beta_14_0_1 + beta_14_1_1 + beta_14_2_1 + beta_14_3_1
   + beta_14_4_1 + beta_14_5_1 <= 1
 Oneprofile_14_2: beta_14_0_2 + beta_14_1_2 + beta_14_2_2 + beta_14_3_2
   + beta_14_4_2 + beta_14_5_2 <= 1
 Oneprofile_14_3: beta_14_0_3 + beta_14_1_3 + beta_14_2_3 + beta_14_3_3
   + beta_14_4_3 + beta_14_5_3 <= 1
 Oneprofile_14_4: beta_14_0_4 + beta_14_1_4 + beta_14_2_4 + beta_14_3_4
   + beta_14_4_4 + beta_14_5_4 <= 1
 Oneprofile_14_5: beta_14_0_5 + beta_14_1_5 + beta_14_2_5 + beta_14_3_5
   + beta_14_4_5 + beta_14_5_5 <= 1
 Oneprofile_15_0: beta_15_0_0 + beta_15_1_0 + beta_15_2_0 + beta_15_3_0
   + beta_15_4_0 + beta_15_5_0 <= 1
 Oneprofile_15_1: beta_15_0_1 + beta_15_1_1 + beta_15_2_1 + beta_15_3_1
   + beta_15_4_1 + beta_15_5_1 <= 1
 Oneprofile_15_2: beta_15_0_2 + beta_15_1_2 + beta_15_2_2 + beta_15_3_2
   + beta_15_4_2 + beta_15_5_2 <= 1
 Oneprofile_15_3: beta_15_0_3 + beta_15_1_3 + beta_15_2_3 + beta_15_3_3
   + beta_15_4_3 + beta_15_5_3 <= 1
 Oneprofile_15_4: beta_15_0_4 + beta_15_1_4 + beta_15_2_4 + beta_15_3_4
   + beta_15_4_4 + beta_15_5_4 <= 1
 Oneprofile_15_5: beta_15_0_5 + beta_15_1_5 + beta_15_2_5 + beta_15_3_5
   + beta_15_4_5 + beta_15_5_5 <= 1
 Oneprofile_16_0: beta_16_0_0 + beta_16_1_0 + beta_16_2_0 + beta_16_3_0
   + beta_16_4_0 + beta_16_5_0 <= 1
 Oneprofile_16_1: beta_16_0_1 + beta_16_1_1 + beta_16_2_1 + beta_16_3_1
   + beta_16_4_1 + beta_16_5_1 <= 1
 Oneprofile_16_2: beta_16_0_2 + beta_16_1_2 + beta_16_2_2 + beta_16_3_2
   + beta_16_4_2 + beta_16_5_2 <= 1
 Oneprofile_16_3: beta_16_0_3 + beta_16_1_3 + beta_16_2_3 + beta_16_3_3
   + beta_16_4_3 + beta_16_5_3 <= 1
 Oneprofile_16_4: beta_16_0_4 + beta_16_1_4 + beta_16_2_4 + beta_16_3_4
   + beta_16_4_4 + beta_16_5_4 <= 1
 Oneprofile_16_5: beta_16_0_5 + beta_16_1_5 + beta_16_2_5 + beta_16_3_5
   + beta_16_4_5 + beta_16_5_5 <= 1
 Oneprofile_17_0: beta_17_0_0 + beta_17_1_0 + beta_17_2_0 + beta_17_3_0
   + beta_17_4_0 + beta_17_5_0 <= 1
 Oneprofile_17_1: beta_17_0_1 + beta_17_1_1 + beta_17_2_1 + beta_17_3_1
   + beta_17_4_1 + beta_17_5_1 <= 1
 Oneprofile_17_2: beta_17_0_2 + beta_17_1_2 + beta_17_2_2 + beta_17_3_2
   + beta_17_4_2 + beta_17_5_2 <= 1
 Oneprofile_17_3: beta_17_0_3 + beta_17_1_3 + beta_17_2_3 + beta_17_3_3
   + beta_17_4_3 + beta_17_5_3 <= 1
 Oneprofile_17_4: beta_17_0_4 + beta_17_1_4 + beta_17_2_4 + beta_17_3_4
   + beta_17_4_4 + beta_17_5_4 <= 1
 Oneprofile_17_5: beta_17_0_5 + beta_17_1_5 + beta_17_2_5 + beta_17_3_5
   + beta_17_4_5 + beta_17_5_5 <= 1
 Onceout_0: alphaout_0_0 + alphaout_0_1 + alphaout_0_2 + alphaout_0_3
   + alphaout_0_4 + alphaout_0_5 <= 1
 Onceout_1: alphaout_1_0 + alphaout_1_1 + alphaout_1_2 + alphaout_1_3
   + alphaout_1_4 + alphaout_1_5 <= 1
 Onceout_2: alphaout_2_0 + alphaout_2_1 + alphaout_2_2 + alphaout_2_3
   + alphaout_2_4 + alphaout_2_5 <= 1
 Onceout_3: alphaout_3_0 + alphaout_3_1 + alphaout_3_2 + alphaout_3_3
   + alphaout_3_4 + alphaout_3_5 <= 1
 Onceout_4: alphaout_4_0 + alphaout_4_1 + alphaout_4_2 + alphaout_4_3
   + alphaout_4_4 + alphaout_4_5 <= 1
 Onceout_5: alphaout_5_0 + alphaout_5_1 + alphaout_5_2 + alphaout_5_3
   + alphaout_5_4 + alphaout_5_5 <= 1
 Onceout_6: alphaout_6_0 + alphaout_6_1 + alphaout_6_2 + alphaout_6_3
   + alphaout_6_4 + alphaout_6_5 <= 1
 Onceout_7: alphaout_7_0 + alphaout_7_1 + alphaout_7_2 + alphaout_7_3
   + alphaout_7_4 + alphaout_7_5 <= 1
 Onceout_8: alphaout_8_0 + alphaout_8_1 + alphaout_8_2 + alphaout_8_3
   + alphaout_8_4 + alphaout_8_5 <= 1
 Onceout_9: alphaout_9_0 + alphaout_9_1 + alphaout_9_2 + alphaout_9_3
   + alphaout_9_4 + alphaout_9_5 <= 1
 Onceout_10: alphaout_10_0 + alphaout_10_1 + alphaout_10_2 + alphaout_10_3
   + alphaout_10_4 + alphaout_10_5 <= 1
 Onceout_11: alphaout_11_0 + alphaout_11_1 + alphaout_11_2 + alphaout_11_3
   + alphaout_11_4 + alphaout_11_5 <= 1
 Onceout_12: alphaout_12_0 + alphaout_12_1 + alphaout_12_2 + alphaout_12_3
   + alphaout_12_4 + alphaout_12_5 <= 1
 Onceout_13: alphaout_13_0 + alphaout_13_1 + alphaout_13_2 + alphaout_13_3
   + alphaout_13_4 + alphaout_13_5 <= 1
 Onceout_14: alphaout_14_0 + alphaout_14_1 + alphaout_14_2 + alphaout_14_3
   + alphaout_14_4 + alphaout_14_5 <= 1
 Onceout_15: alphaout_15_0 + alphaout_15_1 + alphaout_15_2 + alphaout_15_3
   + alphaout_15_4 + alphaout_15_5 <= 1
 Onceout_16: alphaout_16_0 + alphaout_16_1 + alphaout_16_2 + alphaout_16_3
   + alphaout_16_4 + alphaout_16_5 <= 1
 Onceout_17: alphaout_17_0 + alphaout_17_1 + alphaout_17_2 + alphaout_17_3
   + alphaout_17_4 + alphaout_17_5 <= 1
 Onceret_0: alpharet_0_0 + alpharet_0_1 + alpharet_0_2 + alpharet_0_3
   + alpharet_0_4 + alpharet_0_5 <= 1
 Onceret_1: alpharet_1_0 + alpharet_1_1 + alpharet_1_2 + alpharet_1_3
   + alpharet_1_4 + alpharet_1_5 <= 1
 Onceret_2: alpharet_2_0 + alpharet_2_1 + alpharet_2_2 + alpharet_2_3
   + alpharet_2_4 + alpharet_2_5 <= 1
 Onceret_3: alpharet_3_0 + alpharet_3_1 + alpharet_3_2 + alpharet_3_3
   + alpharet_3_4 + alpharet_3_5 <= 1
 Onceret_4: alpharet_4_0 + alpharet_4_1 + alpharet_4_2 + alpharet_4_3
   + alpharet_4_4 + alpharet_4_5 <= 1
 Onceret_5: alpharet_5_0 + alpharet_5_1 + alpharet_5_2 + alpharet_5_3
   + alpharet_5_4 + alpharet_5_5 <= 1
 Onceret_6: alpharet_6_0 + alpharet_6_1 + alpharet_6_2 + alpharet_6_3
   + alpharet_6_4 + alpharet_6_5 <= 1
 Onceret_7: alpharet_7_0 + alpharet_7_1 + alpharet_7_2 + alpharet_7_3
   + alpharet_7_4 + alpharet_7_5 <= 1
 Onceret_8: alpharet_8_0 + alpharet_8_1 + alpharet_8_2 + alpharet_8_3
   + alpharet_8_4 + alpharet_8_5 <= 1
 Onceret_9: alpharet_9_0 + alpharet_9_1 + alpharet_9_2 + alpharet_9_3
   + alpharet_9_4 + alpharet_9_5 <= 1
 Onceret_10: alpharet_10_0 + alpharet_10_1 + alpharet_10_2 + alpharet_10_3
   + alpharet_10_4 + alpharet_10_5 <= 1
 Onceret_11: alpharet_11_0 + alpharet_11_1 + alpharet_11_2 + alpharet_11_3
   + alpharet_11_4 + alpharet_11_5 <= 1
 Onceret_12: alpharet_12_0 + alpharet_12_1 + alpharet_12_2 + alpharet_12_3
   + alpharet_12_4 + alpharet_12_5 <= 1
 Onceret_13: alpharet_13_0 + alpharet_13_1 + alpharet_13_2 + alpharet_13_3
   + alpharet_13_4 + alpharet_13_5 <= 1
 Onceret_14: alpharet_14_0 + alpharet_14_1 + alpharet_14_2 + alpharet_14_3
   + alpharet_14_4 + alpharet_14_5 <= 1
 Onceret_15: alpharet_15_0 + alpharet_15_1 + alpharet_15_2 + alpharet_15_3
   + alpharet_15_4 + alpharet_15_5 <= 1
 Onceret_16: alpharet_16_0 + alpharet_16_1 + alpharet_16_2 + alpharet_16_3
   + alpharet_16_4 + alpharet_16_5 <= 1
 Onceret_17: alpharet_17_0 + alpharet_17_1 + alpharet_17_2 + alpharet_17_3
   + alpharet_17_4 + alpharet_17_5 <= 1
 Countout_0: - alphaout_0_0 - alphaout_1_0 - alphaout_2_0 - alphaout_3_0
   - alphaout_4_0 - alphaout_5_0 - alphaout_6_0 - alphaout_7_0
   - alphaout_8_0 - alphaout_9_0 - alphaout_10_0 - alphaout_11_0
   - alphaout_12_0 - alphaout_13_0 - alphaout_14_0 - alphaout_15_0
   - alphaout_16_0 - alphaout_17_0 + nout_0 = 0
 Countout_1: - alphaout_0_1 - alphaout_1_1 - alphaout_2_1 - alphaout_3_1
   - alphaout_4_1 - alphaout_5_1 - alphaout_6_1 - alphaout_7_1
   - alphaout_8_1 - alphaout_9_1 - alphaout_10_1 - alphaout_11_1
   - alphaout_12_1 - alphaout_13_1 - alphaout_14_1 - alphaout_15_1
   - alphaout_16_1 - alphaout_17_1 + nout_1 = 0
 Countout_2: - alphaout_0_2 - alphaout_1_2 - alphaout_2_2 - alphaout_3_2
   - alphaout_4_2 - alphaout_5_2 - alphaout_6_2 - alphaout_7_2
   - alphaout_8_2 - alphaout_9_2 - alphaout_10_2 - alphaout_11_2
   - alphaout_12_2 - alphaout_13_2 - alphaout_14_2 - alphaout_15_2
   - alphaout_16_2 - alphaout_17_2 + nout_2 = 0
 Countout_3: - alphaout_0_3 - alphaout_1_3 - alphaout_2_3 - alphaout_3_3
   - alphaout_4_3 - alphaout_5_3 - alphaout_6_3 - alphaout_7_3
   - alphaout_8_3 - alphaout_9_3 - alphaout_10_3 - alphaout_11_3
   - alphaout_12_3 - alphaout_13_3 - alphaout_14_3 - alphaout_15_3
   - alphaout_16_3 - alphaout_17_3 + nout_3 = 0
 Countout_4: - alphaout_0_4 - alphaout_1_4 - alphaout_2_4 - alphaout_3_4
   - alphaout_4_4 - alphaout_5_4 - alphaout_6_4 - alphaout_7_4
   - alphaout_8_4 - alphaout_9_4 - alphaout_10_4 - alphaout_11_4
   - alphaout_12_4 - alphaout_13_4 - alphaout_14_4 - alphaout_15_4
   - alphaout_16_4 - alphaout_17_4 + nout_4 = 0
 Countout_5: - alphaout_0_5 - alphaout_1_5 - alphaout_2_5 - alphaout_3_5
   - alphaout_4_5 - alphaout_5_5 - alphaout_6_5 - alphaout_7_5
   - alphaout_8_5 - alphaout_9_5 - alphaout_10_5 - alphaout_11_5
   - alphaout_12_5 - alphaout_13_5 - alphaout_14_5 - alphaout_15_5
   - alphaout_16_5 - alphaout_17_5 + nout_5 = 0
 Countret_0: - alpharet_0_0 - alpharet_1_0 - alpharet_2_0 - alpharet_3_0
   - alpharet_4_0 - alpharet_5_0 - alpharet_6_0 - alpharet_7_0
   - alpharet_8_0 - alpharet_9_0 - alpharet_10_0 - alpharet_11_0
   - alpharet_12_0 - alpharet_13_0 - alpharet_14_0 - alpharet_15_0
   - alpharet_16_0 - alpharet_17_0 + nret_0 = 0
 Countret_1: - alpharet_0_1 - alpharet_1_1 - alpharet_2_1 - alpharet_3_1
   - alpharet_4_1 - alpharet_5_1 - alpharet_6_1 - alpharet_7_1
   - alpharet_8_1 - alpharet_9_1 - alpharet_10_1 - alpharet_11_1
   - alpharet_12_1 - alpharet_13_1 - alpharet_14_1 - alpharet_15_1
   - alpharet_16_1 - alpharet_17_1 + nret_1 = 0
 Countret_2: - alpharet_0_2 - alpharet_1_2 - alpharet_2_2 - alpharet_3_2
   - alpharet_4_2 - alpharet_5_2 - alpharet_6_2 - alpharet_7_2
   - alpharet_8_2 - alpharet_9_2 - alpharet_10_2 - alpharet_11_2
   - alpharet_12_2 - alpharet_13_2 - alpharet_14_2 - alpharet_15_2
   - alpharet_16_2 - alpharet_17_2 + nret_2 = 0
 Countret_3: - alpharet_0_3 - alpharet_1_3 - alpharet_2_3 - alpharet_3_3
   - alpharet_4_3 - alpharet_5_3 - alpharet_6_3 - alpharet_7_3
   - alpharet_8_3 - alpharet_9_3 - alpharet_10_3 - alpharet_11_3
   - alpharet_12_3 - alpharet_13_3 - alpharet_14_3 - alpharet_15_3
   - alpharet_16_3 - alpharet_17_3 + nret_3 = 0
 Countret_4: - alpharet_0_4 - alpharet_1_4 - alpharet_2_4 - alpharet_3_4
   - alpharet_4_4 - alpharet_5_4 - alpharet_6_4 - alpharet_7_4
   - alpharet_8_4 - alpharet_9_4 - alpharet_10_4 - alpharet_11_4
   - alpharet_12_4 - alpharet_13_4 - alpharet_14_4 - alpharet_15_4
   - alpharet_16_4 - alpharet_17_4 + nret_4 = 0
 Countret_5: - alpharet_0_5 - alpharet_1_5 - alpharet_2_5 - alpharet_3_5
   - alpharet_4_5 - alpharet_5_5 - alpharet_6_5 - alpharet_7_5
   - alpharet_8_5 - alpharet_9_5 - alpharet_10_5 - alpharet_11_5
   - alpharet_12_5 - alpharet_13_5 - alpharet_14_5 - alpharet_15_5
   - alpharet_16_5 - alpharet_17_5 + nret_5 = 0
 StandDiscout_0: nout_0 - xstand_0 - xdisc_0 = 0
 StandDiscout_1: nout_1 - xstand_1 - xdisc_1 = 0
 StandDiscout_2: nout_2 - xstand_2 - xdisc_2 = 0
 StandDiscout_3: nout_3 - xstand_3 - xdisc_3 = 0
 StandDiscout_4: nout_4 - xstand_4 - xdisc_4 = 0
 StandDiscout_5: nout_5 - xstand_5 - xdisc_5 = 0
 StandDiscret_0: nret_0 - ystand_0 - ydisc_0 = 0
 StandDiscret_1: nret_1 - ystand_1 - ydisc_1 = 0
 StandDiscret_2: nret_2 - ystand_2 - ydisc_2 = 0
 StandDiscret_3: nret_3 - ystand_3 - ydisc_3 = 0
 StandDiscret_4: nret_4 - ystand_4 - ydisc_4 = 0
 StandDiscret_5: nret_5 - ystand_5 - ydisc_5 = 0
 StandOut_0: - 4 deltaout_0 + xstand_0 <= 0
 StandOut_1: - 4 deltaout_1 + xstand_1 <= 0
 StandOut_2: - 4 deltaout_2 + xstand_2 <= 0
 StandOut_3: - 4 deltaout_3 + xstand_3 <= 0
 StandOut_4: - 4 deltaout_4 + xstand_4 <= 0
 StandOut_5: - 4 deltaout_5 + xstand_5 <= 0
 DiscOut1_0: - 5 deltaout_0 - xdisc_0 <= -5
 DiscOut1_1: - 5 deltaout_1 - xdisc_1 <= -5
 DiscOut1_2: - 5 deltaout_2 - xdisc_2 <= -5
 DiscOut1_3: - 5 deltaout_3 - xdisc_3 <= -5
 DiscOut1_4: - 5 deltaout_4 - xdisc_4 <= -5
 DiscOut1_5: - 5 deltaout_5 - xdisc_5 <= -5
 DiscOut2_0: 18 deltaout_0 + xdisc_0 <= 18
 DiscOut2_1: 18 deltaout_1 + xdisc_1 <= 18
 DiscOut2_2: 18 deltaout_2 + xdisc_2 <= 18
 DiscOut2_3: 18 deltaout_3 + xdisc_3 <= 18
 DiscOut2_4: 18 deltaout_4 + xdisc_4 <= 18
 DiscOut2_5: 18 deltaout_5 + xdisc_5 <= 18
 StandRet_0: - 4 deltaret_0 + ystand_0 <= 0
 StandRet_1: - 4 deltaret_1 + ystand_1 <= 0
 StandRet_2: - 4 deltaret_2 + ystand_2 <= 0
 StandRet_3: - 4 deltaret_3 + ystand_3 <= 0
 StandRet_4: - 4 deltaret_4 + ystand_4 <= 0
 StandRet_5: - 4 deltaret_5 + ystand_5 <= 0
 DiscRet1_0: - 5 deltaret_0 - ydisc_0 <= -5
 DiscRet1_1: - 5 deltaret_1 - ydisc_1 <= -5
 DiscRet1_2: - 5 deltaret_2 - ydisc_2 <= -5
 DiscRet1_3: - 5 deltaret_3 - ydisc_3 <= -5
 DiscRet1_4: - 5 deltaret_4 - ydisc_4 <= -5
 DiscRet1_5: - 5 deltaret_5 - ydisc_5 <= -5
 DiscRet2_0: 18 deltaret_0 + ydisc_0 <= 18
 DiscRet2_1: 18 deltaret_1 + ydisc_1 <= 18
 DiscRet2_2: 18 deltaret_2 + ydisc_2 <= 18
 DiscRet2_3: 18 deltaret_3 + ydisc_3 <= 18
 DiscRet2_4: 18 deltaret_4 + ydisc_4 <= 18
 DiscRet2_5: 18 deltaret_5 + ydisc_5 <= 18
Bounds
 beta_0_5_0 = 0
 beta_0_5_1 = 0
 beta_0_5_2 = 0
 beta_0_5_3 = 0
 beta_0_5_4 = 0
 beta_0_5_5 = 0
 beta_1_5_0 = 0
 beta_1_5_1 = 0
 beta_1_5_2 = 0
 beta_1_5_3 = 0
 beta_1_5_4 = 0
 beta_1_5_5 = 0
 beta_2_5_0 = 0
 beta_2_5_1 = 0
 beta_2_5_2 = 0
 beta_2_5_3 = 0
 beta_2_5_4 = 0
 beta_2_5_5 = 0
 beta_3_5_0 = 0
 beta_3_5_1 = 0
 beta_3_5_2 = 0
 beta_3_5_3 = 0
 beta_3_5_4 = 0
 beta_3_5_5 = 0
 beta_4_5_0 = 0
 beta_4_5_1 = 0
 beta_4_5_2 = 0
 beta_4_5_3 = 0
 beta_4_5_4 = 0
 beta_4_5_5 = 0
 beta_5_5_0 = 0
 beta_5_5_1 = 0
 beta_5_5_2 = 0
 beta_5_5_3 = 0
 beta_5_5_4 = 0
 beta_5_5_5 = 0
 beta_6_5_0 = 0
 beta_6_5_1 = 0
 beta_6_5_2 = 0
 beta_6_5_3 = 0
 beta_6_5_4 = 0
 beta_6_5_5 = 0
 beta_7_5_0 = 0
 beta_7_5_1 = 0
 beta_7_5_2 = 0
 beta_7_5_3 = 0
 beta_7_5_4 = 0
 beta_7_5_5 = 0
 beta_8_5_0 = 0
 beta_8_5_1 = 0
 beta_8_5_2 = 0
 beta_8_5_3 = 0
 beta_8_5_4 = 0
 beta_8_5_5 = 0
 beta_9_5_0 = 0
 beta_9_5_1 = 0
 beta_9_5_2 = 0
 beta_9_5_3 = 0
 beta_9_5_4 = 0
 beta_9_5_5 = 0
 beta_10_5_0 = 0
 beta_10_5_1 = 0
 beta_10_5_2 = 0
 beta_10_5_3 = 0
 beta_10_5_4 = 0
 beta_10_5_5 = 0
 beta_11_5_0 = 0
 beta_11_5_1 = 0
 beta_11_5_2 = 0
 beta_11_5_3 = 0
 beta_11_5_4 = 0
 beta_11_5_5 = 0
 beta_12_5_0 = 0
 beta_12_5_1 = 0
 beta_12_5_2 = 0
 beta_12_5_3 = 0
 beta_12_5_4 = 0
 beta_12_5_5 = 0
 beta_13_5_0 = 0
 beta_13_5_1 = 0
 beta_13_5_2 = 0
 beta_13_5_3 = 0
 beta_13_5_4 = 0
 beta_13_5_5 = 0
 beta_14_5_0 = 0
 beta_14_5_1 = 0
 beta_14_5_2 = 0
 beta_14_5_3 = 0
 beta_14_5_4 = 0
 beta_14_5_5 = 0
 beta_15_5_0 = 0
 beta_15_5_1 = 0
 beta_15_5_2 = 0
 beta_15_5_3 = 0
 beta_15_5_4 = 0
 beta_15_5_5 = 0
 beta_16_5_0 = 0
 beta_16_5_1 = 0
 beta_16_5_2 = 0
 beta_16_5_3 = 0
 beta_16_5_4 = 0
 beta_16_5_5 = 0
 beta_17_5_0 = 0
 beta_17_5_1 = 0
 beta_17_5_2 = 0
 beta_17_5_3 = 0
 beta_17_5_4 = 0
 beta_17_5_5 = 0
Binaries
 beta_0_0_0 beta_0_0_1 beta_0_0_2 beta_0_0_3 beta_0_0_4 beta_0_0_5
 beta_0_1_0 beta_0_1_1 beta_0_1_2 beta_0_1_3 beta_0_1_4 beta_0_1_5
 beta_0_2_0 beta_0_2_1 beta_0_2_2 beta_0_2_3 beta_0_2_4 beta_0_2_5
 beta_0_3_0 beta_0_3_1 beta_0_3_2 beta_0_3_3 beta_0_3_4 beta_0_3_5
 beta_0_4_0 beta_0_4_1 beta_0_4_2 beta_0_4_3 beta_0_4_4 beta_0_4_5
 beta_0_5_0 beta_0_5_1 beta_0_5_2 beta_0_5_3 beta_0_5_4 beta_0_5_5
 beta_1_0_0 beta_1_0_1 beta_1_0_2 beta_1_0_3 beta_1_0_4 beta_1_0_5
 beta_1_1_0 beta_1_1_1 beta_1_1_2 beta_1_1_3 beta_1_1_4 beta_1_1_5
 beta_1_2_0 beta_1_2_1 beta_1_2_2 beta_1_2_3 beta_1_2_4 beta_1_2_5
 beta_1_3_0 beta_1_3_1 beta_1_3_2 beta_1_3_3 beta_1_3_4 beta_1_3_5
 beta_1_4_0 beta_1_4_1 beta_1_4_2 beta_1_4_3 beta_1_4_4 beta_1_4_5
 beta_1_5_0 beta_1_5_1 beta_1_5_2 beta_1_5_3 beta_1_5_4 beta_1_5_5
 beta_2_0_0 beta_2_0_1 beta_2_0_2 beta_2_0_3 beta_2_0_4 beta_2_0_5
 beta_2_1_0 beta_2_1_1 beta_2_1_2 beta_2_1_3 beta_2_1_4 beta_2_1_5
 beta_2_2_0 beta_2_2_1 beta_2_2_2 beta_2_2_3 beta_2_2_4 beta_2_2_5
 beta_2_3_0 beta_2_3_1 beta_2_3_2 beta_2_3_3 beta_2_3_4 beta_2_3_5
 beta_2_4_0 beta_2_4_1 beta_2_4_2 beta_2_4_3 beta_2_4_4 beta_2_4_5
 beta_2_5_0 beta_2_5_1 beta_2_5_2 beta_2_5_3 beta_2_5_4 beta_2_5_5
 beta_3_0_0 beta_3_0_1 beta_3_0_2 beta_3_0_3 beta_3_0_4 beta_3_0_5
 beta_3_1_0 beta_3_1_1 beta_3_1_2 beta_3_1_3 beta_3_1_4 beta_3_1_5
 beta_3_2_0 beta_3_2_1 beta_3_2_2 beta_3_2_3 beta_3_2_4 beta_3_2_5
 beta_3_3_0 beta_3_3_1 beta_3_3_2 beta_3_3_3 beta_3_3_4 beta_3_3_5
 beta_3_4_0 beta_3_4_1 beta_3_4_2 beta_3_4_3 beta_3_4_4 beta_3_4_5
 beta_3_5_0 beta_3_5_1 beta_3_5_2 beta_3_5_3 beta_3_5_4 beta_3_5_5
 beta_4_0_0 beta_4_0_1 beta_4_0_2 beta_4_0_3 beta_4_0_4 beta_4_0_5
 beta_4_1_0 beta_4_1_1 beta_4_1_2 beta_4_1_3 beta_4_1_4 beta_4_1_5
 beta_4_2_0 beta_4_2_1 beta_4_2_2 beta_4_2_3 beta_4_2_4 beta_4_2_5
 beta_4_3_0 beta_4_3_1 beta_4_3_2 beta_4_3_3 beta_4_3_4 beta_4_3_5
 beta_4_4_0 beta_4_4_1 beta_4_4_2 beta_4_4_3 beta_4_4_4 beta_4_4_5
 beta_4_5_0 beta_4_5_1 beta_4_5_2 beta_4_5_3 beta_4_5_4 beta_4_5_5
 beta_5_0_0 beta_5_0_1 beta_5_0_2 beta_5_0_3 beta_5_0_4 beta_5_0_5
 beta_5_1_0 beta_5_1_1 beta_5_1_2 beta_5_1_3 beta_5_1_4 beta_5_1_5
 beta_5_2_0 beta_5_2_1 beta_5_2_2 beta_5_2_3 beta_5_2_4 beta_5_2_5
 beta_5_3_0 beta_5_3_1 beta_5_3_2 beta_5_3_3 beta_5_3_4 beta_5_3_5
 beta_5_4_0 beta_5_4_1 beta_5_4_2 beta_5_4_3 beta_5_4_4 beta_5_4_5
 beta_5_5_0 beta_5_5_1 beta_5_5_2 beta_5_5_3 beta_5_5_4 beta_5_5_5
 beta_6_0_0 beta_6_0_1 beta_6_0_2 beta_6_0_3 beta_6_0_4 beta_6_0_5
 beta_6_1_0 beta_6_1_1 beta_6_1_2 beta_6_1_3 beta_6_1_4 beta_6_1_5
 beta_6_2_0 beta_6_2_1 beta_6_2_2 beta_6_2_3 beta_6_2_4 beta_6_2_5
 beta_6_3_0 beta_6_3_1 beta_6_3_2 beta_6_3_3 beta_6_3_4 beta_6_3_5
 beta_6_4_0 beta_6_4_1 beta_6_4_2 beta_6_4_3 beta_6_4_4 beta_6_4_5
 beta_6_5_0 beta_6_5_1 beta_6_5_2 beta_6_5_3 beta_6_5_4 beta_6_5_5
 beta_7_0_0 beta_7_0_1 beta_7_0_2 beta_7_0_3 beta_7_0_4 beta_7_0_5
 beta_7_1_0 beta_7_1_1 beta_7_1_2 beta_7_1_3 beta_7_1_4 beta_7_1_5
 beta_7_2_0 beta_7_2_1 beta_7_2_2 beta_7_2_3 beta_7_2_4 beta_7_2_5
 beta_7_3_0 beta_7_3_1 beta_7_3_2 beta_7_3_3 beta_7_3_4 beta_7_3_5
 beta_7_4_0 beta_7_4_1 beta_7_4_2 beta_7_4_3 beta_7_4_4 beta_7_4_5
 beta_7_5_0 beta_7_5_1 beta_7_5_2 beta_7_5_3 beta_7_5_4 beta_7_5_5
 beta_8_0_0 beta_8_0_1 beta_8_0_2 beta_8_0_3 beta_8_0_4 beta_8_0_5
 beta_8_1_0 beta_8_1_1 beta_8_1_2 beta_8_1_3 beta_8_1_4 beta_8_1_5
 beta_8_2_0 beta_8_2_1 beta_8_2_2 beta_8_2_3 beta_8_2_4 beta_8_2_5
 beta_8_3_0 beta_8_3_1 beta_8_3_2 beta_8_3_3 beta_8_3_4 beta_8_3_5
 beta_8_4_0 beta_8_4_1 beta_8_4_2 beta_8_4_3 beta_8_4_4 beta_8_4_5
 beta_8_5_0 beta_8_5_1 beta_8_5_2 beta_8_5_3 beta_8_5_4 beta_8_5_5
 beta_9_0_0 beta_9_0_1 beta_9_0_2 beta_9_0_3 beta_9_0_4 beta_9_0_5
 beta_9_1_0 beta_9_1_1 beta_9_1_2 beta_9_1_3 beta_9_1_4 beta_9_1_5
 beta_9_2_0 beta_9_2_1 beta_9_2_2 beta_9_2_3 beta_9_2_4 beta_9_2_5
 beta_9_3_0 beta_9_3_1 beta_9_3_2 beta_9_3_3 beta_9_3_4 beta_9_3_5
 beta_9_4_0 beta_9_4_1 beta_9_4_2 beta_9_4_3 beta_9_4_4 beta_9_4_5
 beta_9_5_0 beta_9_5_1 beta_9_5_2 beta_9_5_3 beta_9_5_4 beta_9_5_5
 beta_10_0_0 beta_10_0_1 beta_10_0_2 beta_10_0_3 beta_10_0_4 beta_10_0_5
 beta_10_1_0 beta_10_1_1 beta_10_1_2 beta_10_1_3 beta_10_1_4 beta_10_1_5
 beta_10_2_0 beta_10_2_1 beta_10_2_2 beta_10_2_3 beta_10_2_4 beta_10_2_5
 beta_10_3_0 beta_10_3_1 beta_10_3_2 beta_10_3_3 beta_10_3_4 beta_10_3_5
 beta_10_4_0 beta_10_4_1 beta_10_4_2 beta_10_4_3 beta_10_4_4 beta_10_4_5
 beta_10_5_0 beta_10_5_1 beta_10_5_2 beta_10_5_3 beta_10_5_4 beta_10_5_5
 beta_11_0_0 beta_11_0_1 beta_11_0_2 beta_11_0_3 beta_11_0_4 beta_11_0_5
 beta_11_1_0 beta_11_1_1 beta_11_1_2 beta_11_1_3 beta_11_1_4 beta_11_1_5
 beta_11_2_0 beta_11_2_1 beta_11_2_2 beta_11_2_3 beta_11_2_4 beta_11_2_5
 beta_11_3_0 beta_11_3_1 beta_11_3_2 beta_11_3_3 beta_11_3_4 beta_11_3_5
 beta_11_4_0 beta_11_4_1 beta_11_4_2 beta_11_4_3 beta_11_4_4 beta_11_4_5
 beta_11_5_0 beta_11_5_1 beta_11_5_2 beta_11_5_3 beta_11_5_4 beta_11_5_5
 beta_12_0_0 beta_12_0_1 beta_12_0_2 beta_12_0_3 beta_12_0_4 beta_12_0_5
 beta_12_1_0 beta_12_1_1 beta_12_1_2 beta_12_1_3 beta_12_1_4 beta_12_1_5
 beta_12_2_0 beta_12_2_1 beta_12_2_2 beta_12_2_3 beta_12_2_4 beta_12_2_5
 beta_12_3_0 beta_12_3_1 beta_12_3_2 beta_12_3_3 beta_12_3_4 beta_12_3_5
 beta_12_4_0 beta_12_4_1 beta_12_4_2 beta_12_4_3 beta_12_4_4 beta_12_4_5
 beta_12_5_0 beta_12_5_1 beta_12_5_2 beta_12_5_3 beta_12_5_4 beta_12_5_5
 beta_13_0_0 beta_13_0_1 beta_13_0_2 beta_13_0_3 beta_13_0_4 beta_13_0_5
 beta_13_1_0 beta_13_1_1 beta_13_1_2 beta_13_1_3 beta_13_1_4 beta_13_1_5
 beta_13_2_0 beta_13_2_1 beta_13_2_2 beta_13_2_3 beta_13_2_4 beta_13_2_5
 beta_13_3_0 beta_13_3_1 beta_13_3_2 beta_13_3_3 beta_13_3_4 beta_13_3_5
 beta_13_4_0 beta_13_4_1 beta_13_4_2 beta_13_4_3 beta_13_4_4 beta_13_4_5
 beta_13_5_0 beta_13_5_1 beta_13_5_2 beta_13_5_3 beta_13_5_4 beta_13_5_5
 beta_14_0_0 beta_14_0_1 beta_14_0_2 beta_14_0_3 beta_14_0_4 beta_14_0_5
 beta_14_1_0 beta_14_1_1 beta_14_1_2 beta_14_1_3 beta_14_1_4 beta_14_1_5
 beta_14_2_0 beta_14_2_1 beta_14_2_2 beta_14_2_3 beta_14_2_4 beta_14_2_5
 beta_14_3_0 beta_14_3_1 beta_14_3_2 beta_14_3_3 beta_14_3_4 beta_14_3_5
 beta_14_4_0 beta_14_4_1 beta_14_4_2 beta_14_4_3 beta_14_4_4 beta_14_4_5
 beta_14_5_0 beta_14_5_1 beta_14_5_2 beta_14_5_3 beta_14_5_4 beta_14_5_5
 beta_15_0_0 beta_15_0_1 beta_15_0_2 beta_15_0_3 beta_15_0_4 beta_15_0_5
 beta_15_1_0 beta_15_1_1 beta_15_1_2 beta_15_1_3 beta_15_1_4 beta_15_1_5
 beta_15_2_0 beta_15_2_1 beta_15_2_2 beta_15_2_3 beta_15_2_4 beta_15_2_5
 beta_15_3_0 beta_15_3_1 beta_15_3_2 beta_15_3_3 beta_15_3_4 beta_15_3_5
 beta_15_4_0 beta_15_4_1 beta_15_4_2 beta_15_4_3 beta_15_4_4 beta_15_4_5
 beta_15_5_0 beta_15_5_1 beta_15_5_2 beta_15_5_3 beta_15_5_4 beta_15_5_5
 beta_16_0_0 beta_16_0_1 beta_16_0_2 beta_16_0_3 beta_16_0_4 beta_16_0_5
 beta_16_1_0 beta_16_1_1 beta_16_1_2 beta_16_1_3 beta_16_1_4 beta_16_1_5
 beta_16_2_0 beta_16_2_1 beta_16_2_2 beta_16_2_3 beta_16_2_4 beta_16_2_5
 beta_16_3_0 beta_16_3_1 beta_16_3_2 beta_16_3_3 beta_16_3_4 beta_16_3_5
 beta_16_4_0 beta_16_4_1 beta_16_4_2 beta_16_4_3 beta_16_4_4 beta_16_4_5
 beta_16_5_0 beta_16_5_1 beta_16_5_2 beta_16_5_3 beta_16_5_4 beta_16_5_5
 beta_17_0_0 beta_17_0_1 beta_17_0_2 beta_17_0_3 beta_17_0_4 beta_17_0_5
 beta_17_1_0 beta_17_1_1 beta_17_1_2 beta_17_1_3 beta_17_1_4 beta_17_1_5
 beta_17_2_0 beta_17_2_1 beta_17_2_2 beta_17_2_3 beta_17_2_4 beta_17_2_5
 beta_17_3_0 beta_17_3_1 beta_17_3_2 beta_17_3_3 beta_17_3_4 beta_17_3_5
 beta_17_4_0 beta_17_4_1 beta_17_4_2 beta_17_4_3 beta_17_4_4 beta_17_4_5
 beta_17_5_0 beta_17_5_1 beta_17_5_2 beta_17_5_3 beta_17_5_4 beta_17_5_5
 alphaout_0_0 alphaout_0_1 alphaout_0_2 alphaout_0_3 alphaout_0_4
 alphaout_0_5 alphaout_1_0 alphaout_1_1 alphaout_1_2 alphaout_1_3
 alphaout_1_4 alphaout_1_5 alphaout_2_0 alphaout_2_1 alphaout_2_2
 alphaout_2_3 alphaout_2_4 alphaout_2_5 alphaout_3_0 alphaout_3_1
 alphaout_3_2 alphaout_3_3 alphaout_3_4 alphaout_3_5 alphaout_4_0
 alphaout_4_1 alphaout_4_2 alphaout_4_3 alphaout_4_4 alphaout_4_5
 alphaout_5_0 alphaout_5_1 alphaout_5_2 alphaout_5_3 alphaout_5_4
 alphaout_5_5 alphaout_6_0 alphaout_6_1 alphaout_6_2 alphaout_6_3
 alphaout_6_4 alphaout_6_5 alphaout_7_0 alphaout_7_1 alphaout_7_2
 alphaout_7_3 alphaout_7_4 alphaout_7_5 alphaout_8_0 alphaout_8_1
 alphaout_8_2 alphaout_8_3 alphaout_8_4 alphaout_8_5 alphaout_9_0
 alphaout_9_1 alphaout_9_2 alphaout_9_3 alphaout_9_4 alphaout_9_5
 alphaout_10_0 alphaout_10_1 alphaout_10_2 alphaout_10_3 alphaout_10_4
 alphaout_10_5 alphaout_11_0 alphaout_11_1 alphaout_11_2 alphaout_11_3
 alphaout_11_4 alphaout_11_5 alphaout_12_0 alphaout_12_1 alphaout_12_2
 alphaout_12_3 alphaout_12_4 alphaout_12_5 alphaout_13_0 alphaout_13_1
 alphaout_13_2 alphaout_13_3 alphaout_13_4 alphaout_13_5 alphaout_14_0
 alphaout_14_1 alphaout_14_2 alphaout_14_3 alphaout_14_4 alphaout_14_5
 alphaout_15_0 alphaout_15_1 alphaout_15_2 alphaout_15_3 alphaout_15_4
 alphaout_15_5 alphaout_16_0 alphaout_16_1 alphaout_16_2 alphaout_16_3
 alphaout_16_4 alphaout_16_5 alphaout_17_0 alphaout_17_1 alphaout_17_2
 alphaout_17_3 alphaout_17_4 alphaout_17_5 alpharet_0_0 alpharet_0_1
 alpharet_0_2 alpharet_0_3 alpharet_0_4 alpharet_0_5 alpharet_1_0
 alpharet_1_1 alpharet_1_2 alpharet_1_3 alpharet_1_4 alpharet_1_5
 alpharet_2_0 alpharet_2_1 alpharet_2_2 alpharet_2_3 alpharet_2_4
 alpharet_2_5 alpharet_3_0 alpharet_3_1 alpharet_3_2 alpharet_3_3
 alpharet_3_4 alpharet_3_5 alpharet_4_0 alpharet_4_1 alpharet_4_2
 alpharet_4_3 alpharet_4_4 alpharet_4_5 alpharet_5_0 alpharet_5_1
 alpharet_5_2 alpharet_5_3 alpharet_5_4 alpharet_5_5 alpharet_6_0
 alpharet_6_1 alpharet_6_2 alpharet_6_3 alpharet_6_4 alpharet_6_5
 alpharet_7_0 alpharet_7_1 alpharet_7_2 alpharet_7_3 alpharet_7_4
 alpharet_7_5 alpharet_8_0 alpharet_8_1 alpharet_8_2 alpharet_8_3
 alpharet_8_4 alpharet_8_5 alpharet_9_0 alpharet_9_1 alpharet_9_2
 alpharet_9_3 alpharet_9_4 alpharet_9_5 alpharet_10_0 alpharet_10_1
 alpharet_10_2 alpharet_10_3 alpharet_10_4 alpharet_10_5 alpharet_11_0
 alpharet_11_1 alpharet_11_2 alpharet_11_3 alpharet_11_4 alpharet_11_5
 alpharet_12_0 alpharet_12_1 alpharet_12_2 alpharet_12_3 alpharet_12_4
 alpharet_12_5 alpharet_13_0 alpharet_13_1 alpharet_13_2 alpharet_13_3
 alpharet_13_4 alpharet_13_5 alpharet_14_0 alpharet_14_1 alpharet_14_2
 alpharet_14_3 alpharet_14_4 alpharet_14_5 alpharet_15_0 alpharet_15_1
 alpharet_15_2 alpharet_15_3 alpharet_15_4 alpharet_15_5 alpharet_16_0
 alpharet_16_1 alpharet_16_2 alpharet_16_3 alpharet_16_4 alpharet_16_5
 alpharet_17_0 alpharet_17_1 alpharet_17_2 alpharet_17_3 alpharet_17_4
 alpharet_17_5 deltaout_0 deltaout_1 deltaout_2 deltaout_3 deltaout_4
 deltaout_5 deltaret_0 deltaret_1 deltaret_2 deltaret_3 deltaret_4
 deltaret_5
Generals
 nout_0 nout_1 nout_2 nout_3 nout_4 nout_5 nret_0 nret_1 nret_2 nret_3
 nret_4 nret_5 xstand_0 xstand_1 xstand_2 xstand_3 xstand_4 xstand_5
 xdisc_0 xdisc_1 xdisc_2 xdisc_3 xdisc_4 xdisc_5 ystand_0 ystand_1 ystand_2
 ystand_3 ystand_4 ystand_5 ydisc_0 ydisc_1 ydisc_2 ydisc_3 ydisc_4 ydisc_5
End
