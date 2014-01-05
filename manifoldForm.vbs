Sub Main
	Form.Visible = True
End Sub


Sub WLCOMMAND1_Click

	Set CompS = document.ComponentSet
	Set StreamClip = CompS("StreamClip")
	Set A = CompS("A")
	Set B = CompS("B")
	Set LU = CompS("LU")
	Set LUC = CompS("LUC")
	Set WW = CompS("WaterWays")
	Set P = CompS("P")
	Set RV = CompS("RV")
	Set ST = CompS("ST")
	Set BF1Q = CompS("BF1Q")
	Set BF2Q = CompS("BF2Q")
	Set BF3Q = CompS("BF3Q")
	Set BF3F = CompS("BF3F")
	Set BF2F = CompS("BF2F")
	CompS.Remove("BF1T")
	CompS.Remove("BF2T")
	CompS.Remove("BF3T")
	Set BF1D = CompS("BF1T 2")
	Set BF2D = CompS("BF2F 2")
	Set BF3D = CompS("BF3F 2")
	Set WTQ = CompS("WTQ")
	Set WW1 = CompS("WW1")
	Set WW2 = CompS("WW2")
	Set WW3 = CompS("WW3")
	Set BF1W = CompS("BF1W")
	CompS.Remove("WeightsTable")
	Set Agriculture = CompS("Agriculture")
	Set BarrenLand = CompS("BarrenLand")
	Set Commercial = CompS("Commercial")
	Set Forest = CompS("Forest")
	Set Industrial = CompS("Industrial")
	Set Institutional = CompS("Institutional")
	Set LowDensRes = CompS("LowDensRes")
	Set MedHighDensRes = CompS("MedHighDensRes")
	Set Other = CompS("Other")
	Set Transportation = CompS("Transportation")
	Set VeryLowDensRes = CompS("VeryLowDensRes")
	Set Water = CompS("Water")
	Set Wetland = CompS("Wetland")	
	Set LUColumn = CompS("LUColumn")
	Set Slope = CompS("Slope")
	Set AVGSlopeC = CompS("AVGSlopeC")
	Set SlopeWeightColumn = CompS("SlopeWeightColumn")
	Set SlopeWeightH = CompS("SlopeWeightH")
	Set SlopeWeightL = CompS("SlopeWeightL")
	Set SlopeWeightM = CompS("SlopeWeightM")
	Set BuildingsWeightC = CompS("BuildingsWeightC")
	Set BuildingsH = CompS("BuildingsH")
	Set BuildingsM = CompS("BuildingsM")
	Set BuildingsL = CompS("BuildingsL")
	Set SQFTAREAB = CompS("SQFTAREAB")
	Set FinalSumWeight = CompS("FinalSumWeight")
	Set FinalSumC = CompS("FinalSumC")


	waterLBufDist = WLText9.Text
	waterMBufDist = WLText10.Text
	waterHBufDist = WLText11.Text
	waterLUnit = WLCombo1.Text
	waterMUnit = WLCombo2.Text
	waterHUnit = WLCombo3.Text
	waterLWeight = WLText14.Text
	waterMWeight = WLText15.Text
	waterHWeight = WLText16.Text

	BF1Q.Text = "SELECT Unionall(Buffer([WaterWays].[Geom (I)], "&WaterLBufDist&", """&waterLUnit&""")) into BF1T FROM [WaterWays]"
	BF1Q.Run
	BF2Q.Text = "SELECT UnionAll(Buffer([WaterWays].[Geom (I)], "&WaterMBufDist&", """&waterMUnit&""")) into BF2T FROM [WaterWays]"
	BF2Q.Run
	BF3Q.Text = "SELECT UnionAll(Buffer([WaterWays].[Geom (I)], "&WaterHBufDist&", """&waterHUnit&""")) into BF3T FROM [WaterWays]"
	BF3Q.Run
	Set BF1T = CompS("BF1T")
	Set BF2T = CompS("BF2T")
	Set BF3T = CompS("BF3T")
	BF2F.Run
	BF3F.Run
	BF1D.Refresh
	BF2D.Refresh
	BF3D.Refresh
	WTQ.Run
	Set WeightsTable = CompS("WeightsTable")
	BF1W.Run
	WW1.Text = "UPDATE (SELECT [BF3F 2].[ID], [WeightsTable].[WaterWeight]" &_
				"FROM [BF3F 2], [WeightsTable]" &_
				"WHERE INTERSECTS([BF3F 2].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
				"SET [WaterWeight] = "&waterHWeight&""
	WW2.Text = "UPDATE (SELECT [BF2F 2].[ID], [WeightsTable].[WaterWeight]" &_
				"FROM [BF2F 2], [WeightsTable]" &_
				"WHERE INTERSECTS([BF2F 2].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
				"SET [WaterWeight] = "&waterHWeight&""
	WW3.Text = "UPDATE (SELECT [BF1T 2].[ID], [WeightsTable].[WaterWeight]" &_
				"FROM [BF1T 2], [WeightsTable]" &_
				"WHERE INTERSECTS([BF1T 2].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
				"SET [WaterWeight] = "&waterHWeight&""
	WW1.RunEx true
	WW2.RunEx true
	WW3.RunEx true
	LUColumn.Run

	Water.Text = "UPDATE (SELECT [LUC].[ID], [WeightsTable].[LandUseWeight]" &_
					"FROM [LUC], [WeightsTable]" &_
					"WHERE ([LUC].[LU_CODE] = 50) AND TOUCHES([LUC].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
					"SET [LandUseWeight] = 1"
	Wetland.Text = "UPDATE (SELECT [LUC].[ID], [WeightsTable].[LandUseWeight]" &_
					"FROM [LUC], [WeightsTable]" &_
					"WHERE ([LUC].[LU_CODE] = 60) AND TOUCHES([LUC].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
					"SET [LandUseWeight] = 1"
	Forest.Text = "UPDATE (SELECT [LUC].[ID], [WeightsTable].[LandUseWeight]" &_
					"FROM [LUC], [WeightsTable]" &_
					"WHERE ([LUC].[LU_CODE] = 41 or [luc].[LU_CODE] = 42 or [luc].[LU_CODE] = 43 or [luc].[LU_CODE] = 44 or [luc].[LU_CODE] = 40) AND TOUCHES([LUC].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
					"SET [LandUseWeight] = 1"
	VeryLowDensRes.Text ="UPDATE (SELECT [LUC].[ID], [WeightsTable].[LandUseWeight]" &_
							"FROM [LUC], [WeightsTable]" &_
							"WHERE ([LUC].[LU_CODE] = 191 or [luc].[LU_CODE] = 192) AND TOUCHES([LUC].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
							"SET [LandUseWeight] = 2"
	BarrenLand.Text = "UPDATE (SELECT [LUC].[ID], [WeightsTable].[LandUseWeight]" &_
						"FROM [LUC], [WeightsTable]" &_
						"WHERE ([LUC].[LU_CODE] = 71 or [luc].[LU_CODE] = 73) AND TOUCHES([LUC].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
						"SET [LandUseWeight] = 2"
	Agriculture.Text = "UPDATE (SELECT [LUC].[ID], [WeightsTable].[LandUseWeight]" &_
						"FROM [LUC], [WeightsTable]" &_
						"WHERE ([LUC].[LU_CODE] = 21 or [luc].[LU_CODE] = 22 or [luc].[LU_CODE] = 23 or [luc].[LU_CODE] = 25 or [luc].[LU_CODE] = 241 or [luc].[LU_CODE] = 242 or [luc].[LU_CODE] = 20) AND TOUCHES([LUC].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
						"SET [LandUseWeight] = 5"
	LowDensRes.Text = "UPDATE (SELECT  [LUC].[ID], [WeightsTable].[LandUseWeight]" &_
						"FROM [LUC], [WeightsTable]" &_
						"WHERE ([LUC].[LU_CODE] = 11 AND TOUCHES([LUC].[Geom (I)], [WeightsTable].[Geom (I)]) ))" &_
						"SET [weightstable].[LandUseWeight] = 5"
	Transportation.Text = "UPDATE (SELECT [LUC].[ID], [WeightsTable].[LandUseWeight]" &_
							"FROM [LUC], [WeightsTable]" &_
							"WHERE ([LUC].[LU_CODE] = 80) AND TOUCHES([LUC].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
							"SET [LandUseWeight] = 5"
	Commercial.Text = "UPDATE (SELECT [LUC].[ID], [WeightsTable].[LandUseWeight]" &_
						"FROM [LUC], [WeightsTable]" &_
						"WHERE ([LUC].[LU_CODE] = 14) AND TOUCHES([LUC].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
						"SET [LandUseWeight] = 7"
	Industrial.Text = "UPDATE (SELECT [LUC].[ID], [WeightsTable].[LandUseWeight]" &_
						"FROM [LUC], [WeightsTable]" &_
						"WHERE ([LUC].[LU_CODE] = 15) AND TOUCHES([LUC].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
						"SET [LandUseWeight] = 7"
	Institutional.Text = "UPDATE (SELECT [LUC].[ID], [WeightsTable].[LandUseWeight]" &_
							"FROM [LUC], [WeightsTable]" &_
							"WHERE ([LUC].[LU_CODE] = 16) AND TOUCHES([LUC].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
							"SET [LandUseWeight] = 7"
	MedHighDensRes.Text = "UPDATE (SELECT [LUC].[ID], [WeightsTable].[LandUseWeight]" &_
							"FROM [LUC], [WeightsTable]" &_
							"WHERE ([LUC].[LU_CODE] = 12 or [luc].[LU_CODE] = 13) AND TOUCHES([LUC].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
							"SET [LandUseWeight] = 7"
	Other.Text = "UPDATE (SELECT [LUC].[ID], [WeightsTable].[LandUseWeight]" &_
					"FROM [LUC], [WeightsTable]" &_
					"WHERE ([LUC].[LU_CODE] = 17 or [luc].[LU_CODE] = 18) AND TOUCHES([LUC].[Geom (I)], [WeightsTable].[Geom (I)]) )" &_
					"SET [LandUseWeight] = 10"

	Water.RunEx true
	Wetland.RunEx true
	Forest.RunEx true
	VeryLowDensRes.RunEx true	
	BarrenLand.RunEx true
	Agriculture.RunEx true
	LowDensRes.RunEx true
	Transportation.RunEx true	
	Commercial.RunEx true
	Industrial.RunEx true
	Institutional.RunEx true
	MedHighDensRes.RunEx true
	Other.RunEx true

	AVGSlopeC.Runex True
	SlopeWeightColumn.RunEx True
	Slope.Runex true
	SlopeWeightM.Text = "UPDATE [WeightsTable] SET [SlopeWeight] = 3 WHERE [AvgSlope] > 2 AND [AvgSlope] < 7"
	SlopeWeightH.Text = "UPDATE [WeightsTable] SET [SlopeWeight] = 5 WHERE [AvgSlope] > 7"
	SlopeWeightL.Text = "UPDATE [WeightsTable] SET [SlopeWeight] = 1 WHERE [AvgSlope] < 2"
	SlopeWeightM.RunEx True
	SlopeWeightH.RunEx True
	SlopeWeightL.RunEx True

	BuildingsWeightC.RunEx True	
	BuildingsL.Text = "UPDATE(SELECT [WeightsTable].[BuildingsWeight], [bc].[ID], [ResidentialParcelsD].[ID]" &_
						"FROM [WeightsTable], [bc], [ResidentialParcelsD]" &_
						"WHERE ([bc].[BuildingSqFTArea] < 1000 )" &_
						"AND Touches ([bc].[Geom (I)], [ResidentialParcelsD].[Geom (I)])" &_
						"And Touches ([ResidentialParcelsD].[Geom (I)], [WeightsTable].[Geom (I)]))" &_
						"Set [BuildingsWeight] = 1"
	BuildingsL.RunEx True
	BuildingsM.Text = "UPDATE(SELECT [WeightsTable].[BuildingsWeight], [bc].[ID], [ResidentialParcelsD].[ID]" &_
							"FROM [WeightsTable], [bc], [ResidentialParcelsD]" &_
							"WHERE ([bc].[BuildingSqFTArea] < 2000)" &_
							"AND ([bc].[BuildingSqFTArea] > 1000)" &_
							"AND Touches ([bc].[Geom (I)], [ResidentialParcelsD].[Geom (I)])" &_
							"And Touches ([ResidentialParcelsD].[Geom (I)], [WeightsTable].[Geom (I)]))" &_
							"Set [BuildingsWeight] = 2"
	BuildingsM.RunEx True
	BuildingsH.Text = "UPDATE(SELECT [WeightsTable].[BuildingsWeight], [bc].[ID], [ResidentialParcelsD].[ID]" &_
						"FROM [WeightsTable], [bc], [ResidentialParcelsD]" &_
						"WHERE ([bc].[BuildingSqFTArea] > 2000 )" &_
						"AND Touches ([bc].[Geom (I)], [ResidentialParcelsD].[Geom (I)])" &_
						"And Touches ([ResidentialParcelsD].[Geom (I)], [WeightsTable].[Geom (I)]))" &_
						"Set [BuildingsWeight] = 4"
	BuildingsH.RunEx True
	
	FinalSumC.RunEx True
	FinalSumWeight.RunEX True
	WeightsTable.Open
End Sub