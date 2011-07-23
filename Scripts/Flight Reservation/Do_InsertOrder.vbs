Option Explicit

''' <summary>
''' Implement a business layer
''' </summary>
''' <remarks></remarks>

Public oInsertOrderGUI

Class ClsDo_InsertOrder
	
	''' <summary>
	''' Execute a business layer
	''' </summary>
	''' <remarks></remarks>
	Public Default Function Run()
		
		Dim intStatus
		Set oInsertOrderGUI = InsertOrderGUI()
		If oInsertOrderGUI.Init() Then
			oInsertOrderGUI.SetDate
			'Create a passed step without expect/actual result
			oReport.ReportPass array("SetDate", "Set successfully"), false
			oInsertOrderGUI.SetFlyFrom
			'Create a passed step without expect/actual result
			oReport.ReportPass array("SetFlyFrom", "Set successfully"), False
			oInsertOrderGUI.SetFlyTo
			'Create a passed step without expect/actual result
			oReport.ReportPass array("SetFlyTo", "Set successfully"), False
			oInsertOrderGUI.SelectFlyLine
			'Create a passed step without expect/actual result
			oReport.ReportPass array("SelectFlyLine", "Select successfully"), False
			oInsertOrderGUI.SetName
			'Create a passed step without expect/actual result
			oReport.ReportPass array("SetName", "Set successfully"), False
			oInsertOrderGUI.InsertOrder
			'Create a passed step without expect/actual result
			oReport.ReportPass array("InsertOrder","InsertOrder successfully", "InsertOrder successfully"), false
		Else
			oReport.ReportFail array("GUI Layer initialization","All InsertOrder GUI objects should be loaded successfully", "Not All InsertOrder GUI objects have loaded successfully"), true	
		End if		
		
	End Function

End Class


''' <summary>
''' Create an instance of the Do_InsertOrder class and execute the business layer
''' </summary>
''' <remarks></remarks>
Public Function Do_InsertOrder()
	
	Dim InsertOrder
	Set InsertOrder = New ClsDo_InsertOrder
	InsertOrder.Run

End function