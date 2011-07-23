Option Explicit

''' <summary>
''' Implement a business layer
''' </summary>
''' <remarks></remarks>

Public oUpdateOrderGUI

Class ClsDo_UpdateOrder
	
	''' <summary>
	''' Execute a business layer
	''' </summary>
	''' <remarks></remarks>
	Public Default Function Run()
		
		Dim intStatus
		Set oUpdateOrderGUI = UpdateOrderGUI()
		If oUpdateOrderGUI.Init() and oInsertOrderGUI.WaitInsertOrderDone Then
			oUpdateOrderGUI.OpenOrder("Demo")
			'Create a passed step without expect/actual result
			oReport.ReportPass array("OpenOrder", "Open successfully"), false
			oUpdateOrderGUI.SetDate
			oReport.ReportPass array("SetDate", "Set successfully"), false
			oUpdateOrderGUI.SetFlyFrom
			oReport.ReportPass array("SetFlyFrom", "Set successfully"), False
			oUpdateOrderGUI.SetFlyTo
			oReport.ReportPass array("SetFlyTo", "Set successfully"), False
			oUpdateOrderGUI.SelectFlyLine
			oReport.ReportPass array("SelectFlyLine", "Select successfully"), False
			oUpdateOrderGUI.SetName
			oReport.ReportPass array("SetName", "Set successfully"), False
			oUpdateOrderGUI.UpdateOrder
			oReport.ReportPass array("UpdateOrder","UpdateOrder successfully", "UpdateOrder successfully"), false
		Else
			oReport.ReportFail array("GUI Layer initialization","All UpdateOrder GUI objects should be loaded successfully", "Not All UpdateOrder GUI objects have loaded successfully"), true	
		End if		
		
	End Function

End Class


''' <summary>
''' Create an instance of the Do_UpdateOrder class and execute the business layer
''' </summary>
''' <remarks></remarks>
Public Function Do_UpdateOrder()

	Dim UpdateOrder
	Set UpdateOrder = New ClsDo_UpdateOrder
	UpdateOrder.Run

End function