Option Explicit

''' <summary>
''' Implement a business layer
''' </summary>
''' <remarks></remarks>
Class ClsDo_DeleteOrder
	
	''' <summary>
	''' Execute a business layer
	''' </summary>
	''' <remarks></remarks>
	Public Default Function Run()
		
		Dim intStatus, oDeleteOrderGUI
		Set oDeleteOrderGUI = DeleteOrderGUI()
		If oDeleteOrderGUI.Init() And oUpdateOrderGUI.WaitUpdateOrderDone Then
			oDeleteOrderGUI.OpenOrder("Demo")
			'Create a passed step without expect/actual result
			oReport.ReportPass array("OpenOrder", "Open successfully"), false
			oDeleteOrderGUI.DeleteOrder
			oReport.ReportPass array("DeleteOrder","DeleteOrder successfully", "DeleteOrder successfully"), false
		Else
			oReport.ReportFail array("GUI Layer initialization","All DeleteOrder GUI objects should be loaded successfully", "Not All DeleteOrder GUI objects have loaded successfully"), true	
		End if		
		
	End Function

End Class


''' <summary>
''' Create an instance of the Do_DeleteOrder class and execute the business layer
''' </summary>
''' <remarks></remarks>
Public Function Do_DeleteOrder()

	Dim DeleteOrder
	Set DeleteOrder = New ClsDo_DeleteOrder
	DeleteOrder.Run

End function