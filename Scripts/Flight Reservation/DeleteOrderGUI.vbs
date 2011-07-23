Option Explicit

''' <summary>
''' Define a container of the DeleteOrder GUI Objects to implement a GUI layer
''' </summary>
''' <remarks></remarks>
Class ClsDeleteOrderGUI

	Private oDictChildObjects
	
	''' <summary>
    ''' Class Initialization procedure. Creates Child Objects dictionary.
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Initialize()

		Set oDictChildObjects = CreateObject("Scripting.Dictionary")
				
	end sub

	''' <summary>
    ''' Class Termination procedure, Release Child Objects dictionary.
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Terminate()
 
		Set oDictChildObjects = nothing
		
	end sub
	
	''' <summary>
    ''' initialized the current and child objects
    ''' </summary>
    ''' <return>True/False</return>
    ''' <remarks></remarks>
	Public Function Init()

		With oDictChildObjects
			.Add "Flight Reservation", FR_Flight_Reservation_Dialog_OR
			.Add "Open Order", FR_Open_Order_button_OR
			.Add "Delete Order", FR_Delete_Order_button_OR
		End With

		' iterates through the dictonary and check if the GUI objects "exist"
		Init = GUILayerContext.IsLoaded(oDictChildObjects, 2, 10)
	
	End Function	
	
	''' <summary>
    ''' Open the specific Order
    ''' </summary>
    ''' <remarks></remarks>
	Public Function OpenOrder(ByVal strOrderName)
		
		Dim index, oDictFlightsTableChildObjects, oDictSearchResultChildObjects
		' Click "Open Order" button
		oDictChildObjects("Open Order").click
		
		' Define child objects for "Open Order" dialog and define the operations		
		Set oDictFlightsTableChildObjects = CreateObject("Scripting.Dictionary")
		With oDictFlightsTableChildObjects
			.Add "CustomerName_checkbox", FR_CustomerName_checkbox_OR
			.Add "CustomerName_text", FR_CustomerName_text_OR
			.Add "OK", FR_Open_Order_OK_OR
		End With
		If GUILayerContext.IsLoaded(oDictFlightsTableChildObjects, 1, 10) then
			oDictFlightsTableChildObjects("CustomerName_checkbox").Set "ON"
			oDictFlightsTableChildObjects("CustomerName_text").Set strOrderName
			oDictFlightsTableChildObjects("OK").click
		End If
		
		' Define child objects for "Search Results" dialog and define the operations
		Set oDictSearchResultChildObjects = CreateObject("Scripting.Dictionary")
		With oDictSearchResultChildObjects
			.Add "SearchResult_OK", FR_SearchResult_OK_OR
		End With
		If GUILayerContext.IsLoaded(oDictSearchResultChildObjects, 1, 10) then
			oDictSearchResultChildObjects("SearchResult_OK").click
		End if	
		' Release the dictionary
		Set oDictSearchResultChildObjects = nothing
		Set oDictFlightsTableChildObjects = nothing
		
	End Function
	
	''' <summary>
    ''' Press the OK button
    ''' </summary>
    ''' <remarks></remarks>
	Public Function DeleteOrder()
		
		Dim oDictFlightsTableChildObjects
		Set oDictFlightsTableChildObjects = CreateObject("Scripting.Dictionary")
		oDictChildObjects("Delete Order").click
		With oDictFlightsTableChildObjects
			.Add "Yes", FR_Yes_OR
		End With
		If GUILayerContext.IsLoaded(oDictFlightsTableChildObjects, 1, 10) then
			oDictFlightsTableChildObjects("Yes").click
		End if	
		Set oDictFlightsTableChildObjects = Nothing
		
	End Function	

End Class


''' <summary>
''' Create an instance of the DeleteOrder class
''' </summary>
''' <remarks></remarks>
Public Function DeleteOrderGUI()
	
	Set DeleteOrderGUI = New ClsDeleteOrderGUI
	
End function