Option Explicit

''' <summary>
''' Define a container of the UpdateOrder GUI Objects to implement a GUI layer
''' </summary>
''' <remarks></remarks>
Class ClsUpdateOrderGUI

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
			.Add "Date", FR_Date_OR
			.Add "FlyFrom", FR_FlyFrom_OR
			.Add "FlyTo", FR_FlyTo_OR
			.Add "FLIGHT_button", FR_FLIGHT_button_OR
		    .Add "Name", FR_Name_OR
			.Add "Update Order", FR_Update_Order_button_OR
			.Add "ProgressBar", FR_ProgressBar_OR
		End With

		' iterates through the dictonary and check if the GUI objects "exist"
		Init = GUILayerContext.IsLoaded(oDictChildObjects, 2, 10)
	
	End Function	
	
	''' <summary>
    ''' Open the specific Order
    ''' </summary>
    ''' <remarks></remarks>
	Public Function OpenOrder(ByVal strOrderName)
		
		Dim index, oDictOpenOrderChildObjects, oDictSearchResultChildObjects
		' Click "Open Order" button
		oDictChildObjects("Open Order").click
		
		' Define child objects for "Open Order" dialog and define the operations		
		Set oDictOpenOrderChildObjects = CreateObject("Scripting.Dictionary")
		With oDictOpenOrderChildObjects
			.Add "CustomerName_checkbox", FR_CustomerName_checkbox_OR
			.Add "CustomerName_text", FR_CustomerName_text_OR
			.Add "OK", FR_Open_Order_OK_OR
		End With
		If GUILayerContext.IsLoaded(oDictOpenOrderChildObjects, 1, 10) then
			oDictOpenOrderChildObjects("CustomerName_checkbox").Set "ON"
			oDictOpenOrderChildObjects("CustomerName_text").Set strOrderName
			oDictOpenOrderChildObjects("OK").click
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
		Set oDictOpenOrderChildObjects = nothing
		
	End Function
	
	''' <summary>
    ''' Set Date of flight
    ''' </summary>
    ''' <remarks></remarks>
	Public Function SetDate()
		
		Dim i
		For i = 0 To 5
			oDictChildObjects("Date").Type micDel
		next
		oDictChildObjects("Date").Type oDateTimeLib.ReFormatDateTime(DateAdd("d", 1, Now), "MMDDYY")
		
	End Function
	
	''' <summary>
    ''' Set the FlyFrom
    ''' </summary>
    ''' <remarks></remarks>
	Public Function SetFlyFrom()

		Dim index
		index = Int((oDictChildObjects("FlyFrom").GetROProperty("items count") - 1 - 0 + 1) * Rnd + 0)'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
		oDictChildObjects("FlyFrom").select(index)
		
	End Function
	
	''' <summary>
    ''' Set the FlyTo
    ''' </summary>
    ''' <remarks></remarks>
	Public Function SetFlyTo()
		
		Dim index
		index = Int((oDictChildObjects("FlyTo").GetROProperty("items count") - 1 - 0 + 1) * Rnd + 0)'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
		oDictChildObjects("FlyTo").select(index)
		
	End Function
	
	''' <summary>
    ''' Select the Fly Line
    ''' </summary>
    ''' <remarks></remarks>
	Public Function SelectFlyLine()
		
		Dim index, oDictFlightsTableChildObjects
		oDictChildObjects("FLIGHT_button").click
		Set oDictFlightsTableChildObjects = CreateObject("Scripting.Dictionary")
		With oDictFlightsTableChildObjects
			.Add "FlyLineWinList", FR_FlyLineWinList_OR
			.Add "OK", FR_Flights_Table_OK_OR
		End With
		If GUILayerContext.IsLoaded(oDictFlightsTableChildObjects, 1, 10) then
			index = Int((oDictFlightsTableChildObjects("FlyLineWinList").GetROProperty("items count") - 1 - 0 + 1) * Rnd + 0)'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
			oDictFlightsTableChildObjects("FlyLineWinList").select(index)
			oDictFlightsTableChildObjects("OK").click
		End if
		Set oDictFlightsTableChildObjects =	nothing
		
	End Function
	
	''' <summary>
    ''' Set the Name
    ''' </summary>
    ''' <remarks></remarks>
	Public Function SetName()
		
		oDictChildObjects("Name").set "Demo1"
		
	End Function
	
	''' <summary>
    ''' Press the OK button
    ''' </summary>
    ''' <remarks></remarks>
	Public Function UpdateOrder()
	
		oDictChildObjects("Update Order").click
		
	End Function	
	
	''' <summary>
    ''' Wait UpdateOrder Done
    ''' </summary>
    ''' <remarks></remarks>
	Public Function WaitUpdateOrderDone()
	
		WaitUpdateOrderDone = oDictChildObjects("ProgressBar").WaitProperty("text", "Update Done...", 10000)
		
	End Function
	

End Class


''' <summary>
''' Create an instance of the UpdateOrder class
''' </summary>
''' <remarks></remarks>
Public Function UpdateOrderGUI()

	Set UpdateOrderGUI = New ClsUpdateOrderGUI

End function