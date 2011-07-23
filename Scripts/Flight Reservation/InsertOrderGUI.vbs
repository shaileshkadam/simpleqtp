Option Explicit

''' <summary>
''' Define a container of the InsertOrder GUI Objects to implement a GUI layer
''' </summary>
''' <remarks></remarks>
Class ClsInsertOrderGUI

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
			.Add "Date", FR_Date_OR
			.Add "FlyFrom", FR_FlyFrom_OR
			.Add "FlyTo", FR_FlyTo_OR
			.Add "FLIGHT_button", FR_FLIGHT_button_OR
		    .Add "Name", FR_Name_OR
			.Add "Insert Order", FR_Insert_Order_button_OR
			.Add "ProgressBar", FR_ProgressBar_OR
		End With

		' iterates through the dictonary and check if the GUI objects "exist"
		Init = GUILayerContext.IsLoaded(oDictChildObjects, 2, 10)
	
	End Function	
	
	''' <summary>
    ''' Set Date of flight
    ''' </summary>
    ''' <remarks></remarks>
	Public Function SetDate()
	
		oDictChildObjects("Date").Type oDateTimeLib.ReFormatDateTime(DateAdd("d",1, Now), "MMDDYY")
		
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
		
		oDictChildObjects("Name").set "Demo"
		
	End Function
	
	''' <summary>
    ''' Press the OK button
    ''' </summary>
    ''' <remarks></remarks>
	Public Function InsertOrder()
	
		oDictChildObjects("Insert Order").click
		
	End Function

	''' <summary>
    ''' Wait InsertOrder Done
    ''' </summary>
    ''' <remarks></remarks>
	Public Function WaitInsertOrderDone()
	
		WaitInsertOrderDone = oDictChildObjects("ProgressBar").WaitProperty("text","Insert Done...",10000)
		
	End Function
	
End Class

''' <summary>
''' Create an instance of the InsertOrder class
''' </summary>
''' <remarks></remarks>
Public Function InsertOrderGUI()

	Set InsertOrderGUI = New ClsInsertOrderGUI

End function