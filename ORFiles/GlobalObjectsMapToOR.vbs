
''' #########################################################
''' <summary>
''' Centralize all Objects definitions mapped to Object Repository that could be accessed from any script.
''' </summary>
''' <remarks>First clear the objects instances in memory</remarks>	 
''' #########################################################


''' <summary>
''' Flight Reservation Login screen Objects Mapping
''' </summary>
Set FR_Login_Dialog_OR = nothing
Set FR_Login_Dialog_OR = Dialog("Login")
Set FR_AgentName_OR = Nothing
Set FR_AgentName_OR = Dialog("Login").WinEdit("Agent Name:")
Set FR_Password_OR = Nothing
Set FR_Password_OR = Dialog("Login").WinEdit("Password:")
Set FR_OK_OR = Nothing
Set FR_OK_OR = Dialog("Login").WinButton("OK")

''' <summary>
''' Flight Reservation screen Objects Mapping
''' </summary>
Set FR_Flight_Reservation_Dialog_OR = Nothing
Set FR_Flight_Reservation_Dialog_OR = Window("Flight Reservation")
Set FR_Date_OR = Nothing
Set FR_Date_OR = Window("Flight Reservation").ActiveX("MaskEdBox")
Set FR_FlyFrom_OR = Nothing
Set FR_FlyFrom_OR = Window("Flight Reservation").WinComboBox("Fly From:")
Set FR_FlyTo_OR = Nothing
Set FR_FlyTo_OR = Window("Flight Reservation").WinComboBox("Fly To:")
Set FR_FLIGHT_button_OR = Nothing
Set FR_FLIGHT_button_OR = Window("Flight Reservation").WinButton("FLIGHT")
Set FR_Name_OR = Nothing
Set FR_Name_OR = Window("Flight Reservation").WinEdit("Name:")
Set FR_Insert_Order_button_OR = Nothing
Set FR_Insert_Order_button_OR = Window("Flight Reservation").WinButton("Insert Order")
Set FR_Open_Order_button_OR = Nothing
Set FR_Open_Order_button_OR = Window("Flight Reservation").WinButton("Button")
Set FR_Update_Order_button_OR = Nothing
Set FR_Update_Order_button_OR = Window("Flight Reservation").WinButton("Update Order")
Set FR_Delete_Order_button_OR = Nothing
Set FR_Delete_Order_button_OR = Window("Flight Reservation").WinButton("Delete Order")
Set FR_ProgressBar_OR = Nothing
Set FR_ProgressBar_OR = Window("Flight Reservation").ActiveX("Threed Panel Control")
Set FR_Yes_OR = Nothing
Set FR_Yes_OR = Window("Flight Reservation").Dialog("Flight Reservations").WinButton("Yes")
' Flights Table dialog
Set FR_FlyLineWinList_OR = Nothing
Set FR_FlyLineWinList_OR = Window("Flight Reservation").Dialog("Flights Table").WinList("From")
Set FR_Flights_Table_OK_OR = Nothing
Set FR_Flights_Table_OK_OR = Window("Flight Reservation").Dialog("Flights Table").WinButton("OK")
' Open Order Dialog
Set FR_CustomerName_checkbox_OR = Nothing
Set FR_CustomerName_checkbox_OR = Window("Flight Reservation").Dialog("Open Order").WinCheckBox("Customer Name")
Set FR_CustomerName_text_OR = Nothing
Set FR_CustomerName_text_OR = Window("Flight Reservation").Dialog("Open Order").WinEdit("Edit")
Set FR_Open_Order_OK_OR = Nothing
Set FR_Open_Order_OK_OR = Window("Flight Reservation").Dialog("Open Order").WinButton("OK")
Set FR_SearchResult_OK_OR = Nothing
Set FR_SearchResult_OK_OR = Window("Flight Reservation").Dialog("Open Order").Dialog("Search Results").WinButton("OK")