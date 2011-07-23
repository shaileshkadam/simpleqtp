Option Explicit

''' <summary>
''' Define a container of the Login GUI Objects to implement a GUI layer
''' </summary>
''' <remarks></remarks>
Class ClsLoginGUI

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
			.Add "Login_Dialog", FR_Login_Dialog_OR
			.Add "AgentName", FR_AgentName_OR
			.Add "Password", FR_Password_OR
			.Add "OK", FR_OK_OR
		End With
		
		' iterates through the dictonary and check if the GUI objects "exist"
		Init = GUILayerContext.IsLoaded(oDictChildObjects, 2, 10)
	
	End Function

	
	''' <summary>
    ''' Set the Agent Name
    ''' </summary>
    ''' <param name="username" type="string">Agent Name</param>
    ''' <remarks></remarks>
	Public Function SetAgentName(ByVal username)
		
		Dim strTimeStamp, AgentName
		oDictChildObjects("AgentName").Set username
		
	End Function
	
	''' <summary>
    ''' Set the Password
    ''' </summary>
    ''' <param name="pwd" type="string">Password</param>
    ''' <remarks></remarks>
	Public Function SetPassword(ByVal pwd)
	
		oDictChildObjects("Password").Set pwd
		
	End Function
	
	''' <summary>
    ''' Press the OK button
    ''' </summary>
    ''' <remarks></remarks>
	Public Function Submit()
	
		oDictChildObjects("OK").click
		
	End Function	
		

End Class


''' <summary>
''' Create an instance of the login class
''' </summary>
''' <remarks></remarks>
Public Function LoginGUI()

	Set LoginGUI = New ClsLoginGUI

End function
