Option Explicit

''' <summary>
''' Implement a business layer
''' </summary>
''' <remarks></remarks>
Class ClsDo_Login
	
	''' <summary>
	''' Execute a business layer
	''' </summary>
	''' <remarks></remarks>
	Public Default Function Run(ByVal strURL, ByVal strUserid, ByVal strPwd)
		
		Dim intStatus, oLoginGUI
		Launch strURL
		Set oLoginGUI = LoginGUI()
		If oLoginGUI.Init() Then
			'Create a passed step without expect/actual result
			oReport.ReportPass array("GUI Layer initialization", "All Login GUI context objects have been loaded successfully"), false
			oLoginGUI.SetAgentName strUserid
			'Create a passed step without expect/actual result
			oReport.ReportPass array("SetAgentName", "Set successfully"), false
			oLoginGUI.SetPassword strPwd
			'Create a passed step without expect/actual result
			oReport.ReportPass array("SetPassword", "Set successfully"), false
			oLoginGUI.Submit
			'Create a passed step without expect/actual result
			oReport.ReportPass array("Submit","Login successfully", "Login successfully"), false
		Else
			oReport.ReportFail array("GUI Layer initialization","All Login GUI objects should be loaded successfully", "Not All Login GUI objects have loaded successfully"), true	
		End if		
		
	End Function
	
	''' <summary>
    ''' Launch specific process
    ''' </summary>
    ''' <param name="strURL" type="string">URL/Location Name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function Launch(ByVal strURL)
		
		CloseProcess oFSO.GetFileName(strURL)
		wait 1
		Dim oShell
		Set oShell = CreateObject("Wscript.shell")
		If oFSO.CheckFileExists(strURL) then
			oShell.Run Chr(34) & strURL & Chr(34)
		Else
			MsgBox "Please change to correct URL/Location for flight4a.exe in PreExecutionSetup.xls file under Config folder"	
		End If 
		
	End Function

	''' <summary>
    ''' Close all IE & Firefox Browsers Using WMI
    ''' </summary>
    ''' <param name="sProcessName" type="string">Process Name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function CloseProcess(ByVal sProcessName)

		Dim strSQL, oWMIService, ProcColl, oElem
		strSQL = "Select * From Win32_Process Where Name = '" & sProcessName & "'"
		Set oWMIService = GetObject("winmgmts:\\.\root\cimv2")
		Set ProcColl = oWMIService.ExecQuery(strSQL)
		For Each oElem in ProcColl
		    oElem.Terminate
		Next
		Set oWMIService = Nothing
	
	End Function

End Class


''' <summary>
''' Create an instance of the Do_login class and execute the business layer
''' </summary>
''' <remarks></remarks>
Public Function Do_Login(ByVal strURL,ByVal strUserid,ByVal strPwd)

	Dim Login
	Set Login = New ClsDo_Login
	Login.Run strURL, strUserid, strPwd

End function
