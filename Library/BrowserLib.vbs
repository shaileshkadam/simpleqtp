Option Explicit

''' #########################################################
''' <summary>
''' A library to work with Browser object
''' </summary>
''' <remarks></remarks>	 
''' #########################################################

Class ClsBrowserLib

    ''' <summary>
    ''' Activate specific Browser
    ''' </summary>
    ''' <param name="objBrowser" type="object">Browser object</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function BrowserActivate(ByVal objBrowser)
		
		Dim hWnd
		hWnd = objBrowser.GetROProperty("hwnd")
		On Error Resume Next
			Window("hwnd:=" & hWnd).Activate
			If Err.Number <> 0 Then
				Window("hwnd:=" & Browser("hwnd:=" & hWnd).Object.hWnd).Activate
				Err.Clear
			End If
		On Error Goto 0
		
	End Function

    ''' <summary>
    ''' Maximize specific Browser
    ''' </summary>
    ''' <param name="objBrowser" type="object">Browser object</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function BrowserMaximize(ByVal objBrowser)
		
		Dim hWnd
		hWnd = objBrowser.GetROProperty("hwnd")
		On Error Resume Next
			Window("hwnd:=" & hWnd).Activate
			If Err.Number <> 0 Then
				hWnd = Browser("hwnd:=" & hWnd).Object.hWnd
				Window("hwnd:=" & hWnd).Activate
				Err.Clear
			End If
			Window("hwnd:=" & hWnd).Maximize
		On Error Goto 0
		
	End Function

	''' <summary>
    ''' Minimize specific Browser
    ''' </summary>
    ''' <param name="objBrowser" type="object">Browser object</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function BrowserMinimize(ByVal objBrowser)
		
		Dim hWnd
		hWnd = objBrowser.GetROProperty("hwnd")
		On Error Resume Next
			Window("hwnd:=" & hWnd).Activate
			If Err.Number <> 0 Then
				hWnd = Browser("hwnd:=" & hWnd).Object.hWnd
				Window("hwnd:=" & hWnd).Activate
				Err.Clear
			End If
			Window("hwnd:=" & hWnd).Minimize
		On Error Goto 0
		
	End Function

	''' <summary>
    ''' Close all Browsers except specific Browser
    ''' </summary>
    ''' <param name="sBrowserName" type="string">Browser Name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function CloseAllBrowserExcept(ByVal sBrowserName)
	
		Dim oDesc, x
		'Create a description object
		Set oDesc = Description.Create
		oDesc( "micclass" ).Value = "Browser"
		'Close all browsers except Quality Center
		If Desktop.ChildObjects(oDesc).Count > 0 Then
		    For x = Desktop.ChildObjects(oDesc).Count - 1 To 0 Step -1
		       If InStr(1, Browser("creationtime:="&x).GetROProperty("name"), sBrowserName) = 0 Then  
		          Browser( "creationtime:=" & x ).Close
		       End If
		    Next
		End If
	
	End Function
	
	''' <summary>
    ''' Close all IE & Firefox Browsers Using WMI
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function CloseAllBrowser()
		
		Dim strSQL, oWMIService, ProcColl, oElem
		strSQL = "Select * From Win32_Process Where Name = 'iexplore.exe' OR Name = 'firefox.exe'"
		Set oWMIService = GetObject("winmgmts:\\.\root\cimv2")
		Set ProcColl = oWMIService.ExecQuery(strSQL)
		For Each oElem in ProcColl
		    oElem.Terminate
		Next
		Set oWMIService = Nothing
	
	End function

End Class
 
Public Function BrowserLib()
	
	Set BrowserLib = New ClsBrowserLib

End Function