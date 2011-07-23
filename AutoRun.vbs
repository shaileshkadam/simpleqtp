Option Explicit

''' ##################################################################
''' <summary>
''' Run one QTP Test automatically using QTP Automation Object Model
''' </summary>
''' <remarks>
''' Please make sure putting this VBS file into the same directory with QTP Test
''' Double-click this file will execute the QTP Test automatically
''' </remarks>
''' ##################################################################

Class ClsAOM

	Private qtApp

	''' <summary>
    ''' Class Initialization procedure.Create the QTP Application object
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Initialize()
		
		Set qtApp = CreateObject("QuickTest.Application") 
			
	End Sub
		
	''' <summary>
    ''' Class Initialization procedure. Release the QTP Application object
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Terminate()
		
		Call EnableScreenSaver
		Set qtApp = nothing
		
	End Sub
	
	''' <summary>
    ''' Auto Run QTP using QTP Automation Object Model
    ''' </summary>
    ''' <param name="sTestName" type="string">The name of the Test</param>
    ''' <remarks></remarks>
	Public Function AutoRunQTP(ByVal sTestName)
		
		Dim objShell,sQTPProjetLocation
		Set objShell = CreateObject("Wscript.Shell")
		sQTPProjetLocation = objShell.CurrentDirectory & "\" & sTestName
		AutoRunExistingTest(sQTPProjetLocation)
	
	End Function
	
	''' <summary>
    ''' Auto Launch QTP, open an existing test and Run the Test
    ''' </summary>
    ''' <param name="sQTPProjetLocation" type="string">The location of the QTP Test</param>
    ''' <remarks></remarks>
	Public Function AutoRunExistingTest(ByVal sQTPProjetLocation)
		
		Dim qtTest
		'If QTP is notopen then open it
		If  qtApp.launched <> True then 
			qtApp.Launch 
		End If 
		
		'Make the QuickTest application visible
		qtApp.Visible = True
		
		'Set QuickTest run options
		'Instruct QuickTest to perform next step when error occurs
		qtApp.Options.Run.ImageCaptureForTestResults = "OnError"
		qtApp.Options.Run.RunMode = "Fast"
		qtApp.Options.Run.ViewResults = true
		
		'Open the test in read-only mode
		qtApp.Open sQTPProjetLocation, True 
		
		'set run settings for the test
		Set qtTest = qtApp.Test
		
		'Instruct QuickTest to perform next step when error occurs
		qtTest.Settings.Run.OnError = "NextStep" 
		
		'Run the test
		qtTest.Run
		
		'Check the results of the test run
		'MsgBox qtTest.LastRunResults.Status
		
		'Close the test
		qtTest.Close 
		
		'Close QTP
		qtApp.quit
		
		'Release Object
		Set qtTest = Nothing
	
	End Function
	
	''' <summary>
    ''' Enable windows 'ScreenSaver'
    ''' </summary>
    ''' <remarks></remarks>
	Public Function EnableScreenSaver()
	
		Dim WshShell  
		Set WshShell = WScript.CreateObject("WScript.Shell")  
		WshShell.RegWrite "HKCU\Control Panel\Desktop\ScreenSaveActive",1,"REG_SZ" 
		Set WshShell = Nothing
	
	End Function
	

End Class

Public Function AOM()
	
	Set AOM = New ClsAOM

End Function

''' <summary>
''' Please change to your Test Name
''' </summary>
Call AOM.AutoRunQTP("MainScript")

