Option Explicit

'Declare an instance of the ExcelLib class
Public oExcelLib
'Declare an Global instance of the ReportLib class
Public oReport
'Declare an Global instance of the FSOLib class
Public oFSO
'Declare an Global instance of the DateTimeLib class
Public oDateTimeLib
	
''' <summary>
''' Implement Flow Engine
''' </summary>
''' <remarks></remarks>
Class ClsFlowEngine
		
	''' <summary>
    ''' Class Initialization procedure
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Initialize()

		'Load all requisite resources
		Call LoadResources	
		'New an instance of the ExcelLib class
		Set oExcelLib = ExcelLib
		'New an instance of the ReportLib class
		Set oReport = ReportLib
		'New an instance of the FSOLib class
		Set oFSO = FSOLib
		'New an instance of the DateTimeLib class
		Set oDateTimeLib = DateTimeLib
				
	end Sub
	
	
	''' <summary>
    ''' Class Termination procedure
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Terminate()
 
		Set oExcelLib = Nothing
		Set oReport = Nothing
		Set oFSO = Nothing
		Set oDateTimeLib = nothing

	end sub
	
	''' <summary>
	''' Load all requisite resources
	''' </summary>
	''' <remarks></remarks>
	Public Function LoadResources
	
		'Load Global constant
		Dim ConfigPath, objFSO
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		ConfigPath =  objFSO.GetParentFolderName(PathFinder.CurrentTestPath) & "\Config\"
		ExecuteFile ConfigPath & "GlobalConstants.vbs"
		
		'Load Specific function libraries
		ExecuteFile GlobalConst.ExcelLib
		ExecuteFile GlobalConst.FSOLib
		ExecuteFile GlobalConst.DateTimeLib
		'ExecuteFile GlobalConst.XlsDictLib
		'ExecuteFile GlobalConst.CompareTwoExcelLib
		'ExecuteFile GlobalConst.ArrayLib
		'ExecuteFile GlobalConst.RegLib
		'ExecuteFile GlobalConst.StringLib
		'ExecuteFile GlobalConst.WordLib
		'ExecuteFile GlobalConst.EmailLib
		'ExecuteFile GlobalConst.LoggerLib
		'ExecuteFile GlobalConst.ObjectGeneralLib
		'ExecuteFile GlobalConst.ZipLib
	
		'Load report library
		ExecuteFile GlobalConst.ReportLib

		'Load Corresponding OR Files
		Dim qtApp, qtRepositories
		Set qtApp = CreateObject("QuickTest.Application")
		'Get the object repositories collection object of the action1
		Set qtRepositories = qtApp.Test.Actions(1).ObjectRepositories
		qtRepositories.RemoveAll
		'If the repository cannot be found in the collection
		If qtRepositories.Find(GlobalConst.FlightOR) = -1 Then
			'Add the repository to the collection
			qtRepositories.Add GlobalConst.FlightOR, 1
		End If
		Set qtApp = nothing

		'Load All scripts
		Dim key,extensionname, scriptfilesPath, oDict, oFSO
		Set oDict = CreateObject("Scripting.Dictionary")
		'New an instance of the FSOLib class
		Set oFSO = FSOLib
		scriptfilesPath = GlobalConst.ScriptPath
		oFSO.GetFilesRecursively oDict, scriptfilesPath
		For Each key In oDict
	        extensionname = objFSO.GetExtensionName(oDict(key))
			If Ucase(extensionname) = "VBS" Then
					ExecuteFile oDict(key)
			End If
		Next
		Set oDict = nothing
		Set oFSO = nothing
		Set objFSO= nothing
		
	End Function

	
	''' <summary>
	''' Run Flow Engine
	''' </summary>
	''' <remarks></remarks>
	Public Default Function Run()
		
		'Execute PreExecutionSetup data file and only run applications when RunStatus = "Y"
		Call ExecutePreSetupFile
	
	End function
	
	''' <summary>
	''' Execute PreExecutionSetup data file and only run applications when RunStatus = "Y"
	''' </summary>
	''' <remarks></remarks>
	Public Function ExecutePreSetupFile
	
		'Read records from PreExecutionSetup data file
		Dim PreSetupFile : PreSetupFile = GlobalConst.PreExecutionSetupFile
		'Declare variables
		Dim arr_PreSetup2DData, GetUsedRowCount, intStartRow,  intRowsCount
		Dim strRunStatus, strAppName, strEnv, strURL, strRelease, strBuild, strUserid, strPwd
	
		'Load data from PreExecutionSetup file
		oExcelLib.LoadFile PreSetupFile, 1
		'Get 2D array data from sheet
		arr_PreSetup2DData = oExcelLib.Get2DArrayFromSheet
		'Count rows for all sheet
		intRowsCount = oExcelLib.GetUsedRowCount
		
		For intStartRow = 2 to intRowsCount
			'Assign field value from 2D array
			strRunStatus = arr_PreSetup2DData(intStartRow, 1)
			strAppName = arr_PreSetup2DData(intStartRow, 2)
			strEnv = arr_PreSetup2DData(intStartRow, 3)
			strURL = arr_PreSetup2DData(intStartRow, 4)
			strRelease = arr_PreSetup2DData(intStartRow, 5)
			strBuild = arr_PreSetup2DData(intStartRow, 6)
			strUserid = arr_PreSetup2DData(intStartRow, 7)
			strPwd = arr_PreSetup2DData(intStartRow, 8)
			If (strRunStatus <> "") and (strAppName <> "")  Then
				If UCase(strRunStatus) = "Y" Then
					If  (strRelease <> "") and (strBuild <> "") Then
						'Create a custom report with header info
						oReport.CreateCustomReportFile "Demo",strAppName,strRelease,strBuild
						'Execute TestSuite file and only run testsuites when RunStatus = "Y"
						Call ExecuteTestSuite(strAppName, strURL,strUserid, strPwd)
					else
						Reporter.ReportEvent micWarning, "Release/Build", "There is no valid Release/Build for application"
					end if
				end if
			else
				Reporter.ReportEvent micWarning, "Run Status/Application name has not been set", "Run Status for the row '" & intStartRow & "' has not been set or Application name has not been set"
			end if
		next
	
	End Function
	
	''' <summary>
	''' Execute Flight TestSuite file and only run testsuites when RunStatus = "Y"
	''' </summary>
	''' <param name="strAppName" type="string">Correspoonding Application Name</param>
	''' <param name="strURL" type="string">URL/Location name to open the application</param>
	''' <param name="strUserid" type="string">Valid userid/username to logon to application</param>
	''' <param name="strPwd" type="string">Valid password to logon to application</param>
	''' <return></return>
	''' <remarks></remarks>
	Public Function ExecuteTestSuite(ByVal strAppName, ByVal strURL,ByVal strUserid, ByVal strPwd)
	
		Dim strTestSuiteFile : strTestSuiteFile = GlobalConst.TestSuitesPath & strAppName & ".xls"
		Dim strTestCaseFile
		'Declare variables
		Dim arr_TestSuite2DData, GetUsedRowCount, intStartRow,  intRowsCount
		Dim strRunStatus, strTestSuiteName
	
		'Load data from PreExecutionSetup file
		oExcelLib.LoadFile strTestSuiteFile, 1
		'Get 2D array data from sheet
		arr_TestSuite2DData = oExcelLib.Get2DArrayFromSheet
		'Count rows for all sheet
		intRowsCount = oExcelLib.GetUsedRowCount
		
		For intStartRow = 2 to intRowsCount
			'Assign field value from 2D array
			strRunStatus = arr_TestSuite2DData(intStartRow, 1)
			strTestSuiteName = arr_TestSuite2DData(intStartRow, 2)
			If (strRunStatus <> "") and (strTestSuiteName <> "")  Then
				If UCase(strRunStatus) = "Y" Then
					'Create a Test Suite Node
				    oReport.AddTestSuiteNode strTestSuiteName
					strTestCaseFile = GlobalConst.TestCasesPath & strTestSuiteName & ".xls"
					'Execute Flight Reservation.xls test case file and only run testcases when RunStatus = "Y"
					Call RunTestCase(strTestCaseFile,strURL,strUserid, strPwd)
				end if
			else
				Reporter.ReportEvent micWarning, "Run Status/TestSuite name has not been set", "Run Status for the row '" & intStartRow & "' has not been set or TestSuite name has not been set"
			end if
		next
		
	End Function

	''' <summary>
	''' Execute Specific test case file and only run testcases when RunStatus = "Y"
	''' </summary>
	''' <param name="strTestCaseFile" type="string">Location of the correspoonding test case file</param>
	''' <param name="strURL" type="string">URL/Location name to open the application</param>
	''' <param name="strUserid" type="string">Valid userid/username to logon to application</param>
	''' <param name="strPwd" type="string">Valid password to logon to application</param>
	''' <return></return>
	''' <remarks></remarks>
	Public Function RunTestCase(ByVal strTestCaseFile, ByVal strURL, ByVal strUserid, ByVal strPwd)
	
		Dim arr_TestCase2DData, intRowsCount, intColsCount, intStartRow, intParamCol, strParamValues, strParamValue
		Dim strRunStatus, strCaseID, strCaseName, strFuncEntryName, strParam1, strParam2, strParam3
		Dim oParamDict, oConvertedParamDict, arrValues, Val, intCount, arrConvertedValues, i
		Set oParamDict = CreateObject("Scripting.Dictionary")
		Set oConvertedParamDict = CreateObject("Scripting.Dictionary")
		Dim oFuncRef
		'Load data from correspoonding test case  file
		oExcelLib.LoadFile strTestCaseFile, 1
		'Get 2D array data from sheet
		arr_TestCase2DData = oExcelLib.Get2DArrayFromSheet
		'Count rows for all sheet
		intRowsCount = oExcelLib.GetUsedRowCount
		'Count Columns for all sheet
		intColsCount = oExcelLib.GetUsedColumnCount
		
		For intStartRow = 2 to intRowsCount
			oParamDict.RemoveAll
			oConvertedParamDict.RemoveAll
			'Assign field value from 2D array
			strRunStatus = arr_TestCase2DData(intStartRow, 1)
			If UCase(strRunStatus) = "Y" Then
				'Assign values in array to the local variables if run status is 'Yes' otherwise none
				strCaseID = arr_TestCase2DData(intStartRow, 2)
				strCaseName = arr_TestCase2DData(intStartRow, 3)
				strFuncEntryName = arr_TestCase2DData(intStartRow, 4)
				
				'Create a Test Case Node
				oReport.AddTestCaseNode strCaseID, strCaseName
				'Load OR file
				ExecuteFile GlobalConst.GlobalObjectsMapToOR
				
				strParamValue = ""
				For intParamCol = 5 To intColsCount
					strParamValue = arr_TestCase2DData(intStartRow, intParamCol)
					If trim(strParamValue) <> "" Then
						oParamDict.Add intParamCol, trim(strParamValue)
					End if
				Next
				If oParamDict.Count > 0 Then
					arrValues = oParamDict.Items
					For i = 0 To oParamDict.Count - 1
						Val = arrValues(i)
						Select Case Trim(UCase(Val))
							case "URL"
								oConvertedParamDict.add i, strURL
							case "USERID"
								oConvertedParamDict.add i, strUserid
							case "PASSWORD"
								oConvertedParamDict.add i, strPwd
							case Else
								oConvertedParamDict.add i, Val
						End select	
					Next
					
					intCount = oConvertedParamDict.Count
					arrConvertedValues = oConvertedParamDict.Items
					Select Case intCount
						Case 1
							Set oFuncRef = GetRef(strFuncEntryName)
							oFuncRef arrConvertedValues(0)
						Case 2
							Set oFuncRef = GetRef(strFuncEntryName)
							oFuncRef arrConvertedValues(0), arrConvertedValues(1)
						Case 3
							Set oFuncRef = GetRef(strFuncEntryName)
							oFuncRef arrConvertedValues(0), arrConvertedValues(1), arrConvertedValues(2)
						Case 4
							Set oFuncRef = GetRef(strFuncEntryName)
							oFuncRef arrConvertedValues(0), arrConvertedValues(1), arrConvertedValues(2), arrConvertedValues(3)
						Case 5
							Set oFuncRef = GetRef(strFuncEntryName)
							oFuncRef arrConvertedValues(0), arrConvertedValues(1), arrConvertedValues(2), arrConvertedValues(3), arrConvertedValues(4)
						Case 6
							Set oFuncRef = GetRef(strFuncEntryName)
							oFuncRef arrConvertedValues(0), arrConvertedValues(1), arrConvertedValues(2), arrConvertedValues(3), arrConvertedValues(4), arrConvertedValues(5)
						Case 7
							Set oFuncRef = GetRef(strFuncEntryName)
							oFuncRef arrConvertedValues(0), arrConvertedValues(1), arrConvertedValues(2), arrConvertedValues(3), arrConvertedValues(4), arrConvertedValues(5), arrConvertedValues(6)
						Case 8
							Set oFuncRef = GetRef(strFuncEntryName)
							oFuncRef arrConvertedValues(0), arrConvertedValues(1), arrConvertedValues(2), arrConvertedValues(3), arrConvertedValues(4), arrConvertedValues(5), arrConvertedValues(6), arrConvertedValues(7)
						Case else
							msgbox "The number of the  Param is more than 8, please reset select-case statement under function ReadTestCase "
					End Select
				Else
					Set oFuncRef = GetRef(strFuncEntryName)
					oFuncRef
				End if
			End if
		next
	
	End Function
	
End Class

Public Function FlowEngine

	Set FlowEngine = new ClsFlowEngine
	
End function	

'Execute Flow Engine
call FlowEngine.run












