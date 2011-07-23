Option Explicit

''' #####################################################
''' <summary>
''' Private Class ReportLib to create a custom report
''' </summary>
''' <remarks></remarks>
''' #####################################################

Class ClsReportLib

'Private Variables
	
	''' <summary>
    ''' FileSystemObject instance created in Class_Initialize
    ''' </summary>
    ''' <remarks></remarks>
	Private oFSO
	
	''' <summary>
    ''' XML Document instance created in Class_Initialize
    ''' </summary>
    ''' <remarks></remarks>
	Private objXMLCustomReport
	
	''' <summary>
    ''' XML Root node instance created in Class_Initialize
    ''' </summary>
    ''' <remarks></remarks>
	Private objXMLroot
	
	''' <summary>
    ''' XML TestSuite node instance
    ''' </summary>
    ''' <remarks></remarks>
	Private objXMLTestSuite
	
	''' <summary>
    ''' XML TestCase node instance
    ''' </summary>
    ''' <remarks></remarks>
    Private objXMLTestCase
    
	''' <summary>
    ''' XML Step node instance
    ''' </summary>
    ''' <remarks></remarks>
   	Private objXMLStep
   	
   	''' <summary>
    ''' Application Name
    ''' </summary>
    ''' <remarks></remarks>
	Private sApplicationName
	
   	''' <summary>
    ''' Release
    ''' </summary>
    ''' <remarks></remarks>
   	Private sRelease
   	
   	''' <summary>
    ''' Build
    ''' </summary>
    ''' <remarks></remarks>
   	Private sBuild
   	
	''' <summary>
    ''' Custom report file path and name
    ''' </summary>
    ''' <remarks></remarks>
   	Private CustomReportFile

	''' <summary>
    ''' Custom report file name
    ''' </summary>
    ''' <remarks></remarks>
	Private CustomReportFileName
	
	''' <summary>
    ''' Path of Result Template Folder
    ''' </summary>
    ''' <remarks></remarks>
   	Private ResultTemplatePath
	
	''' <summary>
    ''' Path of Screenshot Folder
    ''' </summary>
    ''' <remarks></remarks>
	Private ResultScreenshotPath
	
	''' <summary>
    ''' Start Time
    ''' </summary>
    ''' <remarks></remarks>
	Private StartTime
	
	''' <summary>
    ''' End Time
    ''' </summary>
    ''' <remarks></remarks>
	Private EndTime
	
	
'Public Properties

	Public Property Get ApplicationName
		ApplicationName = sApplicationName
	End Property
	
	Public Property Let ApplicationName(ByVal val)
		sApplicationName = val
	End Property   

	Public Property Get Release
		Release  = sRelease
	End Property
	
	Public Property Let Release(ByVal val)
		sRelease = val
	End Property
	
	Public Property Get Build
		Build  = sBuild
	End Property
	
	Public Property Let Build(ByVal val)
		sBuild = val
	End Property	 
	
	''' <summary>
    ''' Class Initialization procedure. Creates XML Document and root node
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Initialize()

		Dim XMLFileHeader,CurrentTestPath
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		XMLFileHeader = "<?xml version='1.0' encoding='UTF-8'?><?xml-stylesheet href='Report.xsl' type='text/xsl'?><Report></Report>"
		Set objXMLCustomReport = XMLUtil.CreateXML() 
		objXMLCustomReport.Load XMLFileHeader
		Set objXMLroot = objXMLCustomReport.GetRootElement()
'		CurrentTestPath = PathFinder.CurrentTestPath
'		ResultTemplatePath = CurrentTestPath & "\Reports\"
		ResultTemplatePath = GlobalConst.ReportPath
		
	End Sub

	''' <summary>
    ''' Class Termination procedure
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Terminate()

		Set oFSO = Nothing
		Set objXMLCustomReport = Nothing
		Set objXMLroot = nothing
		Set objXMLTestSuite = nothing
		Set objXMLTestCase = nothing
		Set objXMLStep = nothing
					  
	End Sub
	
	''' <summary>
    ''' Save XML document with specific path and name
    ''' </summary>
    ''' <param name="Reportfile" type="string">Path and name to the XML document</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
   	Public	Function SaveReport(ByVal Reportfile)

		objXMLCustomReport.SaveFile Reportfile
		Call UpdateTimestamp
		Call ApplyXSL
	
	End Function
	
	''' <summary>
    ''' Create XML document with header info
    ''' </summary>
    ''' <param name="strReportHeader" type="string">Header e.g "SECURITY FINANCE BUSINESS TECHNOLOGY"</param>
    ''' <param name="strApplicationName" type="string">Application Name e.g "SGOWeb"</param>
    ''' <param name="strRelease" type="string">Release e.g "4.8.0"</param>
    ''' <param name="strBuild" type="string">Build e.g "1.0"</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
   	Public Function CreateCustomReportFile(ByVal strReportHeader, ByVal strApplicationName, ByVal strRelease, ByVal strBuild)

		Dim CurrentPath, strTimeStamp
		'Start Timer
		StartTime = Now
		Me.ApplicationName = strApplicationName
		Me.Release = strRelease
		Me.Build = strBuild
		objXMLroot.AddAttribute "Header",CStr(strReportHeader)
		objXMLroot.AddAttribute "ApplicationName",CStr(Me.ApplicationName)
		objXMLroot.AddAttribute "Release",CStr(Me.Release)
		objXMLroot.AddAttribute "Build",CStr(Me.Build)
		objXMLroot.AddAttribute "StartTime", cstr(Now)
		objXMLroot.AddAttribute "EndTime", cstr(Now)
		objXMLroot.AddAttribute "ExecuteHourTime", "0"
		objXMLroot.AddAttribute "ExecuteMinuteTime", "0"
		strTimeStamp = Cstr(Now)
		strTimeStamp = Replace(strTimeStamp, "/", "_")
		strTimeStamp = Replace(strTimeStamp, ":", "_")
		CustomReportFileName = Me.ApplicationName & " " & "Release " & Replace(Me.Release, ".", "_") & " Build " & Replace(Me.Build, ".", "_") & " " & strTimeStamp
		CustomReportFile = ResultTemplatePath & CustomReportFileName  & ".xml"
	    Call SaveReport(CustomReportFile)
		ResultScreenshotPath = ResultTemplatePath & "_Screenshots\" & CustomReportFileName & "\"

	End Function
	
	''' <summary>
    ''' Create TestSuite Node into XML document
    ''' </summary>
    ''' <param name="strDesc" type="string">Description of the TestSuite</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
   	Public Function AddTestSuiteNode(ByVal strDesc)
	
		objXMLroot.AddChildElementByName "TestSuite", ""
		Set objXMLTestSuite = objXMLroot.ChildElements().Item(objXMLroot.ChildElements().count())
		objXMLTestSuite.AddAttribute "Desc" , strDesc
		Call SaveReport(CustomReportFile)
	
	End Function
	
	''' <summary>
    ''' Create TestCase Node into XML document
    ''' </summary>
    ''' <param name="strTestCase_ID" type="string">ID of the TestCase</param>
    ''' <param name="strDesc" type="string">Description of the TestCase</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function AddTestCaseNode(ByVal strTestCase_ID, ByVal strDesc)
	
		objXMLTestSuite.AddChildElementByName "TestCase", ""
		Set objXMLTestCase = objXMLTestSuite.ChildElements().Item(objXMLTestSuite.ChildElements().count())
		objXMLTestCase.AddAttribute "ID" , strTestCase_ID
		objXMLTestCase.AddAttribute "Desc" , strDesc
		Call SaveReport(CustomReportFile)
	
	End Function
	
	''' <summary>
    ''' Capture a screenshot
    ''' </summary>
    ''' <param name="strScreenShotName" type="string">ScreenShotName e.g "test.bmp"</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function CaptureScreenshot(ByVal strScreenShotName)
	
		Dim strScreenShot
		CreateNestedDirs ResultScreenshotPath
		strScreenShot = ResultScreenshotPath & strScreenShotName
		' just capture
		Desktop.CaptureBitmap strScreenShot,True
	
	End Function

	''' <summary>
    ''' Creates multiple level of folders
    ''' By default VBScript can only create one level of folders at a time
    ''' </summary>
    ''' <param name="MyDirName" type="string">folder(s) to be created, single or multi level, absolute or relative, e.g "d:\folder\subfolder" </param>
    ''' <returns></returns>
    ''' <remarks></remarks>	
	Public Function CreateNestedDirs(ByVal MyDirName)
	
	    Dim arrDirs, i, idxFirst, strDir, strDirBuild
	
	    ' Convert relative to absolute path
	    strDir = oFSO.GetAbsolutePathName( MyDirName )
	
	    ' Split a multi level path in its "components"
	    arrDirs = Split( strDir, "\" )
	
	    ' Check if the absolute path is UNC or not
	    If Left( strDir, 2 ) = "\\" Then
	        strDirBuild = "\\" & arrDirs(2) & "\" & arrDirs(3) & "\"
	        idxFirst    = 4
	    Else
	        strDirBuild = arrDirs(0) & "\"
	        idxFirst    = 1
	    End If
	
	    ' Check each (sub)folder and create it if it doesn't exist
	    For i = idxFirst to Ubound( arrDirs )
	        strDirBuild = oFSO.BuildPath( strDirBuild, arrDirs(i) )
	        If Not oFSO.FolderExists( strDirBuild ) Then 
	            oFSO.CreateFolder(strDirBuild)
	        End if
	    Next

	End Function
	
	''' <summary>
    ''' Create a passed Step into XML document and report in QTP
    ''' </summary>
    ''' <param name="arr" type="array">a array with combination of the Step Description/expect result/actual result/link to file</param>
    ''' <param name="bCaptureScreenshot" type="Bool">whether capture screenshot or not</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function ReportPass(ByVal arr, ByVal bCaptureScreenshot)

		Call AddStepNode(arr, "1", bCaptureScreenshot)
		Call SaveReport(CustomReportFile)
		Select Case UBound(arr) 
			Case 0 
				'report in QTP
				Reporter.ReportEvent micPass,arr(0),""
			Case 1
				'report in QTP
				Reporter.ReportEvent micPass,arr(0),arr(1)
			Case Else
				'report in QTP
				Reporter.ReportEvent micPass,arr(0),arr(2)	
		End select

	End Function


	''' <summary>
    ''' Create a failed Step into XML document and report in QTP
    ''' </summary>
    ''' <param name="arr" type="array">a array with combination of the Step Description/expect result/actual result/link to file</param>
    ''' <param name="bCaptureScreenshot" type="Bool">whether capture screenshot or not</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
   	Public Function ReportFail(ByVal arr, ByVal bCaptureScreenshot)

		Call AddStepNode(arr, "2", bCaptureScreenshot)		
		Call SaveReport(CustomReportFile)
		Select Case UBound(arr) 
			Case 0 
				'report in QTP
				Reporter.ReportEvent micFail,arr(0),""
			Case 1
				'report in QTP
				Reporter.ReportEvent micFail,arr(0),arr(1)
			Case Else
				'report in QTP
				Reporter.ReportEvent micFail,arr(0),arr(1) & " but " & arr(2)	
		End select

	End Function
	
	''' <summary>
    ''' Create Step Node into XML document
    ''' </summary>
    ''' <param name="arr" type="array">a array with combination of the Step Description/expect result/actual result/link to file</param>
    ''' <param name="strStatus" type="string">"1" represents Pass, "2" represents Fail</param>
    ''' <param name="bCaptureScreenshot" type="Bool">whether capture screenshot or not</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function AddStepNode(ByVal arr, ByVal strStatus, ByVal bCaptureScreenshot)
	
		Dim v1, v2, v3, v4, strExtensionName
		Select Case UBound(arr)
			Case 0
				v1 = arr(0)
				Call AddStepNodeWithoutResult(v1, strStatus, bCaptureScreenshot)
			Case 1
				v1 = arr(0)
				v2 = arr(1)
				Call AddStepNodeWithDetail(v1, v2, strStatus, bCaptureScreenshot)
			Case 2
				v1 = arr(0)
				v2 = arr(1)
				v3 = arr(2)
				Call AddStepNodeWithResult(v1, v2, v3, strStatus, bCaptureScreenshot)
			Case 3
				v1 = arr(0)
				v2 = arr(1)
				v3 = arr(2)
				v4 = arr(3)
				Call AddStepNodeWithResultAndLinkTOFile(v1, v2, v3, v4, strStatus, bCaptureScreenshot)
			Case Else
				Exit Function
		End Select
	
	End Function
	
	''' <summary>
    ''' Create Step Node without expect/actual result
    ''' </summary>
    ''' <param name="strDesc" type="string">Step Description</param>
    ''' <param name="strStatus" type="string">"1" represents Pass, "2" represents Fail</param>
    ''' <param name="bCaptureScreenshot" type="Bool">whether capture screenshot or not</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function AddStepNodeWithoutResult(ByVal strDesc, ByVal strStatus, ByVal bCaptureScreenshot)
		
		Dim strTimeStamp, strScreenShotName
		objXMLTestCase.AddChildElementByName "Step", strDesc
		Set objXMLStep = objXMLTestCase.ChildElements().Item(objXMLTestCase.ChildElements().count())
		objXMLStep.AddAttribute "Status" , strStatus
		If bCaptureScreenshot Then
			strTimeStamp = Cstr(Now)
			strTimeStamp = Replace(strTimeStamp, "/", "_")
			strTimeStamp = Replace(strTimeStamp, ":", "_")
			strScreenShotName = Me.ApplicationName & " " & Replace(Me.Release, ".", "_") & " " & strTimeStamp & ".bmp"
			Call CaptureScreenshot(strScreenShotName)
			objXMLStep.AddAttribute "ScreenShotPath" , ResultScreenshotPath & strScreenShotName
		End if
	
	End Function
	
	''' <summary>
    ''' Create Step Node with detail
    ''' </summary>
    ''' <param name="strDesc" type="string">Step Name Description</param>
    ''' <param name="strDetail" type="string">Detail Description</param>
    ''' <param name="strStatus" type="string">"1" represents Pass, "2" represents Fail</param>
    ''' <param name="bCaptureScreenshot" type="Bool">whether capture screenshot or not</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function AddStepNodeWithDetail(ByVal strDesc, ByVal strDetail, ByVal strStatus, ByVal bCaptureScreenshot)
		
		Dim strTimeStamp, strScreenShotName
		objXMLTestCase.AddChildElementByName "Step", strDesc
		Set objXMLStep = objXMLTestCase.ChildElements().Item(objXMLTestCase.ChildElements().count())
		objXMLStep.AddAttribute "Status" , strStatus
		objXMLStep.AddAttribute "Detail" , strDetail
		If bCaptureScreenshot Then
			strTimeStamp = Cstr(Now)
			strTimeStamp = Replace(strTimeStamp, "/", "_")
			strTimeStamp = Replace(strTimeStamp, ":", "_")
			strScreenShotName = Me.ApplicationName & " " & Replace(Me.Release, ".", "_") & " " & strTimeStamp & ".bmp"
			Call CaptureScreenshot(strScreenShotName)
			objXMLStep.AddAttribute "ScreenShotPath" , ResultScreenshotPath & strScreenShotName
		End if
	
	End Function

	''' <summary>
    ''' Create Step Node with expect/actual result
    ''' </summary>
    ''' <param name="strDesc" type="string">Step Description</param>
    ''' <param name="strExpectedResult" type="string">Expected Result</param>
    ''' <param name="strActualResult" type="string">Actual Result</param>
    ''' <param name="strStatus" type="string">"1" represents Pass, "2" represents Fail</param>
    ''' <param name="bCaptureScreenshot" type="Bool">whether capture screenshot or not</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
   	Public Function AddStepNodeWithResult(ByVal strDesc, ByVal strExpectedResult, ByVal strActualResult, ByVal strStatus, ByVal bCaptureScreenshot)
	
		Dim strTimeStamp, strScreenShotName
		objXMLTestCase.AddChildElementByName "Step", strDesc
		Set objXMLStep = objXMLTestCase.ChildElements().Item(objXMLTestCase.ChildElements().count())
		objXMLStep.AddAttribute "Status" , strStatus
		objXMLStep.AddAttribute "ExpectedResult" , strExpectedResult
		objXMLStep.AddAttribute "ActualResult" , strActualResult
		If bCaptureScreenshot Then
			strTimeStamp = Cstr(Now)
			strTimeStamp = Replace(strTimeStamp, "/", "_")
			strTimeStamp = Replace(strTimeStamp, ":", "_")
			strScreenShotName = Me.ApplicationName & " " & Replace(Me.Release, ".", "_") & " " & strTimeStamp & ".bmp"
			Call CaptureScreenshot(strScreenShotName)
			objXMLStep.AddAttribute "ScreenShotPath" , ResultScreenshotPath & strScreenShotName
		End if
	
	End Function
	
	''' <summary>
    ''' Create Step Node with Link to file
    ''' </summary>
    ''' <param name="strDesc" type="string">Step Description</param>
    ''' <param name="strExpectedResult" type="string">Expected Result</param>
    ''' <param name="strActualResult" type="string">Actual Result</param>
    ''' <param name="strFilepath" type="string">Link to file</param>
    ''' <param name="strStatus" type="string">"1" represents Pass, "2" represents Fail</param>
    ''' <param name="bCaptureScreenshot" type="Bool">whether capture screenshot or not</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function AddStepNodeWithResultAndLinkTOFile(ByVal strDesc, ByVal strExpectedResult, ByVal strActualResult, ByVal strFilepath, ByVal strStatus, ByVal bCaptureScreenshot)
	
		Dim strTimeStamp, strScreenShotName
		objXMLTestCase.AddChildElementByName "Step", strDesc
		Set objXMLStep = objXMLTestCase.ChildElements().Item(objXMLTestCase.ChildElements().count())
		objXMLStep.AddAttribute "Status" , strStatus
		objXMLStep.AddAttribute "ExpectedResult" , strExpectedResult
		objXMLStep.AddAttribute "ActualResult" , strActualResult
		objXMLStep.AddAttribute "Filepath" , strFilepath
		If bCaptureScreenshot Then
			strTimeStamp = Cstr(Now)
			strTimeStamp = Replace(strTimeStamp, "/", "_")
			strTimeStamp = Replace(strTimeStamp, ":", "_")
			strScreenShotName = Me.ApplicationName & " " & Replace(Me.Release, ".", "_") & " " & strTimeStamp & ".bmp"
			Call CaptureScreenshot(strScreenShotName)
			objXMLStep.AddAttribute "ScreenShotPath" , ResultScreenshotPath & strScreenShotName
		End if
	
	End Function

	''' <summary>
    ''' Update end time and execute time
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function UpdateTimestamp()

		Dim strTimeStamp, intHourTotalTime, intMinuteTotalTime,IntSecondTotalTime
		strTimeStamp = Cstr(Now)
		objXMLroot.RemoveAttribute "EndTime"
		objXMLroot.AddAttribute "EndTime", strTimeStamp
		objXMLroot.RemoveAttribute "ExecuteMinuteTime"
		objXMLroot.RemoveAttribute "ExecuteHourTime"
		'End Timer
		EndTime = Now
		intHourTotalTime = DateDiff("h",StartTime,EndTime)
		intMinuteTotalTime = DateDiff("n",StartTime,EndTime)
		IntSecondTotalTime = DateDiff("s",StartTime,EndTime)
		objXMLroot.AddAttribute "ExecuteHourTime", cstr(intHourTotalTime)
		objXMLroot.AddAttribute "ExecuteMinuteTime", cstr(intMinuteTotalTime+Round((IntSecondTotalTime/60),2))
	
	End Function
	
	''' <summary>
    ''' Transform XML to HTML 
    ''' </summary>  
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function ApplyXSL()
		
		Dim sXMLLib, xmlDoc, xslDoc, outputText, outFile
		sXMLLib = "MSXML.DOMDocument"
		Set xmlDoc = CreateObject(sXMLLib)
		Set xslDoc = CreateObject(sXMLLib)
		xmlDoc.async = False
		xslDoc.async = False
		xslDoc.load ResultTemplatePath & "Report.xsl"
		xmlDoc.load CustomReportFile
		outputText = xmlDoc.transformNode(xslDoc.documentElement)
		'outputText=replace(outputText,"UTF-16","gb2312")
		Set outFile = oFSO.CreateTextFile(ResultTemplatePath & CustomReportFileName & ".html",True)
		outFile.Write outputText
		outFile.Close
		Set outFile = Nothing
		Set xmlDoc = Nothing
		Set xslDoc = Nothing
		
	End Function

End Class

Public Function ReportLib()
	Set	ReportLib = new ClsReportLib
End Function
