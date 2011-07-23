Option Explicit

''' <summary>
''' Check that the current GUI context is loaded
''' </summary>
''' <remarks></remarks>
Class ClsGUILayerContext
	
	''' <summary>
    ''' Check that the current GUI context is loaded by specific time
    ''' </summary>
    ''' <param name="oDictGUIObjects" type="Dictionary">ChildObjects Dictionary</param>
    ''' <param name="Interval" type="int">Check once every specific Interval second</param>
    ''' <param name="TimeOut" type="int">Max timeout</param>
    ''' <return>True/False</return>
    ''' <remarks></remarks>
	Public Function IsLoaded(ByVal oDictGUIObjects, ByVal Interval, ByVal TimeOut)
		
		Dim Starting, Ending, t
		Starting = Now
		Ending = DateAdd("s",TimeOut,Starting)
		IsLoaded = False
		Do
			t = DateDiff("s",Now,Ending)
			If IsContextLoaded(oDictGUIObjects) Then
				IsLoaded = true
				Exit Do
			End If
			wait Interval
		Loop Until t <= 0
		
	End function

	
	''' <summary>
    ''' Check that the current GUI context is loaded
    ''' </summary>
    ''' <param name="oDictGUIObjects" type="Dictionary">ChildObjects Dictionary</param>
    ''' <return>True/False</return>
    ''' <remarks></remarks>
	Public Function IsContextLoaded(ByVal oDictGUIObjects)
	
		Dim keys,items,i,instatus,strDetails,strAdditionalRemarks
		keys = oDictGUIObjects.Keys
		items = oDictGUIObjects.Items
		IsContextLoaded = True
		'Iterates through the oDictGUIObjects items and executes the exist method with 0 as parameter
		For i = 0 To oDictGUIObjects.count - 1 
			IsContextLoaded = IsContextLoaded And items(i).exist(0)
			strDetails = strDetails & vbNewLine & "Object #" & i+1 & keys(i) & " was"	
			If IsContextLoaded Then
				instatus = micPass
				strDetails = strDetails & ""
				strAdditionalRemarks = ""
			Else
				instatus = micWarning
				strDetails = strDetails & " not"
				strAdditionalRemarks = "Please check the object properies"
			End If
			strDetails = strDetails & " found." & strAdditionalRemarks	
		Next
		
		Reporter.ReportEvent instatus, "IsContextLoaded", strDetails
		
	End Function	


End Class

Public Function GUILayerContext()

	Set GUILayerContext = New ClsGUILayerContext
	
End function	