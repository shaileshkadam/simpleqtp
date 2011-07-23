Option Explicit

''' #########################################################
''' <summary>
''' A Library to work with general object
''' </summary>
''' #########################################################


	''' <summary>
    ''' Click a control by its text
    ''' </summary>
    ''' <param name="Obj" type="Object">Parent Window control</param>
    ''' <param name="strText" type="String">Text of the control to be clicked</param>
    ''' <return></return>
    ''' <remarks></remarks>
	Public Function ClickByText(ByVal Obj,ByVal strText)
		
		Dim l, t, r, b
		Dim hwnd
		Dim window_x, window_y, x, y
		Dim Succeeded
		Dim dr
		l = -1
		t= -1
		r = -1
		b = -1
		'Check if the object exist or not
		If Not Obj.Exist Then
			'Reporter.ReportEvent micFail, "Object Does not Exist", "The object does not exist"
			Exit Function
		End If
		hwnd = Obj.GetROProperty("HWND")
		window_x = Obj.GetROProperty("x")
		window_y = Obj.GetROProperty("Y")
		Succeeded = TextUtil.GetTextLocation( strText,hwnd,l,t,r,b)
		If Not Succeeded Then
		MsgBox "Text not found"
		else
		x = window_x +(l+r) / 2
		y = window_y +(t+b) / 2
		Set dr = CreateObject("Mercury.DeviceReplay")
		dr.MouseClick x,y,LEFT_MOUSE_BUTTON
		End If
	
	End Function

    ''' <summary>
    ''' Drag And Drop using DeviceReplay
    ''' </summary>
    ''' <param name="ObjFrom" type="Object">The object to be draged from</param>
    ''' <param name="ObjTo" type="Object">The object to be droped to</param>
    ''' <return></return>
    ''' <remarks></remarks>
    Function DragAndDrop(ByVal ObjFrom, ByVal ObjTo)
    
        'Check if the object exist or not
        If Not (ObjFrom.Exist And ObjTo.Exist) Then
            'Reporter.ReportEvent micFail, "Object Does not Exist", "The object does not exist"
            Exit Function
        End If
     	Dim DC
        Dim ObjFrom_abs_x, ObjFrom_abs_y, ObjFrom_center_x, ObjFrom_center_y
        Dim ObjFrom_width, ObjFrom_height, ObjTo_width, ObjTo_height
        Dim ObjTo_abs_x, ObjTo_abs_y, ObjTo_center_x, ObjTo_center_y     
        'Get the position of the object on the screen
        ObjFrom_abs_x = ObjFrom.GetROProperty("abs_x")
        ObjFrom_abs_y = ObjFrom.GetROProperty("abs_y")
        ObjTo_abs_x = ObjTo.GetROProperty("abs_x")
        ObjTo_abs_y = ObjTo.GetROProperty("abs_y")
        If ObjFrom_abs_x < 0 or ObjFrom_abs_y < 0 or ObjFrom_abs_x = "" or ObjFrom_abs_y = "" Or ObjTo_abs_x < 0 or ObjTo_abs_y < 0 or ObjTo_abs_x = "" or ObjTo_abs_y = "" Then
            'Reporter.ReportEvent micFail, "Object is not Visible", "The object is not visible "
            Exit Function
        End if
     
        ObjFrom_width = ObjFrom.GetROProperty("width")
        ObjFrom_height = ObjFrom.GetROProperty("height")
        ObjTo_width = ObjTo.GetROProperty("width")
        ObjTo_height = ObjTo.GetROProperty("height")
        
        ObjFrom_center_x = ObjFrom_abs_x + ObjFrom_width\2
        ObjFrom_center_y = ObjFrom_abs_y + ObjFrom_height\2
        ObjTo_center_x = ObjTo_abs_x + ObjTo_width\2
        ObjTo_center_y = ObjTo_abs_y + ObjTo_height\2
             
        Set DC = CreateObject("Mercury.DeviceReplay")
     
        'Drag from the center of the ObjFrom and drop to center of the ObjTo
        DC.DragAndDrop ObjFrom_center_x,ObjFrom_center_y,ObjTo_center_x,ObjTo_center_y,micLeftBtn
        
    End Function 



	''' <summary>
    ''' Clicks on the center of an Object using DeviceReplay
    ''' </summary>
    ''' <param name="Obj" type="Object">The object to be clicked on</param>
    ''' <return></return>
    ''' <remarks></remarks>
	Function AsyncClick(ByVal Obj)
	
		'Check if the object exist or not
		If Not Obj.Exist Then
			'Reporter.ReportEvent micFail, "Object Does not Exist", "The object does not exist"
			Exit Function
		End If
	 
		Dim x, y, width, height, DC
	 
		'Get the position of the object on the screen
		x = obj.GetROProperty("abs_x")
		y = obj.GetROProperty("abs_y")
	 
		If x < 0 or y < 0 or x = "" or y = "" Then
			'Reporter.ReportEvent micFail, "Object is not Visible", "The object is not visible "
			Exit Function
		End if
	 
		width = obj.GetROProperty("width")
		height = obj.GetROProperty("height")
	 
		x = x + width\2
		y = y + height\2
	 
		Set DC = CreateObject("Mercury.DeviceReplay")
	 
		'Click on the Middle of the button
		DC.MouseClick x,y,micLeftBtn
		
	End Function

	''' <summary>
    ''' Waits until an object is loaded or the specified timeout expires.
    ''' </summary>
    ''' <param name="obj" type="Object">Any Object</param>
	''' <param name="intTimeoutMSec" type="int">specified timeout</param>
	''' <param name="interval" type="int">Check whether the obj exists or not every specified interval</param>
    ''' <return>True/False</return>
    ''' <remarks></remarks>
	Public Function Exists(ByRef obj, ByVal intTimeoutMSec, ByVal interval)
	
	    Dim objTimer
	 
	    If Not IsNumeric(intTimeoutMSec) Then
	        intTimeoutMSec = Environment("DEFAULT_TIMEOUT_MSEC")
	    End If
	 
	    Set objTimer = MercuryTimers.Timer("ObjectExist")
	 
	    objTimer.Start
	 
	    Do
	        Exists = obj.Exist(0)
	        If Exists Then
	            objTimer.Stop
	            Exit Do
	        End If
	        Wait interval
	    Loop Until objTimer.ElapsedTime > intTimeoutMSec
	    objTimer.Stop
	    
	End Function

	''' <summary>
    ''' Highlights all the objects
    ''' </summary>
    ''' <param name="obj" type="Object">Any Object</param>
    ''' <return></return>
    ''' <remarks></remarks>
	Public Function HighlightAll(ByVal obj)
	    
	    Dim Parent, Desc, Props, PropsCount, MaxIndex, i, Objs 
	    If IsEmpty(obj.GetTOProperty("parent")) Then 
	         Set Parent = Desktop 
	    Else 
	         Set Parent = obj.GetTOProperty("parent") 
	    End If 
	    Set Desc = Description.Create 
	    Set Props = obj.GetTOProperties 
	    PropsCount = Props.Count - 1 
	    For i = 0 to PropsCount 
	        Desc(Props(i).Name).Value = Props(i).Value 
	    Next 
	    Set Objs = Parent.ChildObjects(Desc) 
	    MaxIndex = Objs.Count - 1 
	    For i = 0 to MaxIndex 
	         Objs.Item(i).Highlight 
	    Next
	    
	End Function
	
	''' <summary>
    ''' Retrieve object count (visible + hidden)
    ''' </summary>
    ''' <param name="BaseObject" type="Object">Object containing the ClassName objects</param>
    ''' <param name="strClassName" type="string">MicClass for which the count is being retrieved for</param>
    ''' <return>Integer</return>
    ''' <remarks></remarks>
	Public Function GetObjectCount(ByVal BaseObject,ByVal strClassName ) ' As Integer
	
		BaseObject.Init
	    If Not BaseObject.Exist( 0 ) Then
	      	msgbox("BaseObject was not found.")
	        GetClassCount = -1
	        Exit Function
	    End If
	    Dim oDesc, intCount
	    intCount = 0
	    Set oDesc = Description.Create
	    oDesc( "micclass" ).Value = ClassName
	    intCount = BaseObject.ChildObjects( oDesc ).Count
	    GetClassCount = intCount
	        
	End Function
	
	''' <summary>
    ''' Retrieve visible objects count
    ''' </summary>
    ''' <param name="BaseObject" type="Object">Object containing the ClassName objects</param>
    ''' <param name="strClassName" type="string">MicClass for which the count is being retrieved for</param>
    ''' <return>Integer</return>
    ''' <remarks></remarks>
	Public Function GetVisibleObjectCount(ByVal BaseObject,ByVal ClassName ) ' As Integer
	
		BaseObject.Init
	    If Not BaseObject.Exist( 0 ) Then
	        msgbox("BaseObject was not found.")
	        GetClassCount = -1
	        Exit Function
	    End If
	    Dim oDesc, intCount
	    intCount = 0
	    Set oDesc = Description.Create
	    oDesc( "micclass" ).Value = ClassName
	    intCount = BaseObject.ChildObjects(oDesc).Count
	    oDesc( "x" ).Value = 0
	    intCount = intCount - BaseObject.ChildObjects(oDesc).Count
	    GetVisibleObjectCount = intCount
	        
	End Function

