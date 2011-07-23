Option Explicit

''' ##################################################################
''' <summary>
''' Schedule a task using windows inner Scheduled Task to run one QTP Test automatically
''' </summary>
''' <remarks>
''' Please make sure putting this VBS file into same directory with AutoRun.VBS
''' Double-click this file will schedule one QTP Test automatically
''' </remarks>
''' ##################################################################

Class ClsScheduleAOM
	
	''' <summary>
    ''' Disable windows 'ScreenSaver'
    ''' </summary>
    ''' <remarks></remarks>
	Public Function DisableScreenSaver()
	
		Dim WshShell  
		Set WshShell = WScript.CreateObject("WScript.Shell")  
		WshShell.RegWrite "HKCU\Control Panel\Desktop\ScreenSaveActive",0,"REG_SZ" 
		Set WshShell = Nothing
	
	End Function
	
	''' <summary>
    ''' Delete specific task using windows inner Scheduled Task
    ''' </summary>
    ''' <remarks></remarks>
	Public Function DeleteTask(ByVal sTaskName)
	
		Dim WshShell, DeleteParemeters
		Set WshShell = CreateObject("WScript.Shell")
		DeleteParemeters = "/Delete /tn " & Chr(34) & sTaskName & Chr(34) & " /f"
		WshShell.Run "schtasks.exe " & DeleteParemeters
		
	End Function

	''' <summary>
    ''' Add specific task using windows inner Scheduled Task
    ''' </summary>
    ''' <param name="sTaskName" type="string">Specifies a name for the task</param>
    ''' <param name="sStartTime" type="string">Specifies the time of day that the task starts in HH:MM:SS 24-hour format</param>
    ''' <param name="sSchedule" type="string">Specifies the schedule type. Valid values are MINUTE, HOURLY, DAILY, WEEKLY, MONTHLY, ONCE, ONSTART, ONLOGON, ONIDLE</param>
    ''' <remarks>Detail for schtasks.exe please refer to http://www.microsoft.com/resources/documentation/windows/xp/all/proddocs/en-us/schtasks.mspx?mfr=true</remarks>
	Public Function AddTask(ByVal sTaskName, ByVal sStartTime, ByVal sSchedule)
	
		Dim WshShell, sScriptLocation, AddParemeters
		Set WshShell = CreateObject("WScript.Shell")
		Call DisableScreenSaver
		Call DeleteTask(sTaskName)
		'A Scheduled Task Does Not Run When You Use Schtasks.exe to Create It 
		'and When the Path of the Scheduled Task Contains a Space
		'Please refer to http://support.microsoft.com/kb/823093/en-us
		sScriptLocation = "\" & Chr(34) & WshShell.CurrentDirectory & "\AutoRun.vbs" & "\" & Chr(34)
		AddParemeters = "/create /ru system /tn " & Chr(34) & sTaskName & Chr(34) & " /tr " & Chr(34) & sScriptLocation & Chr(34) &  " /st " & sStartTime & " /sc " & sSchedule 
		WshShell.Run "schtasks.exe " & AddParemeters
		
	End Function

End Class

Public Function ScheduleAOM()
	
	Set ScheduleAOM = New ClsScheduleAOM

End Function

''' <summary>
''' Schedule to run QTP Test automaticcally using windows inner Scheduled Task
''' </summary>
''' <remarks></remarks>
Dim sTaskName, sTime, sSchedule
sTaskName = "Schedule running QTP"
sTime = "21:41:00"
sSchedule = "DAILY"
ScheduleAOM.AddTask sTaskName, sTime, sSchedule