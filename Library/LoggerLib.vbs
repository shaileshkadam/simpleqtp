Option Explicit

''' #########################################################
''' <summary>
''' A logger Library to record debug infos for app
''' </summary>
''' <remarks></remarks>	
''' #########################################################

Class ClsLoggerLib
	
	''' <summary>
    ''' An instance of the Logger class
    ''' </summary>
    ''' <remarks></remarks>
	Private oLogFSO
	
	''' <summary>
    ''' Setting whether to enable/disable logging globally
    ''' </summary>
    ''' <remarks></remarks>
	Private bEnableLogging
	
	''' <summary>
    ''' Setting whether to time stamp Each message with the current date and time
    ''' </summary>
    ''' <remarks></remarks>
	Private bIncludeDateStamp
	
	''' <summary>
    ''' Setting whether to create incremental log files
    ''' </summary>
    ''' <remarks></remarks>
	Private bPrependDateStampInLogFileName
	
	''' <summary>
    ''' Specify the log file location here.
    ''' </summary>
    ''' <remarks> 
    ''' If you would like to log to the same location as the currently running script, set this value to "relative"
    ''' Or sLogFileLocation = "C:\LogFiles\"
	''' </remarks>
	Private sLogFileLocation
	
	''' <summary>
    ''' Specify the log file name here.
    ''' </summary>
    ''' <remarks></remarks>
	Private sLogFileName
	
	''' <summary>
    ''' Setting whether to overwrite/append
    ''' </summary>
    ''' <remarks>"overwrite"/"append"</remarks>
	Private sOverWriteORAppend
	
	''' <summary>
    ''' Setting the maximum number of lines,
    ''' Setting this to a value of 0 will disable this function.
    ''' </summary>
    ''' <remarks></remarks>
	Private vLogMaximumLines
	
	''' <summary>
    ''' Setting the maximum number total size of the log file
    ''' Setting this to a value of 0 will disable this function.
    ''' </summary>
    ''' <remarks></remarks>
	Private vLogMaximumSize
	
	Public Property Get EnableLogging()
		EnableLogging = bEnableLogging
	End Property
	
	Public Property Let EnableLogging(ByVal val)
		bEnableLogging = val
	End Property
	
	Public Property Get IncludeDateStamp()
		IncludeDateStamp = bIncludeDateStamp
	End Property
	
	Public Property Let IncludeDateStamp(ByVal val)
		bIncludeDateStamp = val
	End Property
	
	Public Property Get PrependDateStampInLogFileName()
		PrependDateStampInLogFileName = bPrependDateStampInLogFileName
	End Property
	
	Public Property Let PrependDateStampInLogFileName(ByVal val)
		bPrependDateStampInLogFileName = val
	End Property
	
	Public Property Get LogFileLocation()
		LogFileLocation = sLogFileLocation
	End Property
	
	Public Property Let LogFileLocation(ByVal val)
		sLogFileLocation = val
	End Property
	
	Public Property Get LogFileName()
		LogFileName = sLogFileName
	End Property
	
	Public Property Let LogFileName(ByVal val)
		sLogFileName = val
	End Property
	
	Public Property Get OverWriteORAppend()
		OverWriteORAppend = sOverWriteORAppend
	End Property
	
	Public Property Let OverWriteORAppend(ByVal val)
		sOverWriteORAppend = val
	End Property
	
	Public Property Get LogMaximumLines()
		LogMaximumLines = vLogMaximumLines
	End Property
	
	Public Property Let LogMaximumLines(ByVal val)
		vLogMaximumLines = val
	End Property
	
	Public Property Get LogMaximumSize()
		LogMaximumSize = vLogMaximumSize
	End Property
	
	Public Property Let LogMaximumSize(ByVal val)
		vLogMaximumSize = val
	End Property
	
	''' <summary>
    ''' Class Initialization procedure
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Initialize()
	
		 Set oLogFSO = CreateObject("Scripting.FileSystemObject")
		 bEnableLogging = True
		 bIncludeDateStamp = True
		 bPrependDateStampInLogFileName = true
		 sLogFileLocation = "relative"
		 sLogFileName = "DebugInfo.txt"
		 sOverWriteORAppend = "append"
		 vLogMaximumLines = 0
		 vLogMaximumSize = 0
		 
	End Sub
	
	''' <summary>
    ''' Class Termination procedure
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Terminate()
	
		Set oLogFSO = Nothing
		
	End Sub

	''' <summary>
    ''' Log message to log file
    ''' </summary>
    ''' <param name="Message" type="string">the message to be loged</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function LogToFile(ByVal Message)
	
	    If bEnableLogging = False Then Exit Function
		
		Dim oLogShell, sNow, sLogFile, oLogFile, oReadLogFile, sFileContents,aFileContents
		
	    Const ForReading = 1
	    Const ForWriting = 2
	    Const ForAppending = 8
	
	    If sLogFileLocation = "relative" Then
	        Set oLogShell = CreateObject("Wscript.Shell")
	        sLogFileLocation = oLogShell.CurrentDirectory & "\"
	        Set oLogShell = Nothing
	    End If
	   
	    If bPrependDateStampInLogFileName Then
	        sNow = Replace(Replace(Now(),"/","-"),":",".")
	        sLogFileName = sNow & " - " & sLogFileName
	        bPrependDateStampInLogFileName = False       
	    End If
	   
	    sLogFile = sLogFileLocation & sLogFileName
	   
	    If sOverWriteORAppend = "overwrite" Then
	        Set oLogFile = oLogFSO.OpenTextFile(sLogFile, ForWriting, True)
	        sOverWriteORAppend = "append"
	    Else
	        Set oLogFile = oLogFSO.OpenTextFile(sLogFile, ForAppending, True)
	    End If
	
	    If bIncludeDateStamp Then
	        Message = Now & "   " & Message
	    End If
	
	    oLogFile.WriteLine(Message)
	    oLogFile.Close
	   
	    If vLogMaximumLines > 0 Then
	      Set oReadLogFile = oLogFSO.OpenTextFile(sLogFile, ForReading, True)   
	      sFileContents = oReadLogFile.ReadAll
	      aFileContents = Split(sFileContents, vbCRLF)
	      If Ubound(aFileContents) > vLogMaximumLines Then
	        sFileContents = Replace(sFileContents, aFileContents(0) & _
	        vbCRLF, "", 1, Len(aFileContents(0) & vbCRLF))
	        Set oLogFile = oLogFSO.OpenTextFile(sLogFile, ForWriting, True)
	        oLogFile.Write(sFileContents)
	        oLogFile.Close
	      End If
	      oReadLogFile.Close
	    End If
	    
	    If vLogMaximumSize > 0 Then
	      Set oReadLogFile = oLogFSO.OpenTextFile(sLogFile, ForReading, True)  
	      sFileContents = oReadLogFile.ReadAll
	      oReadLogFile.Close
	      sFileContents = RightB(sFileContents, (vLogMaximumSize*2))
	      Set oLogFile = oLogFSO.OpenTextFile(sLogFile, ForWriting, True)
	      oLogFile.Write(sFileContents)
	      oLogFIle.Close
	    End If

	End Function

End Class

Public Function LoggerLib()

	Set LoggerLib = New ClsLoggerLib
	
End Function
