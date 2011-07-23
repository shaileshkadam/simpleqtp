Option Explicit

''' #########################################################
''' <summary>
''' A Library to work with FTP/SFTP
''' </summary>
''' <remarks></remarks>	 
''' #########################################################

Class ClsFTPLib
	
	Private oFSO
	
	Private Sub Class_Initialize

        Set oFSO = CreateObject("Scripting.FileSystemObject")
        
    End Sub
    
    Private Sub Class_Terminate

        Set oFSO = Nothing
        
    End Sub
	
	''' <summary>
    ''' Copy files from FTP/SFTP
    ''' </summary>
    ''' <param name="sHost" type="string">Host Name</param>
    ''' <param name="sUsername" type="string">Login user name</param>
    ''' <param name="sPassword" type="string">Login password</param>
    ''' <param name="sLocalPath" type="string">To Local destination</param>
    ''' <param name="sRemotePath" type="string">From Romote Path</param>
    ''' <param name="sRemoteFile" type="string">Files to be copied</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <example>
	''' Call oFTPLib.FTPDownload("SVDG0050", "a106403", "Qwerty8$", "C:\FTP", "/dsgo01/sgo/appdata/swiftout/pros", "*0504*")
    ''' </example>
	Public Function FTPDownload(sHost, sUsername, sPassword, sLocalPath, sRemotePath, sRemoteFile)
	
	  Dim oShell, sOriginalWorkingDirectory, sFTPScript
	  Dim sFTPTemp, sFTPTempFile
	  Dim sFTPResults, fFTPScript, fFTPResults, sResults
	  Const OpenAsDefault = -2
	  Const FailIfNotExist = 0
	  Const ForReading = 1
	  Const ForWriting = 2

	  Set oShell = CreateObject("WScript.Shell")
	
	  sRemotePath = Trim(sRemotePath)
	  sLocalPath = Trim(sLocalPath)
	  
	  '----------Path Checks---------
	  'Here we will check the remote path, if it contains
	  'spaces then we need to add quotes to ensure
	  'it parses correctly.
	  If InStr(sRemotePath, " ") > 0 Then
	    If Left(sRemotePath, 1) <> """" And Right(sRemotePath, 1) <> """" Then
	      sRemotePath = """" & sRemotePath & """"
	    End If
	  End If
	  
	  'Check to ensure that a remote path was
	  'passed. If it's blank then pass a "\"
	  If Len(sRemotePath) = 0 Then
	    'Please note that no premptive checking of the
	    'remote path is done. If it does not exist for some
	    'reason. Unexpected results may occur.
	    sRemotePath = "\"
	  End If
	  
	  'If the local path was blank. Pass the current
	  'working direcory.
	  If Len(sLocalPath) = 0 Then
	    sLocalpath = oShell.CurrentDirectory
	  End If
	  
	  If Not oFSO.FolderExists(sLocalPath) Then
	    'destination not found
		CreateNestedDirs(sLocalPath)
	  End If
	  
	  sOriginalWorkingDirectory = oShell.CurrentDirectory
	  oShell.CurrentDirectory = sLocalPath
	  '--------END Path Checks---------
	  
	  'build input file for ftp command
	  'For FTP subcommands, Please refer to http://msdn.microsoft.com/en-us/library/cc755356
	  sFTPScript = sFTPScript & "USER " & sUsername & vbCRLF
	  sFTPScript = sFTPScript & sPassword & vbCRLF
	  sFTPScript = sFTPScript & "cd " & sRemotePath & vbCRLF
	  sFTPScript = sFTPScript & "binary" & vbCRLF
	  sFTPScript = sFTPScript & "prompt n" & vbCRLF
	  sFTPScript = sFTPScript & "mget " & sRemoteFile & vbCRLF
	  sFTPScript = sFTPScript & "quit" & vbCRLF & "quit" & vbCRLF & "quit" & vbCRLF
	
	
	  sFTPTemp = oShell.ExpandEnvironmentStrings("%TEMP%")
	  'GetTempName: Returns a randomly generated temporary file name
	  'The GetTempName method does not create a file. It provides only a temporary file name that can be used with CreateTextFile to create a file
	  CreateNestedDirs(sFTPTemp & "\" & "FTP")
	  sFTPTempFile = sFTPTemp & "\FTP\" & oFSO.GetTempName
	  sFTPResults = sFTPTemp & "\FTP\" & oFSO.GetTempName
	
	  'Write the input file for the ftp command to a temporary file.
	  Set fFTPScript = oFSO.CreateTextFile(sFTPTempFile, True)
	  fFTPScript.WriteLine(sFTPScript)
	  fFTPScript.Close
	  Set fFTPScript = Nothing  
	
	  'For FTP Command, please refer to http://msdn.microsoft.com/en-us/library/cc756013
	  oShell.Run "%comspec% /c FTP -n -s:" & sFTPTempFile & " " & sHost & " > " & sFTPResults, 0, TRUE
	  
	  Wscript.Sleep 1000
	  
	  'Check results of transfer.
	  Set fFTPResults = oFSO.OpenTextFile(sFTPResults, ForReading, FailIfNotExist, OpenAsDefault)
	  sResults = fFTPResults.ReadAll
	  fFTPResults.Close
	  
	  'oFSO.DeleteFile(sFTPTempFile)
	  'oFSO.DeleteFile (sFTPResults)
	  
	  If InStr(sResults, "226 Transfer complete.") > 0 Then
	    FTPDownload = True
	  ElseIf InStr(sResults, "File not found") > 0 Then
	    FTPDownload = "Error: File Not Found"
	  ElseIf InStr(sResults, "cannot log in.") > 0 Then
	    FTPDownload = "Error: Login Failed."
	  Else
	    FTPDownload = "Error: Unknown."
	  End If
	  
	  Set oShell = Nothing
	  
	End Function

	''' <summary>
    ''' Creates multiple folders like CMD.EXE's internal MD command
    ''' </summary>
    ''' <param name="DirName" type="string">folder(s) to be created, single or multi level, absolute or relative</param>
    ''' <returns></returns>
    ''' <remarks>By default VBScript can only create one level of folders at a time</remarks>
    ''' <example>
	''' UNC path
	''' oFSOLib.CreateNestedDirs "\\MYSERVER\D$\Test01\Test02\Test03\Test04"
	''' Absolute path
	''' oFSOLib.CreateNestedDirs "C:\Test11\Test12\Test13\Test14"
	''' Relative path
	''' oFSOLib.CreateNestedDirs "Test21\Test22\Test23\Test24"
    ''' </example>
	Public Function CreateNestedDirs(MyDirName)
	
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

End Class

Public Function FTPLib()
	
	Set FTPLib = New ClsFTPLib

End Function
