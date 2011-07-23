Option Explicit

''' #########################################################
''' <summary>
''' A library to work with file/folder system
''' </summary>
''' <remarks></remarks>	 
''' #########################################################

Public Const ForReading = 1
Public Const ForWriting = 2
Public Const ForAppending = 8
  
Class ClsFSOLib

	Private oFSO
	
	Private Sub Class_Initialize

        Set oFSO = CreateObject("Scripting.FileSystemObject")
        
    End Sub
    
    Private Sub Class_Terminate

        Set oFSO = Nothing
        
    End Sub

	''' <summary>
    ''' Create a new txt file
    ''' </summary>
    ''' <param name="FilePath" type="string">location of the file and its name</param>
    ''' <returns>Created file</returns>
    ''' <remarks></remarks>
	Public Function CreateFile(ByVal FilePath)
		
		' varibale that will hold the new file object
		dim NewFile
		' create the new text ile
		set NewFile = oFSO.CreateTextFile(FilePath, True)
		set CreateFile = NewFile
		
	End Function

	''' <summary>
    ''' Create a new Folder
    ''' </summary>
    ''' <param name="FolderPath" type="string">location of the folder</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function CreateFolder(ByVal FolderPath)
		
		If NOT oFSO.FolderExists(FolderPath) Then
			oFSO.CreateFolder(FolderPath)
		End If
		
	End Function
	
	''' <summary>
    ''' Creates multiple folders like CMD.EXE's internal MD command
    ''' </summary>
    ''' <param name="DirName" type="string">folder(s) to be created, single or multi level, absolute or relative</param>
    ''' <returns></returns>
    ''' <remarks>By default VBScript can only create one level of folders at a time</remarks>
    ''' <example>
	''' UNC path
	''' FSOLib.CreateNestedDirs "\\MYSERVER\D$\Test01\Test02\Test03\Test04"
	''' Absolute path
	''' FSOLib.CreateNestedDirs "C:\Test11\Test12\Test13\Test14"
	''' Relative path
	''' FSOLib.CreateNestedDirs "Test21\Test22\Test23\Test24"
    ''' </example>
	Public Function CreateNestedDirs(ByVal DirName)
	
	    Dim arrDirs, i, idxFirst, strDir, strDirBuild
	    ' Convert relative to absolute path
	    strDir = oFSO.GetAbsolutePathName(DirName)
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
	        If Not CheckFolderExists( strDirBuild ) Then 
	            CreateFolder strDirBuild
	        End if
	    Next

	End Function

	''' <summary>
    ''' Check if a specific file exist
    ''' </summary>
    ''' <param name="FilePath" type="string">location of the file and its name</param>
    ''' <returns>true/false</returns>
    ''' <remarks></remarks>
	Public Function CheckFileExists(ByVal FilePath)
	    
	    CheckFileExists = oFSO.FileExists(FilePath)
	    
	End Function

	''' <summary>
    ''' Check if a specific Folder exist
    ''' </summary>
    ''' <param name="FolderPath" type="string">location of the Folder</param>
    ''' <returns>true/false</returns>
    ''' <remarks></remarks>
	Public Function CheckFolderExists(ByVal FolderPath)
	
	    CheckFolderExists = oFSO.FolderExists(FolderPath)
	    
	End Function

	''' <summary>
    ''' Write data to file
    ''' </summary>
    ''' <param name="FileRef" type="object">reference to the file</param>
    ''' <param name="str" type="string">data to be written to the file</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function WriteToFile(ByRef FileRef,ByVal str)

	   FileRef.WriteLine(str)
	   
	End Function
	 
	''' <summary>
    ''' Read line from file
    ''' </summary>
    ''' <param name="FileRef" type="object">reference to the file</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function ReadLineFromFile(ByRef FileRef)

	    ReadLineFromFile = FileRef.ReadLine
	    
	End Function

	''' <summary>
    ''' Read line from file
    ''' </summary>
    ''' <param name="FileRef" type="object">reference to the file</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function CloseFile(ByRef FileRef)
	
	    FileRef.close
	    
	End Function
	
	''' <summary>
    ''' Opens a specified file and returns an object that can be used to read from, write to, or append to the file.
    ''' </summary>
    ''' <param name="FilePath" type="string">location of the file and its name</param>
    ''' <param name="mode" type="int">ForReading - 1, ForWriting - 2, ForAppending - 8</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function OpenFile(ByVal FilePath,ByVal mode)
	    
	    ' open the txt file and retunr the File object
	    set OpenFile = oFSO.OpenTextFile(FilePath, mode, True)
	    
	End Function

	''' <summary>
    ''' Compare two text files.
    ''' </summary>
    ''' <param name="FilePath1" type="string">location of the first file to be compared</param>
    ''' <param name="FilePath2" type="string">location of the second file to be compared</param>
    ''' <param name="FilePathDiff" type="string">location of the diffrences file</param>
    ''' <param name="ignoreWhiteSpace" type="bool">determine whether or not to ignore differences in whitespace characters</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <example>
	''' call FSOLib.FileCompare("C:\1.txt", "C:\2.txt", "C:\3.txt", True)
    ''' </example>
	Public Function FileCompare(ByVal FilePath1, ByVal FilePath2, ByVal FilePathDiff,ByVal ignoreWhiteSpace)
	    
	    dim differentFiles
	    differentFiles = false
	    dim f1, f2, f_diff
	    ' open the files
	    set f1 = OpenFile(FilePath1,1)
	    set f2 = OpenFile(FilePath2,1)
	    set f_diff = OpenFile(FilePathDiff,8)
	    dim rowCountF1, rowCountF2
	    rowCountF1 = 0
	    rowCountF2 = 0
	    dim str
	    ' count how many lines there are in first file
	    While not f1.AtEndOfStream 
	        str = ReadLineFromFile(f1)
	        rowCountF1= rowCountF1 + 1
	    Wend
	 
	    ' count how many lines there are in second file
	    While not f2.AtEndOfStream 
	        str = ReadLineFromFile(f2)
	        rowCountF2= rowCountF2 + 1
	    Wend
	 
	    ' re-open the files to go back to the first line in the files
	    set f1 = OpenFile(FilePath1,1)
	    set f2 = OpenFile(FilePath2,1)
	 
	    ' compare the number of lines in the two files.
	    ' assign biggerFile - the file that contain more lines
	    ' assign smallerFile - the file that contain less lines
	    dim biggerFile, smallerFile
	    set biggerFile = f1
	    set smallerFile = f2
	    If ( rowCountF1 < rowCountF2) Then
	        set smallerFile = f1
	        set biggerFile = f2
	    End If
	 
	    dim lineNum,str1, str2
	    lineNum = 1
	    str = "Line" & vbTab & "File1" & vbTab & vbTab & "File2"
	    WriteToFile f_diff,str
	     ' loop on all the lines in the samller file
	    While not smallerFile.AtEndOfStream 
	        ' read line from both files
	        str1 = ReadLineFromFile(f1)
	        str2 = ReadLineFromFile(f2)
	 
	        ' check if we need to ignore white spaces, if yes, trim the two lines
	        If Not ignoreWhiteSpace Then
	           Trim(str1)
	           Trim(str2)
	        End If
	 
	        ' if there is a diffrence between the two lines, write them to the diffrences file 
	        If not (str1 = str2) Then
	            differentFiles = true
	            str = lineNum & vbTab & str1 & vbTab & vbTab & str2
	            WriteToFile f_diff,str
	        End If
	        lineNum = lineNum + 1
	    Wend
	 
	    ' loop on the bigger lines, to write its line two the diffrences file
	    While not biggerFile.AtEndOfStream 
	        str1 = ReadLineFromFile(biggerFile)
	        str = lineNum & vbTab & "" & vbTab & vbTab & str2
	        WriteToFile f_diff,str
	        lineNum = lineNum + 1
	    Wend
	  
	    FileCompare = Not differentFiles
	    
	End Function
	
	''' <summary>
    ''' Appends a line to a file
    ''' </summary>
    ''' <param name="FilePath" type="string">location of the file and its name</param>
    ''' <param name="sLine" type="string">the line to be appended</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function AppendToFile(ByVal FilePath,ByVal sLine)
	    
	    Set f = OpenFile(FilePath, ForAppending)
	    f.WriteLine sLine
	    f.Close
	    
	End Function

	''' <summary>
    ''' Copy a file to another path
    ''' </summary>
    ''' <param name="FilePathSource" type="string">location of the source file and its name</param>
    ''' <param name="FilePathDest" type="string">location of the destination file and its name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function FileCopy(ByVal FilePathSource,ByVal FilePathDest)
	    
	    ' copy source file to destination file
	    If oFSO.FileExists(FilePathSource) Then
		    oFSO.CopyFile FilePathSource, FilePathDest
	    End If
	    
	End Function

	''' <summary>
    ''' Copy a folder to destination path
    ''' </summary>
    ''' <param name="FolderPathSource" type="string">location of the source folder</param>
    ''' <param name="FolderPathDest" type="string">location of the destination folder</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function FolderCopy(ByVal FolderPathSource,ByVal FolderPathDest)
		
		If oFSO.FolderExists(FolderPathSource) Then
			oFSO.CopyFolder FolderPathSource, FolderPathDest
		End If
		
	End Function

	''' <summary>
    ''' Count the nubmer of Files in that directory and all sub directories
    ''' </summary>
    ''' <param name="strFolder" type="string">location of the directory</param>
    ''' <returns>the nubmer of Files</returns>
    ''' <remarks></remarks>
	Public Function CounteFiles(ByVal strFolder)
		
		Dim ParentFld
		Dim SubFld
		Dim IntCount

		Set ParentFld = oFSO.GetFolder(strFolder)
		' count the number of files in the current directory
		IntCount = ParentFld.Files.Count
		For Each SubFld In ParentFld.SubFolders
		' count all files in each subfolder ¨C recursion point
		IntCount = IntCount + CounteFiles(SubFld.Path)
		Next
		' return counted files
		CounteFiles = IntCount
		
	End Function

	''' <summary>
    ''' Count the nubmer of folders in that directory and all sub directories
    ''' </summary>
    ''' <param name="strFolder" type="string">location of the directory</param>
    ''' <returns>the nubmer of folders</returns>
    ''' <remarks></remarks>
	Public Function CounterFolders(ByVal strFolder)
		
		Dim ParentFld
		Dim SubFld
		Dim IntCount

		Set ParentFld = oFSO.GetFolder(strFolder)
		' count the number of files in the current directory
		IntCount = ParentFld.subfolders.Count
		For Each SubFld In ParentFld.SubFolders
		' count all files in each subfolder ¨C recursion point
		IntCount = IntCount + CounterFolders(SubFld.Path)
		Next
		' return counted files
		CounterFolders = IntCount
		
	End Function
	 
	''' <summary>
    ''' Delete a file
    ''' </summary>
    ''' <param name="FilePath" type="string">location of the file to be deleted</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function FileDelete(ByVal FilePath)
	  
	  If oFSO.FileExists(FilePath) Then
	    oFSO.DeleteFile(FilePath)
	  End If
	  
	End Function
	
	''' <summary>
    ''' Delete a folder
    ''' </summary>
    ''' <param name="FolderPath" type="string">location of the folder to be deleted</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function FolderDelete(ByVal FolderPath)
		
		If oFSO.FolderExists(FolderPath) Then
		    oFSO.GetFolder(FolderPath).Delete
	    End If
	    
	End Function
	
	''' <summary>
    ''' Get file name for a given file path
    ''' </summary>
    ''' <param name="FilePath" type="string">the path of the file</param>
    ''' <returns>file name</returns>
    ''' <remarks></remarks>
    ''' <example>
	''' MsgBox FSOLib.GetFileName("C:\Config.xls")
    ''' </example>
	Public Function GetFileName(ByVal FilePath)

	   GetFileName = oFSO.GetFileName(FilePath)
	      
	End Function
	
	''' <summary>
    ''' Get Parent folder path for a given file/folder path
    ''' </summary>
    ''' <param name="Path" type="string">the path of the file/folder</param>
    ''' <returns>Parent folder path</returns>
    ''' <remarks></remarks>
    ''' <example>
	''' MsgBox FSOLib.GetParentFolderPath("C:\Documents and Settings\lwfwind\Desktop\Config.xls")
	''' MsgBox FSOLib.GetParentFolderPath("C:\Documents and Settings\lwfwind\Desktop\Master")
    ''' </example>
	Public Function GetParentFolderPath(ByVal Path)

	   GetParentFolderPath = oFSO.GetParentFolderName(Path)
	         
	End Function
	
	''' <summary>
    ''' Get Parent folder name for a given file/folder path
    ''' </summary>
    ''' <param name="Path" type="string">the location of the file/folder</param>
    ''' <returns>Parent folder name</returns>
    ''' <remarks></remarks>
    ''' <example>
	''' MsgBox FSOLib.GetParentFolderName("C:\Documents and Settings\lwfwind\Desktop\Config.xls")
	''' MsgBox FSOLib.GetParentFolderName("C:\Documents and Settings\lwfwind\Desktop\Master")
    ''' </example>
	Public Function GetParentFolderName(ByVal Path)
		
		strParentFolderPath = oFSO.GetParentFolderName(Path)
		arrFolder = Split(strParentFolderPath, "\")
		GetParentFolderName = arrFolder(UBound(arrFolder))
	         
	End Function

	''' <summary>
    ''' Enumerate files in folder and subfolders and save to Dictionary
    ''' </summary>
    ''' <param name="oDict" type="Dictionary">the oDict to be saved</param>
    ''' <param name="FolderPath" type="string">the path of the folder</param>
    ''' <returns>array of files location</returns>
    ''' <remarks></remarks>
    ''' <example>
    ''' Dim oDict,key
	''' Set oDict = CreateObject("Scripting.Dictionary")
	''' call FSOLib.GetFilesRecursively (oDict, "C:\Documents and Settings\a106403\Desktop\Browser")
    ''' For Each key In oDict
	''' 	msgbox oDict(key)
	''' Next
    ''' </example>
	Public Function GetFilesRecursively(ByRef oDict, ByVal FolderPath)

		Dim objFolder, objFile
		set objFolder = oFSO.GetFolder(FolderPath)
		For each objFile in objFolder.Files
			 oDict.add objFile.path, objFile.path
		next	
		For each objfolder in objFolder.SubFolders
			GetFilesRecursively oDict, objfolder.Path
		next
		
	end function

	''' <summary>
    ''' Get a List of All the Folders and Files in a Folder and Its Subfolders by recursion and save to array
    ''' </summary>
    ''' <param name="arr" type="array">the arr to be saved</param>
    ''' <param name="FolderPath" type="string">the path of the folder</param>
    ''' <returns>array of files/folder location</returns>
    ''' <remarks></remarks>
    ''' <example>
    ''' Dim arr()
	''' call FSOLib.GetFilesFoldersRecursively (arr, "C:\Documents and Settings\a106403\Desktop\Browser")
    ''' </example>
	Public Function GetFilesFoldersRecursively(ByRef arr, ByVal FolderPath)

		Dim oParentfolder, objFolder, objFile,i
		Set oParentfolder = oFSO.GetFolder(FolderPath)
		' Get all files in Parent Folders
		For Each objFile in oParentfolder.Files
			ReDim Preserve arr(i)
			arr(i) = objFile.path
			i = i + 1
		Next
		
		For Each objFolder in oParentfolder.SubFolders
			ReDim Preserve arr(i)
		    arr(i) = objFolder.path
			i = i + 1
		    GetSubFolders objFolder.Path
		Next
		GetFilesFoldersRecursively = arr
		
	End Function
	
	Public Function GetSubFolders(strFolderName2)
			
		Dim oParentfolder2,objFolder2,objFile2
	    Set oParentfolder2 = oFSO.GetFolder(strFolderName2)
	
	    For Each objFolder2 in oParentfolder2.SubFolders
				ReDim Preserve arr(i)
		        arr(i) = objFolder2.path
				i = i + 1
	
	        For Each objFile2 in oParentfolder2.Files
				ReDim Preserve arr(i)
		        arr(i) = objFile2.path
				i = i + 1
	        Next
	
	        GetSubFolders objFolder2.Path
	    Next
		    
	End Function
	
	''' <summary>
    ''' Finds the first instance of a file within the root folder or one of its subfolders
    ''' </summary>
    ''' <param name="strRootFolder" type="array">the root folder to be searched</param>
    ''' <param name="strFilename" type="string">the name of the file</param>
    ''' <returns>String with full file pathname based on root folder and file name</returns>
    ''' <remarks></remarks>
    ''' <example>
	''' Dim strRootFolder, strFilename
	''' strRootFolder = "C:\Program Files"
	''' strFilename = "pdfshell.dll"
	''' MsgBox FSOLib.FindFileRecursively(strRootFolder, strFilename)
    ''' </example>
	Public Function FindFileRecursively(ByVal strRootFolder, ByVal strFilename)
	
		Dim FSO    
		Dim strFullPathToSearch    
		Dim objSubFolders, subfolder    
		
		Set FSO = CreateObject("Scripting.FileSystemObject")    
		'Initialize function    
		FindFileRecursively = ""    
		'Check that filename is not empty    
		If strFileName = "" Then Exit Function    
		'Get full file pathname    
		strFullPathToSearch = strRootFolder & "\" & strFilename    
		'Check if root folder exists    
		If FSO.FolderExists(strRootFolder) Then        
			'Check if file exists under root folder        
			If FSO.FileExists(strFullPathToSearch) Then            
			  FindFileRecursively = strFullPathToSearch        
			Else            
			  'Get subfolders            
			  Set objSubFolders = FSO.GetFolder(strRootFolder).SubFolders            
			  For Each subfolder in objSubFolders                
			      strFullPathToSearch = strRootFolder & "\" & subfolder.name                
			      FindFileRecursively = FindFileRecursively(strFullPathToSearch, strFilename)                
			      If FindFileRecursively <> "" Then                    
			          Exit For                
			      End If            
			  Next        
			End If    
		End If
	
	End Function
	
	''' <summary>
    ''' Finds the specific String within file under the root folder or one of its subfolders
    ''' </summary>
    ''' <param name="sFolderLoction" type="array">the folder to be searched</param>
    ''' <param name="sSearchString" type="string">the string to be searched</param>
    ''' <returns>True/False</returns>
    ''' <remarks></remarks>
    ''' <example>
	''' Dim sFolderLoction, sSearchString
	''' sFolderLoction = "C:\FTP"
	''' sSearchString = "1132005"
	''' MsgBox FSOLib.IsTextFoundInFolder(sFolderLoction, sSearchString)
    ''' </example>
	Public Function IsTextFoundInFolder(ByVal sFolderLoction, ByVal sSearchString)

		Dim arrfilename()
		Dim i
		Dim strLine
		call GetFilesRecursively (arrfilename, sFolderLoction)
		For i = LBound(arrfilename) To UBound(arrfilename)
			Set f = OpenFile(arrfilename(i), 1) 
			Do While f.AtEndOfStream <> True 
				strLine = f.ReadLine 
				if InStr(1, Trim(strLine),Trim(sSearchString)) <> 0 Then
					IsTextFoundInFolder = True 
					Exit Function
				End if
			Loop 
		Next
		IsTextFoundInFolder = false 
	
	End Function
	
	''' <summary>
    ''' Find and replace text in a text file
    ''' </summary>
    ''' <param name="filename" type="string">the location of the text file</param>
    ''' <param name="oldstr" type="string">the old string to be replaced</param>
    ''' <param name="newstr" type="string">the new string used to replace the old staing</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    ''' <example>
	''' FSOLib.ReplaceTextInTextFile "C:\test.txt", "Johnston", "JohnstonJohnston"
    ''' </example>
	Public Function ReplaceTextInTextFile(Byval filename, Byval oldstr, Byval newstr)
	
		Dim objFile, strText,strNewText
		Set objFile = oFSO.OpenTextFile(filename, ForReading)	
		strText = objFile.ReadAll
		objFile.Close
		strNewText = Replace(strText, oldstr, newstr)
		Set objFile = oFSO.OpenTextFile(filename, ForWriting)
		objFile.WriteLine strNewText
		objFile.Close
	
	End Function

	''' <summary>
    ''' Reads the file line by line and save into a array
    ''' </summary>
    ''' <param name="strFileName" type="string">the location of the text file</param>
    ''' <returns>the number of lines</returns>
    ''' <remarks></remarks>
    ''' <example>
	''' Dim arr, i
	''' arr = FSOLib.ReadFileIntoArray(DESKTOP_Path & "\VBscript Resource.txt")
	''' For i = LBound(arr) To UBound(arr) - 1
	''' 	MsgBox arr(i)
	''' Next
    ''' </example>
	Public Function ReadFileIntoArray(ByVal strFileName) 
	
		Dim intLineCount, f, intIndex, strLine, fileArray() 
		'Size of File 
		intLineCount = FileLineCount(strFileName) - 1 
		ReDim Preserve fileArray(intLineCount) 'resize array to proper size 
		
		' Open file and read contents into an array 
		Set f = OpenFile(strFileName, ForReading) 
		intIndex = 0 
		Do While f.AtEndOfStream <> True 
			strLine = f.ReadLine 
			fileArray(intIndex) = strline 
			intIndex = intIndex + 1 
		Loop 
		f.close 'Close File 
		ReadFileIntoArray = fileArray 'Return Array 
	
	End Function 

	''' <summary>
    ''' Get the number of lines in a file
    ''' </summary>
    ''' <param name="strFileName" type="string">the location of the text file</param>
    ''' <returns>the number of lines</returns>
    ''' <remarks></remarks>
    ''' <example>
	''' MsgBox FSOLib.FileLineCount("c:\VBscript Resource.txt")
    ''' </example>
	Public Function FileLineCount(ByVal strFileName) 
	 
		' Create a File System Object 
		Dim fileContents, lineCount, sReadLine 	
		' Get the file contents 
		Set fileContents = OpenFile(strFileName, 1) 
		' Loop through counting the lines 
		lineCount = 0 
		Do While fileContents.AtEndOfStream <> True 
			sReadLine = fileContents.ReadLine 
			lineCount = lineCount + 1 
		Loop 
		' Return the file's line count 
		FileLineCount = lineCount 
		' Cleanup the file system objects 
		fileContents.Close 
		Set fileContents = Nothing 
	
	End Function 
	
End Class	
	
Public Function FSOLib()
	
	Set FSOLib = New ClsFSOLib

End Function
