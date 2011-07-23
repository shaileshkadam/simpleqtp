Option Explicit

''' #########################################################
''' <summary>
''' A Library to work with registry
''' </summary>
''' <remarks></remarks>	
''' <example>

''' Dim strRegPath, strValueName, RegValueData, RegDataType

''' ----------------------------------------------------------------------------
''' strRegPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\test"
''' RegLib.CreateRegKey(strRegPath)
''' RegLib.DeleteRegKey(strRegPath)

''' ----------------------------------------------------------------------------
''' strRegPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{241F2BF7-69EB-42A4-9156-96B2426C7504}" 
''' ----------------------------------------------------------------------------
''' strValueName = "test REG_SZ"
''' RegValueData = "test"
''' RegDataType = REG_SZ
''' call RegLib.SetRegValue(strRegPath, strValueName,RegValueData, RegDataType)

''' ----------------------------------------------------------------------------
''' strValueName = "test REG_EXPAND_SZ"
''' RegValueData = "http://go.microsoft.com/fwlink/?LinkId=81488"
''' RegDataType = REG_EXPAND_SZ
''' call RegLib.SetRegValue(strRegPath, strValueName,RegValueData, RegDataType)

''' ----------------------------------------------------------------------------
''' strValueName = "test REG_BINARY"
''' RegValueData = "&H01,&Ha2"
''' RegDataType = REG_BINARY
''' call RegLib.SetRegValue(strRegPath, strValueName,RegValueData, RegDataType)

''' ----------------------------------------------------------------------------
''' strValueName = "test REG_DWORD"
''' RegValueData = 2
''' RegDataType = REG_DWORD
''' call RegLib.SetRegValue(strRegPath, strValueName,RegValueData, RegDataType)

''' ----------------------------------------------------------------------------
''' strValueName = "test REG_MULTI_SZ"
''' RegValueData = "a,b,c"
''' RegDataType = REG_MULTI_SZ
''' call RegLib.SetRegValue(strRegPath, strValueName,RegValueData, RegDataType)

''' ----------------------------------------------------------------------------
''' strRegPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{241F2BF7-69EB-42A4-9156-96B2426C7504}" 
''' strValueName = "test REG_SZ"
''' MsgBox RegLib.ReadRegValue(strRegPath, strValueName)
''' Call RegLib.DeleteRegValue(strRegPath, strValueName)

''' ----------------------------------------------------------------------------
''' strValueName = "test REG_EXPAND_SZ"
''' MsgBox RegLib.ReadRegValue(strRegPath, strValueName)
''' Call RegLib.DeleteRegValue(strRegPath, strValueName)

''' ----------------------------------------------------------------------------
''' strValueName = "test REG_BINARY"
''' MsgBox RegLib.ReadRegValue(strRegPath, strValueName)
''' Call RegLib.DeleteRegValue(strRegPath, strValueName)

''' ----------------------------------------------------------------------------
''' strValueName = "test REG_DWORD"
''' MsgBox RegLib.ReadRegValue(strRegPath, strValueName)
''' Call RegLib.DeleteRegValue(strRegPath, strValueName)

''' ----------------------------------------------------------------------------
''' strValueName = "test REG_MULTI_SZ"
''' MsgBox RegLib.ReadRegValue(strRegPath, strValueName)
''' Call RegLib.DeleteRegValue(strRegPath, strValueName)

''' </example>
''' #########################################################

Const strComputer = "."
'If I use &H syntax in library, it gives an error and QTP does not allow to import this file in my test due to syntax errors. 
'But, when to use &H syntax in test itself, and use it. It works fine.
'Solution: using ExecuteGlobal
ExecuteGlobal "Const HKEY_CLASSES_ROOT = &H80000000"
ExecuteGlobal "const HKEY_CURRENT_USER = &H80000001"
ExecuteGlobal "Const HKEY_LOCAL_MACHINE = &H80000002"
ExecuteGlobal "Const HKEY_USERS = &H80000003"
ExecuteGlobal "Const HKEY_CURRENT_CONFIG = &H80000005"

Const REG_SZ                         =  1
Const REG_EXPAND_SZ                  =  2
Const REG_BINARY                     =  3
Const REG_DWORD                      =  4
Const REG_DWORD_BIG_ENDIAN           =  5
Const REG_LINK                       =  6
Const REG_MULTI_SZ                   =  7
Const REG_RESOURCE_LIST              =  8
Const REG_FULL_RESOURCE_DESCRIPTOR   =  9
Const REG_RESOURCE_REQUIREMENTS_LIST = 10
Const REG_QWORD                      = 11
		
Class ClsRegLib
	
	Private oReg

	
	''' <summary>
    ''' Class Initialization procedure. Creates Excel Singleton.
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Initialize()
			
		Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_ 
		strComputer & "\root\default:StdRegProv") 
		
	End Sub
	
	''' <summary>
    ''' Class Termination procedure
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Terminate()
		
		Set oReg = Nothing
		
	End Sub

 	''' <summary>
    ''' Reads a value from the registry of any WMI enabled computer.
    ''' </summary>
    ''' <param name="strRegPath" type="string">a full registry key path, e.g.HKLM\SOFTWARE\Microsoft\DirectX OR HKEY_CLASSES_ROOT\SOFTWARE\Microsoft\DirectX
    ''' <param name="strValueName" type="string">the value name to be queried, e.g.InstalledVersion
    ''' <returns>Reg Value</returns>    
    ''' <remarks></remarks>
	Public Function ReadRegValue(ByVal strRegPath, ByVal strValueName)
	
		Dim EndStrPos, Length, sHive, Hive, strKeyPath, strValue
		Dim arrValueNames, arrValueTypes, valRegType, sResult, i ,valRegVal
		
		sResult = ""
		
		'Initiate custom error handling
    	On Error Resume Next
    	
		EndStrPos = Instr(strRegPath,"\")
		Length = EndStrPos  - 1
		sHive = Mid(strRegPath,1,Length)
		strKeyPath = Mid(strRegPath,EndStrPos + 1)
		
		select Case UCase(sHive)
		
			Case "HKCR","HKEY_CLASSES_ROOT"
				Hive = HKEY_CLASSES_ROOT
			Case "HKCU","HKEY_CURRENT_USER"
				Hive = HKEY_CURRENT_USER
			Case "HKLM","HKEY_LOCAL_MACHINE"
				Hive = HKEY_LOCAL_MACHINE
			Case "HKU","HKEY_USERS"
				Hive = HKEY_USERS
			Case "HKCC","HKEY_CURRENT_CONFIG"
				Hive = HKEY_CURRENT_CONFIG	
			Case Else
				ReadRegValue = "The Top key is not correct"
				Exit Function
			
		End Select
		
		' Get a list of all values in the registry path;
	    ' we need to do this in order to find out the
	    ' exact data type for the requested value
	    oReg.EnumValues Hive, strKeyPath, arrValueNames, arrValueTypes
	
	    ' If no values were found, we'll need to retrieve a default value
	    If Not IsArray( arrValueNames ) Then
	        arrValueNames = Array( "" )
	        arrValueTypes = Array( REG_SZ )
	    else
			' Loop through all values in the list . . .
			For i = 0 To UBound(arrValueNames)
				' . . . and find the one requested
				If Trim(UCase(arrValueNames(i))) = Trim(UCase(strValueName)) Then
					' Read the requested value's data type
					valRegType = arrValueTypes(i)
					' Based on the data type, use the appropriate query to retrieve the data
					' http://msdn.microsoft.com/en-us/library/aa393664(v=VS.85).aspx
					Select Case valRegType
						Case REG_SZ
							oReg.GetStringValue Hive, strKeyPath, _
												  strValueName, valRegVal
						Case REG_EXPAND_SZ
							oReg.GetExpandedStringValue Hive, strKeyPath, _
														  strValueName, valRegVal
						Case REG_BINARY ' returns an array of bytes
							oReg.GetBinaryValue Hive, strKeyPath, _
												  strValueName, valRegVal
						Case REG_DWORD
							oReg.GetDWORDValue Hive, strKeyPath, _
												 strValueName, valRegVal
						Case REG_MULTI_SZ ' returns an array of strings
							oReg.GetMultiStringValue Hive, strKeyPath, _
													   strValueName, valRegVal
						Case REG_QWORD
							' Windows Server 2003, Windows XP, Windows 2000, Windows NT 4.0, and Windows Me/98/95:  This method is not available.
							oReg.GetQWORDValue Hive, strKeyPath, _
												 strValueName, valRegVal
						Case Else
							ReadRegValue = "The RegType is not correct"
							Exit Function
					End Select
				End If
			Next
	    end if
		If isnull(valRegVal) Or IsEmpty(valRegVal) then
			ReadRegValue = "null"
		Else
			If valRegType = REG_BINARY Or valRegType = REG_MULTI_SZ Then
				For i = 0 To UBound(valRegVal)
					If sResult = "" then
		            	sResult = valRegVal(i)
		            Else
		            	sResult = sResult & "," & valRegVal(i)
		            End if
		        Next
		        ReadRegValue = sResult
			else
				ReadRegValue = valRegVal
			End if
		End if

	    If Err.Number > 0 Then
	        Err.Clear
	        On Error Goto 0
			ReadRegValue = "Unknown error"
	        Exit Function
	    End if
	        
	End Function
	
	
 	''' <summary>
    ''' Set a value under the specific key from the registry of any WMI enabled computer.
    ''' </summary>
    ''' <param name="strRegPath" type="string">a full registry key path, e.g.HKLM\SOFTWARE\Microsoft\DirectX OR HKEY_CLASSES_ROOT\SOFTWARE\Microsoft\DirectX
    ''' <param name="strValueName" type="string">the value name to be set, e.g.InstalledVersion
    ''' <param name="RegValueData" type="string">data value e.g. 11
    ''' <param name="RegDataType" type="string">registry data type, e.g.REG_SZ Please refer to http://msdn.microsoft.com/en-us/library/aa392326(v=vs.85).aspx
    ''' <returns></returns>    
    ''' <remarks></remarks>
	Public Function SetRegValue(ByVal strRegPath, ByVal strValueName, ByVal RegValueData, ByVal RegDataType)
	
		Dim EndStrPos, Length, sHive, Hive, strKeyPath, arrRegValueData
		
		'Initiate custom error handling
    	On Error Resume Next
    	
		EndStrPos = Instr(strRegPath,"\")
		Length = EndStrPos  - 1
		sHive = Mid(strRegPath,1,Length)
		strKeyPath = Mid(strRegPath,EndStrPos + 1)
		
		select Case UCase(sHive)
		
			Case "HKCR","HKEY_CLASSES_ROOT"
				Hive = HKEY_CLASSES_ROOT
			Case "HKCU","HKEY_CURRENT_USER"
				Hive = HKEY_CURRENT_USER
			Case "HKLM","HKEY_LOCAL_MACHINE"
				Hive = HKEY_LOCAL_MACHINE
			Case "HKU","HKEY_USERS"
				Hive = HKEY_USERS
			Case "HKCC","HKEY_CURRENT_CONFIG"
				Hive = HKEY_CURRENT_CONFIG	
			Case Else
				SetRegValue = "The Top key is not correct"
				Exit Function
			
		End Select
		

		' Based on the data type, use the appropriate query to set the data
		' http://msdn.microsoft.com/en-us/library/aa393664(v=VS.85).aspx
		Select Case RegDataType
			Case REG_SZ
				oReg.SetStringValue Hive, strKeyPath, _
									  strValueName, RegValueData
			Case REG_EXPAND_SZ
				oReg.SetExpandedStringValue Hive, strKeyPath, _
											  strValueName, RegValueData
			Case REG_BINARY
				arrRegValueData = split(RegValueData, ",")
				oReg.SetBinaryValue Hive, strKeyPath, _
									  strValueName, arrRegValueData
			Case REG_DWORD
				oReg.SetDWORDValue Hive, strKeyPath, _
									 strValueName, RegValueData
			Case REG_MULTI_SZ
				arrRegValueData = split(RegValueData, ",")
				oReg.SetMultiStringValue Hive, strKeyPath, _
										   strValueName, arrRegValueData
			Case REG_QWORD
				' Windows Server 2003, Windows XP, Windows 2000, Windows NT 4.0, and Windows Me/98/95:  This method is not available.
				oReg.SetQWORDValue Hive, strKeyPath, _
									 strValueName, RegValueData
			Case Else
				SetRegValue = "The RegType is not correct"
				Exit Function
		End Select


	    If Err.Number > 0 Then
	        Err.Clear
	        On Error Goto 0
			SetRegValue = "Unknown error"
	        Exit Function
	    End if
	        
	End Function	
	
		
 	''' <summary>
    ''' Create a key from the registry of any WMI enabled computer.
    ''' </summary>
    ''' <param name="strRegPath" type="string">a full registry key path, e.g.HKLM\SOFTWARE\Microsoft\DirectX OR HKEY_CLASSES_ROOT\SOFTWARE\Microsoft\DirectX
    ''' <returns></returns>    
    ''' <remarks></remarks>
	Public Function CreateRegKey(ByVal strRegPath)
			
		Dim EndStrPos, Length, Hive, strKeyPath
		
		'Initiate custom error handling
    	On Error Resume Next
    	
		EndStrPos = Instr(strRegPath,"\")
		Length = EndStrPos  - 1
		Hive = Mid(strRegPath,1,Length)
		strKeyPath = Mid(strRegPath,EndStrPos + 1)
		
		select Case UCase(Hive)
		
			Case "HKCR","HKEY_CLASSES_ROOT"
				oReg.CreateKey HKEY_CLASSES_ROOT,strKeyPath
			Case "HKCU","HKEY_CURRENT_USER"
				oReg.CreateKey HKEY_CURRENT_USER,strKeyPath
			Case "HKLM","HKEY_LOCAL_MACHINE"
				oReg.CreateKey HKEY_LOCAL_MACHINE,strKeyPath
			Case "HKU","HKEY_USERS"
				oReg.CreateKey HKEY_USERS,strKeyPath
			Case "HKCC","HKEY_CURRENT_CONFIG"
				oReg.CreateKey HKEY_CURRENT_CONFIG,strKeyPath	
			Case Else
				Exit Function
			
		End Select
		
		'Abort on failure to create the object
	    If Err.Number > 0 Then
	        Err.Clear
	        On Error Goto 0
	        Exit Function
	    End If
    
	End Function
	
	''' <summary>
    ''' Delete a key from the registry of any WMI enabled computer.
    ''' </summary>
    ''' <param name="strRegPath" type="string">a full registry key path, e.g.HKLM\SOFTWARE\Microsoft\DirectX OR HKEY_CLASSES_ROOT\SOFTWARE\Microsoft\DirectX
    ''' <returns></returns>    
    ''' <remarks></remarks>
	Public Function DeleteRegKey(ByVal strRegPath)
			
		Dim EndStrPos, Length, Hive, strKeyPath
		
		'Initiate custom error handling
    	On Error Resume Next
    	
		EndStrPos = Instr(strRegPath,"\")
		Length = EndStrPos  - 1
		Hive = Mid(strRegPath,1,Length)
		strKeyPath = Mid(strRegPath,EndStrPos + 1)
		
		select Case UCase(Hive)
		
			Case "HKCR","HKEY_CLASSES_ROOT"
				oReg.DeleteKey HKEY_CLASSES_ROOT,strKeyPath
			Case "HKCU","HKEY_CURRENT_USER"
				oReg.DeleteKey HKEY_CURRENT_USER,strKeyPath
			Case "HKLM","HKEY_LOCAL_MACHINE"
				oReg.DeleteKey HKEY_LOCAL_MACHINE,strKeyPath
			Case "HKU","HKEY_USERS"
				oReg.DeleteKey HKEY_USERS,strKeyPath
			Case "HKCC","HKEY_CURRENT_CONFIG"
				oReg.DeleteKey HKEY_CURRENT_CONFIG,strKeyPath	
			Case Else
				Exit Function
			
		End Select
		
		'Abort on failure to create the object
	    If Err.Number > 0 Then
	        Err.Clear
	        On Error Goto 0
	        Exit Function
	    End If
    
	End Function
	
	''' <summary>
    ''' Delete a key from the registry of any WMI enabled computer.
    ''' </summary>
    ''' <param name="strRegPath" type="string">a full registry key path, e.g.HKLM\SOFTWARE\Microsoft\DirectX OR HKEY_CLASSES_ROOT\SOFTWARE\Microsoft\DirectX
    ''' <param name="strValueName" type="string">the value name to be deleted, e.g.InstalledVersion
    ''' <returns></returns>    
    ''' <remarks></remarks>
	Public Function DeleteRegValue(ByVal strRegPath, ByVal strValueName)
			
		Dim EndStrPos, Length, Hive, strKeyPath
		
		'Initiate custom error handling
    	On Error Resume Next
    	
		EndStrPos = Instr(strRegPath,"\")
		Length = EndStrPos  - 1
		Hive = Mid(strRegPath,1,Length)
		strKeyPath = Mid(strRegPath,EndStrPos + 1)
		
		select Case UCase(Hive)
		
			Case "HKCR","HKEY_CLASSES_ROOT"
				oReg.DeleteValue HKEY_CLASSES_ROOT,strKeyPath, strValueName
			Case "HKCU","HKEY_CURRENT_USER"
				oReg.DeleteValue HKEY_CURRENT_USER,strKeyPath, strValueName
			Case "HKLM","HKEY_LOCAL_MACHINE"
				oReg.DeleteValue HKEY_LOCAL_MACHINE,strKeyPath, strValueName
			Case "HKU","HKEY_USERS"
				oReg.DeleteValue HKEY_USERS,strKeyPath, strValueName
			Case "HKCC","HKEY_CURRENT_CONFIG"
				oReg.DeleteValue HKEY_CURRENT_CONFIG,strKeyPath, strValueName	
			Case Else
				Exit Function
			
		End Select
		
		'Abort on failure to create the object
	    If Err.Number > 0 Then
	        Err.Clear
	        On Error Goto 0
	        Exit Function
	    End If
    
	End Function
	
End Class

Public Function RegLib()
	
	Dim objRegLib
	Set objRegLib = New ClsRegLib
	Set RegLib = objRegLib

End Function

'##################################################################################
''' Example
'##################################################################################

'Dim strRegPath, strValueName, RegValueData, RegDataType

'strRegPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\test"
'RegLib.CreateRegKey(strRegPath)
'RegLib.DeleteRegKey(strRegPath)

'strRegPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{241F2BF7-69EB-42A4-9156-96B2426C7504}" 
'strValueName = "test REG_SZ"
'RegValueData = "test"
'RegDataType = REG_SZ
'call RegLib.SetRegValue(strRegPath, strValueName,RegValueData, RegDataType)

'strValueName = "test REG_EXPAND_SZ"
'RegValueData = "http://go.microsoft.com/fwlink/?LinkId=81488"
'RegDataType = REG_EXPAND_SZ
'call RegLib.SetRegValue(strRegPath, strValueName,RegValueData, RegDataType)

'strValueName = "test REG_BINARY"
'RegValueData = "&H01,&Ha2"
'RegDataType = REG_BINARY
'call RegLib.SetRegValue(strRegPath, strValueName,RegValueData, RegDataType)

'strValueName = "test REG_DWORD"
'RegValueData = 2
'RegDataType = REG_DWORD
'call RegLib.SetRegValue(strRegPath, strValueName,RegValueData, RegDataType)

'strValueName = "test REG_MULTI_SZ"
'RegValueData = "a,b,c"
'RegDataType = REG_MULTI_SZ
'call RegLib.SetRegValue(strRegPath, strValueName,RegValueData, RegDataType)

'strRegPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{241F2BF7-69EB-42A4-9156-96B2426C7504}" 
'strValueName = "test REG_SZ"
'MsgBox RegLib.ReadRegValue(strRegPath, strValueName)
'Call RegLib.DeleteRegValue(strRegPath, strValueName)

'strValueName = "test REG_EXPAND_SZ"
'MsgBox RegLib.ReadRegValue(strRegPath, strValueName)
'Call RegLib.DeleteRegValue(strRegPath, strValueName)

'strValueName = "test REG_BINARY"
'MsgBox RegLib.ReadRegValue(strRegPath, strValueName)
'Call RegLib.DeleteRegValue(strRegPath, strValueName)

'strValueName = "test REG_DWORD"
'MsgBox RegLib.ReadRegValue(strRegPath, strValueName)
'Call RegLib.DeleteRegValue(strRegPath, strValueName)

'strValueName = "test REG_MULTI_SZ"
'MsgBox RegLib.ReadRegValue(strRegPath, strValueName)
'Call RegLib.DeleteRegValue(strRegPath, strValueName)






