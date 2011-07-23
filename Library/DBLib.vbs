Option Explicit

''' #########################################################
''' <summary>
''' The library to work with DB
''' </summary>
''' <remarks></remarks>
''' <example>
''' Dim oDBLib, sqlQuery
''' Dim Array_2D
''' Set oDBLib = DBLib

''' ------Initial the Login account information ------------------
''' oDBLib.Init "a106403", "QWERTY5", "O06SLD1"

''' sqlQuery = "select * from spj.daily_G1_claim_payments"

''' ------Export DB query result to 2D Array ---------------------
''' Array_2D = oDBLib.ExportDBTo2DArray(sqlQuery)

''' ------Export DB query result to excel sheet ------------------
''' oDBLib.ExportDBToExcel "C:\test.xls", sqlQuery
''' </example>
''' #########################################################

Class ClsDBLib
	
	Private xlsApp
	Private sUserID
	Private sPassword
	Private sDataSource
	Private cnObj
	Private rsObj
	
	Public Property Get UserID
		UserID = sUserID
	End Property
	
	Public Property let UserID(ByVal val)
		sUserID = val
	End Property
	
	Public Property Get Password
		Password = sPassword
	End Property
	
	Public Property let Password(ByVal val)
		sPassword = val
	End Property
	
	Public Property Get DataSource
		DataSource = sDataSource
	End Property
	
	Public Property let DataSource(ByVal val)
		sDataSource = val
	End Property
	
	''' <summary>
    ''' Class Initialization procedure
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Initialize()
	
		Set xlsApp = CreateObject("Excel.Application")
			
	End Sub
	
	''' <summary>
    ''' Class Termination procedure
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Terminate()
	
	 	xlsApp.Quit
		Set xlsApp = Nothing
		Set cnObj = Nothing
		Set rsObj = Nothing
		
	End Sub
	
	''' <summary>
    ''' Initial the Login information
    ''' </summary>
    ''' <param name="strUserID" type="string">The User ID to log in the DB</param>
    ''' <param name="strPassword" type="string">The Password to log in the DB</param>
    ''' <param name="strDataSource" type="string">The DataSource to log in the DB</param>
    ''' <returns></returns>
	Public Function Init(ByVal strUserID, ByVal strPassword, ByVal strDataSource)
		
		Me.UserID = strUserID
		Me.Password = strPassword
		Me.DataSource = strDataSource
	
	End Function
	
	''' <summary>
    ''' Create a connection to a oracle
    ''' </summary>
    ''' <returns>true/false</returns>
	Public Function CreateConnection()	
		
		Dim cnString, retVal
		' Create database connection object
		Set cnObj = CreateObject("ADODB.Connection")
	
		' If rsObj is set to False then don't create a record set
		' else create the record set object
		If Not rsObj Then
			Set rsObj = CreateObject("ADODB.Recordset")
		End If
	
		' Create the connection. Report error if there is a problem
		On Error Resume Next
		' if connection type is Basic, then "Data Source" should be 'Hostname:Port/Service name'
		' if conncetion type is TNS, then "Data Source" should be 'Network Alias'
		cnString = "Provider=MSDAORA.1;User ID=" & Me.UserID & ";Password=" & Me.Password & ";Data Source=" & Me.DataSource
		cnObj.Open(cnString)
		If Err.Number <> 0 Then
			'MsgBox "Error on database connection" & vbNewLine & "ERROR: " & _
			'		Err.Description & vbNewLine & "CONNECTION: " & cnString, _
			'		vbCritical, "Database Connection Error"
			retVal = false
		Else
			retVal = true
		End If
		On Error GoTo 0
	
		CreateConnection = retVal
		
	End Function
	
	''' <summary>
    ''' Close the connection to the oracle
    ''' </summary>
    ''' <returns>true/false</returns>
	Public Function CloseConnection()
	
		' Close the record set object if it exists
		If IsObject(rsObj) Then
			On Error Resume Next
				rsObj.Close
				Set rsObj = Nothing
			On Error GoTo 0
		End If
	
		' Close the connection object
		cnObj.Close
		Set cnObj = Nothing
	
		' Return True
		CloseConnection = true
	
	End Function
	
	''' <summary>
    ''' Count the number count of the records
    ''' </summary>
    ''' <param name="rsObj" type="Object">Record set object</param>
    ''' <returns>The number count of the records</returns>
	Public Function NoRecs(ByVal rsObj)
	
		Dim count: count = 1
		Do Until rsObj.EOF
			count = count + 1
			rsObj.MoveNext
		Loop
		NoRecs = count
		
	End Function
	
	''' <summary>
    ''' Export DB query result to 2D Array
    ''' </summary>
    ''' <param name="sqlQuery" type="string">SQL statement to be executed</param>
    ''' <returns>2D Array</returns>
	Public Function ExportDBTo2DArray(ByVal sqlQuery)
		
		Dim retVal
		Dim Array_2D
		If CreateConnection Then
			' Perform the query
			On Error Resume Next
				'rsObj.CursorLocation = adUseClient
				'msgbox sqlQuery
				rsObj.Open sqlQuery, cnObj ', adOpenStatis, adLockOptimistic
	
				If Err.Number <> 0 Then
					' MsgBox "Error on database query" & vbNewLine & "ERROR: " & _
					'		 Err.Description & vbNewLine & "SQL: " & sqlQuery, vbCritical, _
					'		 "Database Query Error"
					retVal = false
				Else
					retVal = true
				End If
			On Error GoTo 0
	
			' Put the results into the return array
			' Array structure:
			'	row    = record from database
			'	column = column from database
			' First index = row
			' Second index = column
			' Zero entry in column = table column name
			If retVal Then
				'MsgBox sqlQuery & " = " & rsObj.RecordCount & " records and " & rsObj.Fields.Count & " fields"
				If (rsObj.RecordCount = -1) Then
					ReDim Array_2D(NoRecs(rsObj) - 1, rsObj.Fields.Count - 1)
				Else
					ReDim Array_2D(rsObj.RecordCount, rsObj.Fields.Count - 1)
				End If
				Dim j, i
				j = 1
				If rsObj.BOF Then 
					'MsgBox "No records returned for the query" & vbNewLine & "SQL: " & sqlQuery
				Else
					rsObj.MoveFirst
					While ((Not rsObj.EOF) And (Not rsObj.BOF))
						For i = 0 To rsObj.Fields.Count - 1
							Array_2D(0, i) = rsObj.Fields(i).Name
							Array_2D(j, i) = rsObj.Fields(i).Value
						Next
		
						rsObj.MoveNext
						j = j + 1
					Wend
				End If
			End If
	
			' Close the database connections
			CloseConnection
			
		End If
		
		ExportDBTo2DArray = Array_2D
	
	End Function
	
	''' <summary>
    ''' Export DB query result to excel sheet
    ''' </summary>
    ''' <param name="sWorkbook" type="string">Path to the Excel WorkBook</param>
    ''' <param name="sqlQuery" type="string">SQL statement to be executed</param>
    ''' <returns></returns>
	Public Function ExportDBToExcel(ByVal sWorkbook, ByVal sqlQuery)
		
		Dim objWorkbook1, objWorksheet1, FirstCell, strCon, xlQuery
		Set objWorkbook1= xlsApp.Workbooks.Open(sWorkbook)
		Set objWorksheet1 = objWorkbook1.Worksheets(1)
		FirstCell = objWorksheet1.Cells(1,1).Address(0,0)
		' if connection type is Basic, then "Data Source" should be 'Hostname:Port/Service name'
		' if conncetion type is TNS, then "Data Source" should be 'Network Alias'
		strCon = "Provider=MSDAORA.1;User ID=" & Me.UserID & ";Password=" & Me.Password & ";Data Source=" & Me.DataSource	
		Set xlQuery = objWorksheet1.QueryTables.Add("OLEDB;" & strCon, objWorksheet1.Range(FirstCell), sqlQuery)
		xlQuery.Refresh
		objWorkbook1.Save
		objWorkbook1.Close
		Set objWorkbook1 = Nothing
		Set objWorksheet1 = Nothing
		Set xlQuery = nothing
	
	End Function
	
End class	

Public Function DBLib()

  	Set DBLib = New ClsDBLib
  	
End Function


