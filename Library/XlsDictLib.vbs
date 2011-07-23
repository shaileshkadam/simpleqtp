Option Explicit

''' #########################################################
''' <summary>
''' A Library to create a high-effective TestData Dictionary based on excel
''' </summary>
''' <remarks></remarks>	 
''' <example>
''' Dim sWorkBook, vSheet, mDict
''' sWorkBook = "C:\Config.xls"
''' vSheet = "DB"              
''' Set mDict = XlsDictLib.Load(sWorkBook, vSheet, "SGO_UXT")
''' MsgBox mDict("UserName")    
''' MsgBox mDict("Password")    
''' MsgBox mDict("Data Source")     
''' Set mDict = XlsDictLib.Load(sWorkBook, vSheet, 4)
''' MsgBox mDict("UserName")    
''' MsgBox mDict("Password")    
''' MsgBox mDict("Data Source")  
''' </example>
''' #########################################################

Class ClsXlsDictLib

 	Private mDict       'Local Instance of Scripting.Dictionary
	Private strWorkBook    'Excel WorkBook
	Private strSheet       'Excel WorkSheet
	Private iCurrentRow     'Excel Current Row 
    Private strUniquecol   'Unique column value to access the specfic row
    
	''' <summary>
	''' Create a dictionary object
	''' </summary>
	''' <remarks></remarks>	 
    Private Sub Class_Initialize
  
        Set mDict = CreateObject("Scripting.Dictionary")
        
    End Sub
    
    ''' <summary>
	''' Release the dictionary object
	''' </summary>
	''' <remarks></remarks>	    
    Private Sub Class_Terminate

        Set mDict = Nothing
        
    End Sub
    
	''' <summary>
    ''' Makes the TestData dictionary available to the test
    ''' </summary>
    ''' <param name="sWorkBook" type="string">Path to the Workbook where test data is stored</param>
    ''' <param name="sSheet" type="string">Name of the Worksheet where the data is stored</param>
    ''' <param name="sUniquecolVal_Or_iRow" type="int/string">Specfic row/Unique column value to access the specfic row</param>
    ''' <returns>Object- Scripting.Dictionary</returns>
    ''' <remarks></remarks>
	Public Default Function Load(sWorkBook, sSheet, sUniquecolVal_Or_iRow)
 		
 		With Me
			.WorkBook = sWorkBook
			.Sheet = sSheet
			.UniquecolVal_Or_iRow = sUniquecolVal_Or_iRow
		End With
		
		Set Load = BuildContext
		
	End Function

	''' <summary>
	''' Do core operation of building the test data dictionary
	''' </summary>
	''' <remarks></remarks>	 
	Private Function BuildContext
		Dim oConn, oRS, x, sQuery
		CONST adOpenStatic = 3
		CONST adLockOptimistic = 3
		CONST adCmdText = "&H0001"

		Set oConn = CreateObject("ADODB.Connection")
		Set oRS = CreateObject("ADODB.RecordSet")

		'Open Connection
		'Connection strings for Excel http://www.connectionstrings.com/excel
		'"HDR=Yes;" indicates that the first row contains columnnames, not data. "HDR=No;" indicates the opposite.
		'"IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. Note that this option might affect excel sheet write access negative.
		oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" &_
					"Data Source=" & Me.WorkBook & ";" & _
					"Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"

		'Query
		'How To Use ADO with Excel Data from Visual Basic or VBA
		'http://support.microsoft.com/kb/257819/EN-US
		sQuery = "Select * From [" & Me.Sheet & "$]"

		'Run query against WorkBook
		'Open Method (ADO Recordset) http://msdn.microsoft.com/en-us/library/ms675544(v=VS.85).aspx
		oRS.Open sQuery, oConn, 3, 3, 1
		
		'Increment x depending upon the column
		'COlumn 1 in Excel: x = 0
		'Column 2 in Excel: x = 1
		' ... and so on ...
		'Check if the parameter is iRow or Uniquecol
		oRS.MoveFirst
		If bIsNumber(Me.UniquecolVal_Or_iRow) Then
			'Move RecordSet to the target Row
			For iCurrentRow = 2 to Me.UniquecolVal_Or_iRow - 1
				oRS.MoveNext
			Next
		Else
			'Unique column locate at column 1
			x = 0
			iCurrentRow = 2
			Do Until oRS.EOF
				If Not IsNull(oRS.Fields(x)) then
					If Trim(LCase(CStr(oRS.Fields(x)))) = LCase(CStr(Me.UniquecolVal_Or_iRow)) Then
						Exit Do
					End If
				End If
				iCurrentRow = iCurrentRow + 1
				oRS.MoveNext
			Loop
		End If


		'Use a For..Loop to Build Scripting.Dictionary
		For x = 0 to oRS.Fields.Count - 1
			With mDict
				.Add "" & oRS(x).Name, "" & oRS.Fields(x)
			End With
		Next
		
		Set oRS = nothing
		Set oConn = Nothing
		
		Set BuildContext = mDict
		
	End Function
	
	''' <summary>
	''' check if a string is a number
	''' </summary>
	''' <remarks></remarks>
	Public Function bIsNumber(sInput)
	
		Dim myRegExp
		Set myRegExp = New RegExp
		myRegExp.Pattern = "^\d+?$"
		If myRegExp.Test(sInput) Then
			bIsNumber = True
		Else
			bIsNumber = False
		End If
	    Set myRegExp = Nothing
	    
	End Function
	
	Public Property Get WorkBook
		WorkBook = strWorkBook 
	End Property

	Public Property Let WorkBook(Byval Val)
		strWorkBook = Val
	End Property	
	

	Public Property Let Sheet(Byval Val)
		strSheet = Val
	End Property	

	Public Property Get Sheet
		Sheet = strSheet 
	End Property	
	

	Public Property Let UniquecolVal_Or_iRow(Byval Val)
		strUniquecol = Val
	End Property	

	Public Property Get UniquecolVal_Or_iRow
		UniquecolVal_Or_iRow = strUniquecol 
	End Property	
	
End Class

''' <summary>
''' Any class declaration is private in QTP. So use a code like "Set objMyClass = new MyClass" only within the function library where it is declared
''' To workaround that you need to define a class constructor function returning an instance to the class.
''' </summary>
''' <remarks></remarks>
Public Function XlsDictLib()

	Set XlsDictLib = New ClsXlsDictLib
	
End Function 