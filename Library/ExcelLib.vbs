Option Explicit

''' #########################################################
''' <summary>
''' A library to work with excel
''' </summary>
''' <remarks></remarks>
''' <example>

'''	Dim xlsFile : xlsFile = "C:\Config.xls"
'''	Dim oExcelLib
'''	Set oExcelLib = ExcelLib

''' -------------------------Load excel data-----------------------------
'''	Example 1
'''	oExcelLib.LoadFile xlsFile, 1
'''	Example 2
'''	Dim var
'''	var = oExcelLib.LoadFile(xlsFile, 1).GetCellValue(2, "B")
'''	MsgBox var

''' -------------------------GetCellValue--------------------------------
'''	Example 1
'''	MsgBox "GetCellValue: " & oExcelLib.GetCellValue(1, 1)
'''	Example 2
'''	MsgBox "GetCellValue: " & oExcelLib.GetCellValue(2, 1)
'''	Example 3
'''	MsgBox "GetCellValue: " & oExcelLib.GetCellValue(1, "A")
'''	Example 4
'''	MsgBox "GetCellValue: " & oExcelLib.GetCellValue(2, "A")

''' -------------------------Get2DArrayFromSheet--------------------------
'''	Dim arr
'''	arr = oExcelLib.Get2DArrayFromSheet
'''	MsgBox "(Get2DArrayFromSheet) UBound: " & UBound(arr)
'''	MsgBox "(Get2DArrayFromSheet) UBound: " & UBound(arr, 2)
'''	MsgBox "(Get2DArrayFromSheet) UBound: " & LBound(arr)
'''	MsgBox "(Get2DArrayFromSheet) UBound: " & LBound(arr, 2)

''' -------------------------GetWorkSheetRange----------------------------
'''	Dim rng
'''	Set rng = oExcelLib.GetWorkSheetRange
'''	Msgbox "(GetWorkSheetRange) TypeName: " & TypeName(rng)
'''	MsgBox "(GetWorkSheetRange) Rows: " & rng.Rows.Count
'''	MsgBox "(GetWorkSheetRange) Columns: " & rng.Columns.Count

''' -------------------------BuildRowHeadingDictionary-------------------
'''	Dim Dict
'''	Set Dict = oExcelLib.BuildRowHeadingDictionary(2, 1)
'''	MsgBox "(BuildRowHeadingDictionary) Count: " & Dict.Count
'''	MsgBox "BuildRowHeadingDictionary) UserName: " & Dict("UserName")
'''	MsgBox "BuildRowHeadingDictionary) Password: " & Dict("Password")
'''	Dict.RemoveAll
'''	Set Dict = Nothing

''' -------------------------FindCellContainingValue----------------------
'''	Dim cell
'''	Set cell = oExcelLib.FindCellContainingValue("UserName")
'''	MsgBox "(FindCellContainingValue) Row: " & cell.Row
'''	MsgBox "(FindCellContainingValue) Column: " & cell.Column

''' -------------------------FindNextCell----------------------------------
'''	Set cell = oExcelLib.FindNextCell
'''	MsgBox "(FindNextCell) Row: " & cell.Row
'''	MsgBox "(FindNextCell) Column: " & cell.Column
'''	Set cell = oExcelLib.FindNextCell
'''	MsgBox "(FindCellContainingValue) Row: " & cell.Row
'''	MsgBox "(FindNextCell) Column: " & cell.Column
'''	Set cell = Nothing

''' -------------------------GetUsedRowCount--------------------------------
'''	MsgBox "GetUsedRowCount: " & oExcelLib.GetUsedRowCount

''' -------------------------GetUsedColumnCount-----------------------------
'''	MsgBox "GetUsedColumnCount: " & oExcelLib.GetUsedColumnCount

''' -------------------------GetUsedRowCountByColumn------------------------
'''	Example 1
'''	MsgBox "GetUsedRowCountByColumn: " & oExcelLib.GetUsedRowCountByColumn(1)
'''	Example 2
'''	MsgBox "GetUsedRowCountByColumn: " & oExcelLib.GetUsedRowCountByColumn(2)

''' -------------------------GetUsedColumnCountByRow------------------------
'''	Example 1
'''	MsgBox "GetUsedColumnCountByRow: " & oExcelLib.GetUsedColumnCountByRow(1)
'''	Example 2
'''	MsgBox "GetUsedColumnCountByRow: " & oExcelLib.GetUsedColumnCountByRow(2)

''' -------------------------WriteCellValue---------------------------------
'''	Example 1
'''	oExcelLib.WriteCellValue "SSN", 1, 7
'''	Example 2
'''	oExcelLib.WriteCellValue "123-456-6789", 2, "G"

''' ------------------------InsertComment-----------------------------------
'''	Example 1
'''	oExcelLib.InsertComment "This is a comment", true, 2, 3
'''	Example 2
'''	oExcelLib.InsertComment "This is a comment", false, 4, 5

''' -----------------------ChangeCellFontColor------------------------------
'''	oExcelLib.ChangeCellFontColor 36, 3, 1

''' -----------------------ChangeCellBGColor--------------------------------
'''	oExcelLib.ChangeCellBGColor 3, 2, 3

''' -----------------------ChangeFontSize-----------------------------------
'''	oExcelLib.ChangeFontSize 24, 1, 2

''' -----------------------DrawBorder---------------------------------------
'''	Example 1
'''	oExcelLib.DrawBorder "B2", "left"
'''	Example 2
'''	oExcelLib.DrawBorder "B2", "right"
'''	Example 3
'''	oExcelLib.DrawBorder "B2", "top"
'''	Example 4
'''	oExcelLib.DrawBorder "B2", "bottom"

''' ------------------------MergeCells--------------------------------------
'''	oExcelLib.MergeCells "A1:A2"

''' ------------------------UnmergeCells------------------------------------
'''	oExcelLib.UnmergeCells "A1:A2"

''' -----------------------Convert to CSV formate---------------------------
'''	oExcelLib.ConvertoCSV xlsFile

''' -----------------------CreateNewWorkBook--------------------------------
'''	Example 1
'''	Call oExcelLib.CreateNewWorkBook("c:\New1.xls", true)
'''	Example 2
'''	Call oExcelLib.CreateNewWorkBook("c:\New2.xls", false)

''' -----------------------AddWorkSheet-------------------------------------
'''	Example 1	
'''	oExcelLib.AddWorkSheet xlsFile, "test1"
'''	Example 2
'''	oExcelLib.AddWorkSheet xlsFile, "test2"

''' -----------------------InsertImageInCell---------------------------------
'''	oExcelLib.InsertImageInCell 4, 1, "C:\qtp.jpg"

''' -----------------------Release the excel object--------------------------
'''	Set oExcelLib = Nothing
''' </example>
''' #########################################################

'A global Excel.Application instance
Public oExcel

Class ClsExcelLib

'Private Variables
	
	''' <summary>
    ''' Range object created in FindCellContainingValue and passed to FindNextCell
    ''' </summary>
    ''' <remarks></remarks>
	Private rngFound
	
	''' <summary>
    ''' Region Excel.Application instance created in Class_Initialize
    ''' </summary>
    ''' <remarks></remarks>
	Private xlsApp
	
	''' <summary>
    ''' Region Excel WorkBook instance created in LoadFile
    ''' </summary>
    ''' <remarks></remarks>
    ''' <seealso>LoadFile()</seealso>
	Private oxlsbook
	
	''' <summary>
    ''' Region Excel WorkSheet instance created in LoadFile
    ''' </summary>
    ''' <remarks></remarks>
	Private oxlssheet
	
	''' <summary>
    ''' WorkBook path
    ''' </summary>
    ''' <remarks></remarks>
	Private sWorkBook
	
	''' <summary>
    ''' WorkSheet name
    ''' </summary>
    ''' <remarks></remarks>
	Private sWorkSheet


'Public Properties
	
	Public Property Get xlsBook
		Set xlsBook = oxlsbook
	End Property
	
	Public Property Set xlsBook(ByVal val)
		Set oxlsbook = val
	End Property	
		
	Public Property Get xlsSheet
		Set xlsSheet = oxlssheet
	End Property
	
	Public Property Set xlsSheet(ByVal val)
		Set oxlssheet = val
	End Property	


'Private Methods

	''' <summary>
    ''' Class Initialization procedure. Creates Excel Singleton.
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Initialize()
		
		Dim bCreated : bCreated = False
		
		If IsObject(oExcel) Then
			If Not oExcel Is Nothing Then
				If TypeName(oExcel) = "Application" Then
					bCreated = True
				End If
			End If
		End If
		
		If Not bCreated Then 
			On Error Resume Next
				Set oExcel = GetObject("", "Excel.Application")

				If Err.Number <> 0 Then
					Err.Clear

					Set oExcel = CreateObject("Excel.Application")
				End If
				
				If Err.Number <> 0 Then
					MsgBox "Please install Excel before using ExcelLib", vbOKOnly, "Excel.Application Exception!"
					Err.Clear
					Exit Sub
				End If
			On Error Goto 0
		End If
		Set xlsApp = oExcel
		Set Me.xlsBook = Nothing
		Set Me.xlsSheet = Nothing
		
	End Sub
	
	''' <summary>
    ''' Class Termination procedure
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Terminate()
	
	 	xlsApp.Quit
		Set xlsApp = Nothing
		
		If IsObject(Me.xlsBook) Then
			If Not Me.xlsBook Is Nothing Then
				Set Me.xlsBook = Nothing
			End If
		End If
		
		If IsObject(Me.xlsSheet) Then
			If Not Me.xlsSheet Is Nothing Then
				Set Me.xlsSheet = Nothing
			End If
		End If
	End Sub

	''' <summary>
    ''' Sets the region instances for Excel WorkBook and WorkSheet. These instances for the
	''' Excel source are created only once and used by other methods.
	''' NOTE: For any method to execute, LoadFile must be executed first to set the WorkBook and WorkSheet.
    ''' </summary>
    ''' <param name="WorkBook" type="string">Path to the Excel WorkBook</param>
    ''' <param name="WorkSheet" type="string">Name or Item Number of the WorkSheet</param>
    ''' <returns>instance of the ExcelLib class</returns>
	Public Property Get LoadFile(ByVal WorkBook, ByVal WorkSheet)
		
		Dim fso
		Set LoadFile = Me
		If xlsApp Is Nothing Then Exit Property
		
		'c#: this.sWorkBook = WorkBook;
		'vb: Me.sWorkBook = WorkBook
		If sWorkBook = "" Then sWorkBook = WorkBook
		'c#: this.sWorkSheet = WorkSheet;
		'vb: Me.sWorkSheet = WorkSheet
		If sWorkSheet = "" Then sWorkSheet = WorkSheet

		If sWorkBook <> WorkBook Then
			Me.xlsBook.Close

			sWorkBook = WorkBook
		End If

		If sWorkSheet <> WorkSheet Then
			sWorkSheet = WorkSheet
		End If
		
		On Error Resume Next

			Set fso = CreateObject("Scripting.FileSystemObject")

			If Not fso.FileExists(WorkBook) Then
				MsgBox "Unable to find the Excel WorkBook with the given path: " & _
					WorkBook, vbOKOnly, "ExcelFile.LoadFile->'File Not Found' Exception!"
				Set fso = Nothing
				Exit Property
			End If

			Set Me.xlsBook = xlsApp.WorkBooks.Open(WorkBook)
			
			If Err.Number <> 0 Then
				MsgBox "Unable to load the WorkBook: " & WorkBook, vbOKOnly, _
					"LoadFile->'xlsApp.WorkBooks.Open(WorkBook)' Exception!"
				Err.Clear
				Exit Property
			End If
			
			If Not IsNumeric(WorkSheet) Then
				Set Me.xlsSheet = Me.xlsBook.WorkSheets(WorkSheet)
			Else
				Set Me.xlsSheet = Me.xlsBook.WorkSheets.Item(WorkSheet)
			End If
			
			If Err.Number <> 0 Then
				MsgBox "Unable to bind to the WorkSheet: " & WorkSheet, vbOKOnly, _
					"ExcelLib.LoadFile->'xlsApp.WorkBooks.WorkSheets(Sheet)' Exception!"
				Err.Clear
				Exit Property
			End If
			
		On Error Goto 0
	End Property

	''' <summary>
    ''' Returns a Scripting.Dictionary object with heading & row pair.
    ''' </summary>
    ''' <param name="iRow" type="integer">Data Row</param>
	''' <param name="iHeadingRow" type="integer">Heading Row</param>
    ''' <returns>Scripting.Dictionary</returns>
	Public Property Get BuildRowHeadingDictionary(ByVal iRow, ByVal iHeadingRow)
		
		Dim oRange, arrRange, iColumns, dic, iCol
		Set oRange = GetWorkSheetRange
		arrRange = oRange.Value
		
		iColumns = UBound(oRange.Value, 2)
		
		Set dic = CreateObject("Scripting.Dictionary")
		dic.CompareMode = vbTextCompare
		
		For iCol = LBound(arrRange, 2) To UBound(arrRange, 2)
			If Not dic.Exists(arrRange(1, iCol)) Then
				dic.Add CStr(arrRange(iHeadingRow, iCol)), CStr(arrRange(iRow, iCol))
			End If
		Next
		
		Set BuildRowHeadingDictionary = dic
	End Property

	''' <summary>
    ''' Reads the value of a cell in an Excel WorkSheet
    ''' </summary>
    ''' <param name="iRow" type="integer">Row number</param>
    ''' <param name="vColumn" type="variant">Column letter or number</param>
    ''' <returns>String</returns>
	Public Property Get GetCellValue(ByVal iRow, ByVal vColumn)
		GetCellValue = Me.xlsSheet.Cells(iRow, vColumn).Value
	End Property

	''' <summary>
    ''' Returns the complete WorkSheet Range object
    ''' </summary>
    ''' <returns>Range</returns>
	Public Property Get GetWorkSheetRange()
		Set GetWorkSheetRange = Me.xlsSheet.UsedRange
	End Property
	
	''' <summary>
    ''' Returns a 2D array from the WorkSheet
    ''' </summary>
    ''' <returns>Array</returns>
	Public Property Get Get2DArrayFromSheet()
		Get2DArrayFromSheet = GetWorkSheetRange.Value
	End Property

	''' <summary>
    ''' Returns a Range object if the supplied argument is found in the WorkSheet
    ''' </summary>
    ''' <param name="arg" type="variant">Value being searched for</param>
    ''' <returns>Range</returns>
	Public Property Get FindCellContainingValue(ByVal arg)
		
		Dim cell
		Set cell = Me.xlsSheet.UsedRange.Find(arg)
		'c#: this.rngFound = cell;
		'vb: Me.rngFound = cell
		Set rngFound = cell
		Set FindCellContainingValue = cell
		
	End Property

	''' <summary>
    ''' Finds the next cell from the supplied argument in FindCellContainingValue
    ''' </summary>
    ''' <returns>Range</returns>
	''' <seealso>FindCellContainingValue</seealso>
	Public Property Get FindNextCell()
		
		Dim cell
		Set cell = Me.xlsSheet.UsedRange.FindNext(rngFound)
		Set rngFound = cell
		Set FindNextCell = cell
		
	End Property
	
	''' <summary>
    ''' Finds the number of used rows in the Excel WorkSheet
    ''' </summary>
    ''' <returns>Integer</returns>
	Public Property Get GetUsedRowCount()
		GetUsedRowCount = Me.xlsSheet.UsedRange.Rows.Count
	End Property
	
	''' <summary>
    ''' Finds the number of used columns in the Excel WorkSheet
    ''' </summary>
    ''' <returns>Integer</returns>
	Public Property Get GetUsedColumnCount()
		GetUsedColumnCount = Me.xlsSheet.UsedRange.Columns.Count
	End Property
	
	''' <summary>
    ''' Finds the number of used rows in an Excel WorkSheet by column
    ''' </summary>
    ''' <param name="vColumn" type="variant">Column letter or number</param>
    ''' <returns>Integer</returns>
	Public Property Get GetUsedRowCountByColumn(ByVal vColumn)
		Const xlDown = -4121
		
		GetUsedRowCountByColumn = Me.xlsSheet.Cells(1, vColumn).End(xlDown).Row
	End Property
	
	''' <summary>
    ''' Finds the number of used columns in an Excel WorkSheet by row
    ''' </summary>
    ''' <param name="iRow" type="integer">Row number</param>
    ''' <returns>Integer</returns>
	Public Property Get GetUsedColumnCountByRow(ByVal iRow)
		Const xlToRight = -4161
		
		GetUsedColumnCountByRow = Me.xlsSheet.Cells(iRow, 1).End(xlToRight).Column
	End Property


'Public Methods
	
	''' <summary>
    ''' Inputs a value to an Excel cell
    ''' </summary>
    ''' <param name="iRow" type="integer">Value input</param>
    ''' <param name="vColumn" type="variant">Row number</param>
    ''' <param name="TheValue" type="variant">Column letter or number</param>
    ''' <remarks></remarks>
	Public Function WriteCellValue(ByVal TheValue, ByVal iRow, ByVal vColumn)
		If TheValue = "" Then Exit Function
		
		Me.xlsSheet.Cells(iRow, vColumn).Value = TheValue
		Me.xlsBook.Save
	End Function

	''' <summary>
    ''' Inserts an image in a Excel cell
    ''' </summary>
    ''' <param name="iRow" type="integer">Row number</param>
    ''' <param name="vColumn" type="variant">Column letter or number</param>
    ''' <param name="ImagePath" type="string">Path to the image file</param>
    ''' <remarks></remarks>
	Public Function InsertImageInCell(ByVal iRow, ByVal vColumn, ByVal ImagePath)
		
		Dim fso, pic
		Set fso = CreateObject("Scripting.FileSystemObject")
		If Not fso.FileExists(ImagePath) Then
			MsgBox "Unable to find the Image  with the given path: " & _
				ImagePath & ".", vbOKOnly, "ExcelLib.InsertImageInCell->'File Not Found' Exception!"
			Set fso = Nothing
			Exit Function
		End If
			
		Me.xlsSheet.Cells(iRow, vColumn).Select

		With Me.xlsSheet
			Set pic = .Pictures.Insert(ImagePath)

			With .Cells(iRow, vColumn)
				pic.top = .Top
				pic.left = .Left
			
				pic.ShapeRange.height = .RowHeight * 1
				pic.ShapeRange.width = .ColumnWidth * .ColumnWidth
			End With
		End With

		oxlsbook.Save
	End Function
	
	''' <summary>
    ''' Changes the background color of a cell
    ''' </summary>
    ''' <param name="ColorCode" type="variant">Value of the custom color</param>
	''' <param name="iRow" type="integer">Row number</param>
    ''' <param name="vColumn" type="variant">Column letter or number</param>
    ''' <remarks></remarks>
	Public Function ChangeCellBGColor(ByVal ColorCode, ByVal iRow, ByVal vColumn)
		Me.xlsSheet.Cells(iRow, vColumn).Interior.ColorIndex = ColorCode
		Me.xlsBook.Save
	End Function

	''' <summary>
    ''' Changes the font color of a cell
    ''' </summary>
    ''' <param name="ColorCode" type="variant">Value of the custom color</param>
	''' <param name="iRow" type="integer">Row number</param>
    ''' <param name="vColumn" type="variant">Column letter or number</param>
    ''' <remarks></remarks>
	Public Function ChangeCellFontColor(ByVal ColorCode, ByVal iRow, ByVal vColumn)
		Me.xlsSheet.Cells(iRow, vColumn).Font.ColorIndex = ColorCode
		Me.xlsBook.Save
	End Function
	
	''' <summary>
    ''' Changes the font size
    ''' </summary>
    ''' <param name="iFontSize" type="integer">New font size</param>
	''' <param name="iRow" type="integer">Row number</param>
    ''' <param name="vColumn" type="variant">Column letter or number</param>
    ''' <remarks></remarks>
	Public Function ChangeFontSize(ByVal iFontSize, ByVal iRow, ByVal vColumn)
		Me.xlsSheet.Cells(iRow, vColumn).Font.Size = iFontSize
		Me.xlsBook.Save
	End Function
	
	''' <summary>
    ''' Draws a border to the left, right, top, or bottom of a given range
    ''' </summary>
    ''' <param name="Range" type="range">Excel Range</param>
    ''' <param name="Direction" type="variant">Direction: left, right, top, bottom</param>
    ''' <remarks></remarks>
	Public Function DrawBorder(ByVal Range, ByVal Direction)
 		If IsNumeric(Direction) Then Direction = CStr(Direction)

		Direction = LCase(Direction)
		
		With Me.xlsSheet.Range(Range)
			Select Case Direction
				Case "1", "left"
					.Borders(1).LineStyle = 1
				Case "2", "right"
					.Borders(2).LineStyle = 1
				Case "3", "top"
					.Borders(3).LineStyle = 1
				Case "4", "bottom"
					.Borders(4).LineStyle = 1
				Case "5", "all"
					Dim ix					
					For ix = 1 To 4
						.Borders(ix).LineStyle = 1
					Next
				Case Else
					MsgBox "Invalid Direction: ' " & Direction & " '" & vbNewLine & _
						"Please provide the correct Direction to draw the border." & _
						Direction, vbOKOnly, "DrawBorder->'Invalid Direction' Exception!"
					Exit Function
			End Select
		End With

		Me.xlsBook.Save
	End Function
	
	''' <summary>
    ''' Merges the cells in a range
    ''' </summary>
    ''' <param name="Range" type="range">Excel Range</param>
    ''' <remarks></remarks>
	Public Function MergeCells(ByVal Range)
 		
 		xlsApp.DisplayAlerts = False
		Me.xlsSheet.Range(Range).MergeCells = True
		xlsApp.DisplayAlerts = True
		Me.xlsBook.Save
		
	End Function

	''' <summary>
    ''' Removes the merge feature from cells of a given range
    ''' </summary>
    ''' <param name="Range" type="range">Excel Range</param>
    ''' <remarks></remarks>
	Public Function UnmergeCells(ByVal Range)
	
 		xlsApp.DisplayAlerts = False
		Me.xlsSheet.Range(Range).MergeCells = False
		xlsApp.DisplayAlerts = True
		Me.xlsBook.Save
		
	End Function

	''' <summary>
    ''' Inserts a hidden or visible comment in a cell
    ''' </summary>
    ''' <param name="CommentText" type="variant">Comment text</param>
    ''' <param name="iRow" type="integer">Row number</param>
    ''' <param name="vColumn" type="variant">Column letter or number</param>
    ''' <param name="bMakeVisible" type="bool">Make the comment visible or hidden</param>
    ''' <remarks></remarks>
	Public Function InsertComment(ByVal CommentText, ByVal bMakeVisible, ByVal iRow, ByVal vColumn)
		
		With Me.xlsSheet.Cells(iRow, vColumn)
			If Not .Comment Is Nothing Then .Comment.Delete
			
			.AddComment CommentText
			.Comment.Visible = bMakeVisible
		End With

		Me.xlsBook.Save
		
	End Function
	
	''' <summary>
    ''' Creates and saves a new WorkBook for a given path
    ''' </summary>
    ''' <param name="WorkBookPath" type="string">Path of the Excel file</param>
    ''' <remarks></remarks>
	Public Function CreateNewWorkBook(ByVal WorkBookPath, ByVal bReplaceOldFile)
		
		Dim fso, xlsBook
		Set fso = CreateObject("Scripting.FileSystemObject")
		If fso.FileExists(WorkBookPath) Then
			If bReplaceOldFile Then
				fso.DeleteFile(WorkBookPath)
			Else
				Exit Function
			End If
		End If
		
		Set xlsBook = xlsApp.Workbooks.Add
		xlsBook.SaveAs WorkBookPath
		
	End Function
	
	''' <summary>
    ''' Adds a WorkSheets to a given WorkBook
    ''' </summary>
    ''' <param name="WorkBook" type="string">Path to the Excel file</param>
    ''' <param name="WorkSheetName" type="string">New WorkSheet name</param>
    ''' <remarks></remarks>
	Public Function AddWorkSheet(ByVal WorkBook, ByVal WorkSheetName)
		
		Dim fso, xlsBook, xlsSheet
		Set fso = CreateObject("Scripting.FileSystemObject")
		If Not fso.FileExists(WorkBook) Then
			MsgBox "Unable to find the Excel WorkBook with the given path: " & _
				WorkBook, vbOKOnly, "NewWorkSheet->'File Not Found' Exception!"
			Set fso = Nothing
			Exit Function
		End If
		
		Set xlsBook = xlsApp.Workbooks.Open(WorkBook)
		
		For Each xlsSheet in xlsBook.WorkSheets
			If LCase(xlsSheet.Name) = LCase(WorkSheetName) Then Exit Function
		Next
		
		Set xlsSheet = xlsBook.Worksheets.Add
		xlsSheet.Name = WorkSheetName
		xlsBook.Save
	End Function
	
	''' <summary>
    ''' Convert excel to CSV file
    ''' </summary>
    ''' <param name="WorkBook" type="string">Path to the Excel file</param>
    ''' <remarks></remarks>
	Public Function ConvertoCSV (Byval WorkBook)
		
		Dim blnForceOverwrite
		Dim objFSO
		Dim strFileOut
		
		Const xlCSV = 6
		blnForceOverwrite = True
	
		Set objFSO   = CreateObject( "Scripting.FileSystemObject" )
		
		With objFSO
			If .FileExists( WorkBook ) Then
				WorkBook = .GetAbsolutePathName( WorkBook )
				xlsApp.Workbooks.Open WorkBook, , True
				strFileOut = .BuildPath( .GetParentFolderName( WorkBook ), .GetBaseName( WorkBook ) & ".csv" )
				If blnForceOverwrite And .FileExists( strFileOut ) Then
					.DeleteFile strFileOut, True
					WScript.Echo "Existing CSV file replaced."
				End If
				On Error Resume Next
				xlsApp.ActiveWorkbook.SaveAs strFileOut, xlCSV
				If Err.Number = 1004 Then
					WScript.Echo "Existing CSV file not replaced."
				End If
				On Error Goto 0
				xlsApp.ActiveWorkbook.Close False
			End If
		End With
		
		Set objFSO   = Nothing
    End Function

	''' <summary>
    ''' Closes the WorkBook opened in LoadFile
    ''' </summary>
    ''' <remarks></remarks>
	Public Function CloseWorkBook()
		On Error Resume Next
			oxlsbook.Close

			If Err.Number <> 0 Then Err.Clear
		On Error Goto 0
	End Function

End Class
'=======================================================================================================
'Any class declaration is private in QTP. So use a code like "Set objMyClass = new MyClass" only within the function library where it is declared.
'To workaround that you need to define a class constructor function returning an instance to the class.
'=======================================================================================================
Public  Function  ExcelLib()
  	Set ExcelLib = New ClsExcelLib
End Function

	