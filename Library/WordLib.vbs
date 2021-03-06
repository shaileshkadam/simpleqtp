Option Explicit

''' #########################################################
''' <summary>
''' A Library to work with word
''' </summary>
''' <remarks></remarks>	
''' <example>

''' Dim sDocPath : sDocPath = "C:\test.doc"
''' Dim oWordLib : set oWordLib = WordLib

''' -------------------------Load word data-----------------------------
'''	Example 1
''' oWordLib.LoadFile(sDocPath)
'''	Example 2
''' MsgBox oWordLib.LoadFile(sDocPath).ExtractAllTexts()

''' -------------------------Insert the specific text------------------
''' oWordLib.InsertText(Date())

''' -------------------------Insert the specific image-----------------
''' oWordLib.InsertImage("C:\Blue hills.jpg")

''' -------------------------Search specific string--------------------
''' MsgBox oWordLib.IsStringFound("insert", True, false)
''' MsgBox oWordLib.IsStringFound("insert", False, false)
''' MsgBox oWordLib.IsStringFound("*sert*", True, true)
''' MsgBox oWordLib.IsStringFound("*sert*", True, false)

''' -------------------------Replace specific string--------------------
''' oWordLib.ReplaceString "a", "bbb"

''' -------------------------Set the page orientation------------------
''' oWordLib.SetOrientationToLandscape
''' oWordLib.SetOrientationToPortrait

''' -------------------------Covert-----------------------------
''' oWordLib.ConvertDOCtoHTMLPDF sDocPath, "pdf"
''' oWordLib.ConvertDOCtoHTMLPDF sDocPath, "HTML"

''' set oWordLib = nothing

''' </example> 
''' #########################################################

'A global Word.Application instance
Public oWord

Class ClsWordLib
	
	''' <summary>
    ''' Region Word.Application instance created in Class_Initialize
    ''' </summary>
    ''' <remarks></remarks>
	Private oDocApp
	
	''' <summary>
    ''' Region Word instance created in LoadFile
    ''' </summary>
    ''' <remarks></remarks>
    ''' <seealso>LoadFile()</seealso>
	Private oDoc
	
	''' <summary>
    ''' Word Document path
    ''' </summary>
    ''' <remarks></remarks>
	Private sDoc
	
	Public Property Get Doc
		Set Doc = oDoc
	End Property
	
	Public Property Set Doc(ByVal val)
		Set oDoc = val
	End Property
	
	''' <summary>
    ''' Class Initialization procedure. Creates Word Singleton.
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Initialize()
		
		Dim bCreated : bCreated = False
		
		If IsObject(oWord) Then
			If Not oWord Is Nothing Then
				If TypeName(oWord) = "Application" Then
					bCreated = True
				End If
			End If
		End If
		
		If Not bCreated Then 
			On Error Resume Next
				Set oWord = GetObject("",  "Word.Application")

				If Err.Number <> 0 Then
					Err.Clear

					Set oWord = CreateObject("Word.Application")
					oWord.DisplayAlerts = 0
				End If
				
				If Err.Number <> 0 Then
					MsgBox "Please install Word before using WordLib", vbOKOnly, "Word.Application Exception!"
					Err.Clear
					Exit Sub
				End If
			On Error Goto 0
		End If
		Set oDocApp = oWord
		Set Me.Doc = Nothing
		
	End Sub
	
	''' <summary>
    ''' Class Termination procedure
    ''' </summary>
    ''' <remarks></remarks>
	Private Sub Class_Terminate()
	
	 	oDocApp.Quit
		Set oDocApp = Nothing
		
		If IsObject(Me.Doc) Then
			If Not Me.Doc Is Nothing Then
				Set Me.Doc = Nothing
			End If
		End If
	
	End Sub
	
	''' <summary>
    ''' Sets the region instances for Word.
	''' NOTE: For any method to execute, LoadFile must be executed first to set the Word.
    ''' </summary>
    ''' <param name="sDocPath" type="string">Path to the Doc</param>
    ''' <returns>WordLib</returns>
	Public Property Get LoadFile(ByVal sDocPath)
		
		Dim fso
		Set LoadFile = Me
		If oDocApp Is Nothing Then Exit Property

		If sDoc = "" Then sDoc = sDocPath
	
		If sDoc <> sDocPath Then
			Me.Doc.Close
			sDoc = sDocPath
		End If

		
		On Error Resume Next

			Set fso = CreateObject("Scripting.FileSystemObject")

			If Not fso.FileExists(sDocPath) Then
				MsgBox "Unable to find the Word with the given path: " & _
					sDocPath, vbOKOnly, "WordFile.LoadFile->'File Not Found' Exception!"
				Set fso = Nothing
				Exit Property
			End If

			Set Me.Doc = oDocApp.Documents.Open(sDocPath)
			
			If Err.Number <> 0 Then
				MsgBox "Unable to load the WorkBook: " & sDocPath, vbOKOnly, _
					"LoadFile->'xlsApp.WorkBooks.Open(WorkBook)' Exception!"
				Err.Clear
				Exit Property
			End If
			
		On Error Goto 0

	End Property
	
	''' <summary>
    ''' Sets the page orientation in Microsoft Word to landscape
    ''' </summary>
	Public Function SetOrientationToLandscape()
		
		Const wdOrientLandscape = 1
		Me.Doc.PageSetup.Orientation = wdOrientLandscape
		Me.Doc.Save
		
	End Function
	
	''' <summary>
    ''' Sets the page orientation in Microsoft Word to Portrait
    ''' </summary>
	Public Function SetOrientationToPortrait()
		
		Const wdOrientPortrait = 0
		Me.Doc.PageSetup.Orientation = wdOrientPortrait
		Me.Doc.Save
		
	End Function

	''' <summary>
    ''' Creates and saves a new Doc for a given path
    ''' </summary>
    ''' <param name="NewWordPath" type="string">Path of the Doc file</param>
    ''' <param name="bReplaceOldFile" type="Bool">if overwrite</param>
    ''' Notes: Please refer to http://support.microsoft.com/kb/973904
    ''' <remarks></remarks>
	Public Function CreateNewDoc(ByVal NewDocPath, ByVal bReplaceOldFile)
		
		Dim fso, DocWord, selection
		Set fso = CreateObject("Scripting.FileSystemObject")
		If fso.FileExists(NewDocPath) Then
			If bReplaceOldFile Then
				fso.DeleteFile(NewDocPath)
			Else
				Exit Function
			End If
		End If
		
		Set DocWord = oDocApp.Documents.Add()
		oDocApp.Selection.TypeText ""
		oDocApp.ActiveDocument.SaveAs NewDocPath
		LoadFile NewDocPath
		Set fso = nothing
		
	End Function
	
	''' <summary>
    ''' Extract all texts from the Doc
    ''' </summary>
    ''' <remarks></remarks>
	Public Function ExtractAllTexts()

		Dim nWords, n, sText
		sText = ""
		nWords = Doc.Words.count
		For n = 1 to nWords
		     sText = sText & Doc.Words(n)
		Next
		ExtractAllTexts = sText
		
	End Function
	
	''' <summary>
    ''' Check the specific string can be found
    ''' </summary>
    ''' <param name="sSearchString" type="string">string to be searched</param>
    ''' <param name="bMatchCase" type="Bool">if match case</param>
    ''' <param name="bMatchWildcards" type="Bool">if match wildcards</param>
    ''' <returns>Bool</returns>
    ''' <remarks></remarks>
	Public Function IsStringFound(ByVal sSearchString, ByVal bMatchCase, ByVal bMatchWildcards)
		
		Dim n, startRange, endRange, tRange
		IsStringFound = false
		For n = 1 To Doc.Paragraphs.Count
			startRange = Doc.Paragraphs(n).Range.Start
			endRange = Doc.Paragraphs(n).Range.End
			Set tRange = Doc.Range(startRange, endRange)
			tRange.Find.Text = sSearchString
			tRange.Find.MatchCase = bMatchCase
			tRange.Find.MatchWildcards = bMatchWildcards
			tRange.Find.Execute
			If tRange.Find.Found Then
				IsStringFound = True
				Exit Function
			End If
		Next
	
	End Function
	
	''' <summary>
    ''' Replace the specific string
    ''' </summary>
    ''' <param name="sSearchString" type="string">string to be searched</param>
    ''' <param name="sReplaceString" type="string">string to be replaced</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function ReplaceString(ByVal sSearchString, ByVal sReplaceString)
				
		Const wdReplaceAll  = 2
		Dim objSelection
		Set objSelection = oDocApp.Selection	
		objSelection.Find.Text = sSearchString
		objSelection.Find.Forward = TRUE
		'objSelection.Find.MatchWholeWord = TRUE
		objSelection.Find.Replacement.Text = sReplaceString
		objSelection.Find.Execute ,,,,,,,,,,wdReplaceAll
		Me.Doc.Save
		Set objSelection = Nothing
	
	End Function
	
	''' <summary>
    ''' Insert the specific text into word document
    ''' </summary>
    ''' <param name="sText" type="string">text to be inserted</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function InsertText(ByVal sText)
	
		Const END_OF_STORY = 6
		Const MOVE_SELECTION = 0
		Dim selection
		Set selection = oDocApp.Selection
		selection.EndKey END_OF_STORY, MOVE_SELECTION
		selection.TypeParagraph()
		'selection.Font.Size = "14"
		selection.TypeText "" & sText
		selection.TypeParagraph()
		'selection.Font.Size = "10"
		Me.Doc.Save
		Set selection = Nothing

	End Function
	
	''' <summary>
    ''' Insert the specific image into word document
    ''' </summary>
    ''' <param name="sImagePath" type="string">image to be inserted</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function InsertImage(ByVal sImagePath)

    	Const END_OF_STORY = 6
		Const MOVE_SELECTION = 0
		Dim selection, oImg
		Set selection = oDocApp.Selection
		selection.EndKey END_OF_STORY, MOVE_SELECTION
		selection.TypeParagraph()
	    With selection
	        Set oImg = .InlineShapes.AddPicture(sImagePath, False, True)
	        oImg.Width = oImg.Width*1
	        oImg.Height = oImg.Height*1
	        'Center alignment
	        oImg.Range.ParagraphFormat.Alignment = 1
	        .TypeParagraph
	    End With
	    Me.Doc.Save
		Set selection = Nothing
    
   end Function
   
   	''' <summary>
    ''' Convert the word document to HTML/PDF
    ''' </summary>
    ''' <param name="SrcFile" type="string">word document to be converted</param>
    ''' <param name="sFormat" type="string">"PDF" or "HTML"</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
   Public Function ConvertDOCtoHTMLPDF(ByVal SrcFile, ByVal sFormat)
	   
	   Dim fso, oFile, FilePath, myFile
	   Set fso = CreateObject( "Scripting.FileSystemObject" )
	   Set oFile = fso.GetFile(SrcFile)
	   FilePath = oFile.Path
	   oDocApp.Documents.Open FilePath
	   If ucase(sFormat) = "PDF" Then
	       myFile = fso.BuildPath (ofile.ParentFolder ,fso.GetBaseName(ofile) & ".pdf")
	       oDocApp.Activedocument.Saveas myfile, 17
	   elseIf ucase(sFormat) = "HTML" Then
	       myFile = fso.BuildPath (ofile.ParentFolder ,fso.GetBaseName(ofile) & ".html")
	       oDocApp.Activedocument.Saveas myfile, 8
	   End If
	   Set fso = Nothing
	   Set oFile = Nothing
	   
	End Function

End Class

Public Function WordLib()
	
	Dim obj
	Set obj = New ClsWordLib
	Set WordLib = obj

End function


