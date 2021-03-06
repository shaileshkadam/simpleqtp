Option Explicit

''' #########################################################
''' <summary>
''' A Library to work with string
''' </summary>
''' <remarks></remarks>	 
''' #########################################################

Class ClsStringLib

	''' <summary>
    ''' Get the mid string between two strings
    ''' </summary>
    ''' <param name="strSource" type="string">The source string to be extracted</param>
    ''' <param name="StartStr" type="string">start-string</param>
    ''' <param name="EndStr" type="string">end-string</param>
    ''' <return>extracted string</return>
    ''' <remarks></remarks>
    Public Function GetStrBetween(ByVal strSource,ByVal StartStr,ByVal EndStr)
	    
	    Dim StartStrPos, EndStrPos, Length, Res
		StartStrPos = Instr(strSource, StartStr)+Len(StartStr)  
		EndStrPos = Instr(strSource,EndStr)  
		Length = EndStrPos  - StartStrPos   
		Res = Mid(strSource,StartStrPos,Length)  
		GetStrBetween = Res
		
	End Function
	
	''' <summary>
    ''' Extracts the Nth occurence of a substring from a target-string delimited by a from-string and a to-string.
    ''' </summary>
    ''' <param name="strSource" type="string">The source string to be extracted</param>
    ''' <param name="iPos" type="array">Nth occurence of the sFrom string</param>
    ''' <param name="sFrom" type="string">from-string</param>
    ''' <param name="sTo" type="string">to-string</param>
    ''' <return>extracted string</return>
    ''' <remarks></remarks>
	Public Function strExtractN(ByVal strSource, ByVal iPos, ByVal sFrom, ByVal sTo)
	    
	    Dim iLoop, iLen, iPosn, sLine
	
	    sLine = strSource
	    iLen = Len(sFrom)
	    For iLoop = 1 To iPos
	        iPosn = inStr(sLine,sFrom)
	        If iPosn = 0 Then
	            'Print "strExtractN - Failed - From String Not Found - occurence: " & cStr(iLoop)
	            strExtractN = ""
	            Exit Function
	        End If
	        sLine = Mid(sLine,iPosn + iLen)
	    Next
	    iPosn = inStr(sLine,sTo)
	    If iPosn = 0 Then
	        'Print "strExtractN - Failed - To String Not Found"
	        strExtractN = ""
	        Exit Function
	    End If
	    strExtractN = Left(sLine,iPosn-1)
	    
	End Function

	''' <summary>
    ''' tokenize a string with sing multiple token separators
    ''' </summary>
    ''' <param name="TokenString" type="string">The source string to be token</param>
    ''' <param name="SeparatorsArr" type="array">multiple token separators</param>
    ''' <return>tokenized array</return>
    ''' <remarks></remarks>
	Public Function Tokenize(ByVal TokenString, ByVal SeparatorsArr)
	
		Dim NumWords, arr
		NumWords = 0
		
		Dim NumSeps
		NumSeps = UBound(SeparatorsArr)
		
		Do 
			Dim SepIndex, SepPosition
			SepPosition = 0
			SepIndex    = -1
			
			for i = 0 to NumSeps-1
			
				' Find location of separator in the string
				Dim pos
				pos = InStr(TokenString, SeparatorsArr(i))
				
				' Is the separator present, and is it closest to the beginning of the string?
				If pos > 0 and ( (SepPosition = 0) or (pos < SepPosition) ) Then
					SepPosition = pos
					SepIndex    = i
				End If
				
			Next
	
			' Did we find any separators?	
			If SepIndex < 0 Then
	
				' None found - so the token is the remaining string
				redim preserve arr(NumWords+1)
				arr(NumWords) = TokenString
				
			Else
	
				' Found a token - pull out the substring		
				Dim substr
				substr = Trim(Left(TokenString, SepPosition-1))
		
				' Add the token to the list
				redim preserve arr(NumWords+1)
				arr(NumWords) = substr
			
				' Cutoff the token we just found
				Dim TrimPosition
				TrimPosition = SepPosition+Len(SeparatorsArr(SepIndex))
				TokenString = Trim(Mid(TokenString, TrimPosition))
							
			End If	
			
			NumWords = NumWords + 1
		loop while (SepIndex >= 0)
		
		Tokenize = arr
		
	End Function
	
	''' <summary>
    ''' Return a dictionary containing the locations and matched expressions of the string within the source string
    ''' </summary>
    ''' <param name="strSource" type="string">The source to be looked for</param>
    ''' <param name="sExpression" type="string">The string that is being looked for</param>
    ''' <param name="useCase" type="bool">True, case sensitive search, False not</param>
    ''' <param name="Dict" type="Dictionary">dictionary</param>
    ''' <return>
    ''' Dict - The locations and matched strings in a dictionary
    ''' True - The string was found
    ''' False - The string was not found
    ''' </return>
    ''' <remarks></remarks>
    Public Function FindLocationOfString(ByVal strSource, ByVal sExpression, ByVal useCase, ByRef Dict)
	   	
	   	Dim regEx, Match, Matches
	
		' Make sure that Dict is a dictionary object
		Set Dict = CreateObject("Scripting.Dictionary")
	
		' Create the regular expression properties for the compare
	   	Set regEx = New RegExp
	   	regEx.Pattern = sExpression
		If useCase Then
			regEx.IgnoreCase = False
		Else
			regEx.IgnoreCase = True
		End If
	   	regEx.Global = True
	
		' Perform the compare
	   	Set Matches = regEx.Execute(strSource)
	
	   	For Each Match in Matches
			Dict.Add Match.FirstIndex, Match.Value
	   	Next
	
		If Dict.Count > 0 Then
			FindLocationOfString = True
		Else
			FindLocationOfString = False
		End If
		
	End Function

	''' <summary>
    ''' Determine whether or not an expression is within a string
    ''' </summary>
    ''' <param name="strSource" type="string">The source to be looked for</param>
    ''' <param name="sExpression" type="string">The string that is being looked for</param>
    ''' <param name="useCase" type="bool">True, case sensitive search, False not</param>
    ''' <return>true/false</return>
    ''' <remarks></remarks>
	Public Function isInString(ByVal strSource, ByVal sExpression, ByVal useCase)
	
		Dim retVal, regEx
		Set regEx = New RegExp
		regEx.Pattern = sExpression
		If useCase Then
			regEx.IgnoreCase = False
		Else
			regEx.IgnoreCase = True
		End If
		isInString = regEx.Test(strSource)
		
	End Function
		
	''' <summary>
    ''' Return the number of times a string is matched within a string
    ''' </summary>
    ''' <param name="strSource" type="string">The source to be looked for</param>
    ''' <param name="sExpression" type="string">The string that is being looked for</param>
    ''' <param name="useCase" type="bool">True, case sensitive search, False not</param>
    ''' <return>The number of matches</return>
    ''' <remarks></remarks>
	Public Function NumberOfMatchesInString(ByVal strSource, ByVal sExpression, ByVal useCase)
	   	
	   	Dim regEx, Matches
	
		' Create the regular expression properties for the compare
	   	Set regEx = New RegExp
	   	regEx.Pattern = sExpression
		If useCase Then
			regEx.IgnoreCase = False
		Else
			regEx.IgnoreCase = True
		End If
	   	regEx.Global = True
	
		' Perform the compare
	   	Set Matches = regEx.Execute(strSource)
		NumberOfMatchesInString = Matches.Count
		
	End Function
	
	''' <summary>
    ''' Replace the occurrences of a string with another string	
    ''' </summary>
    ''' <param name="strSource" type="string">The source to be looked for</param>
    ''' <param name="sExpression" type="string">The string that is being looked for</param>
    ''' <param name="newExpression" type="string">The string to replace the sExpression</param>
    ''' <param name="useCase" type="bool">True, case sensitive search, False not</param>
    ''' <return></return>
    ''' <remarks></remarks>
	Public Function ReplaceValue(ByRef strSource, ByVal sExpression, ByVal newExpression, ByVal useCase)
	   	
	   	Dim regEx
	
		' Create the regular expression properties for the compare
	   	Set regEx = New RegExp
	   	regEx.Pattern = sExpression
		If useCase Then
			regEx.IgnoreCase = False
		Else
			regEx.IgnoreCase = True
		End If
	   	regEx.Global = True
	
		' Perform the replace
	   	ReplaceValue = regEx.Replace(strSource, newExpression)
	   	
	End Function
	
	''' <summary>
    ''' Reverse the contents of the string
    ''' </summary>
    ''' <param name="strSource" type="string">The source to be reversed</param>
    ''' <return>Reversed string</return>
    ''' <remarks></remarks>
    Public Function ReverseString(ByVal strSource)
		
		ReverseString = StrReverse(strSource)
	
	End Function
	
	''' <summary>
    ''' Return the length of the string
    ''' </summary>
    ''' <param name="strSource" type="string">The source to be counted</param>
    ''' <return>The length of the string</return>
    ''' <remarks></remarks>
	Public Function lengthOfString(ByVal strSource)
	
		lengthOfString = len(theString)
		
	End Function

	''' <summary>
    ''' Check whether a value exists in  a given array
    ''' </summary>
    ''' <param name="str" type="string">The str to be checked</param>
    ''' <param name="arr" type="array">The array source</param>
    ''' <return>true/false</return>
    ''' <remarks></remarks>
	Public Function CheckValueExistsinArray(ByVal str, ByVal arr)
		
		CheckValueExistsinArray=False
		Dim i
		For i = LBound(arr) to UBound(arr)-1
			If LCase(Trim(arr(i))) = LCase(Trim(str)) then
				CheckValueExistsinArray = true
				Exit Function
			else
				CheckValueExistsinArray = false
			end if
		Next
	    
	End Function
	
	''' <summary>
    ''' Check if a string is ENTIRELY in lower case
    ''' </summary>
    ''' <param name="str" type="string">The string to be checked</param>
    ''' <return>true/false</return>
    ''' <remarks></remarks>
	Public Function bStr_CheckIsAllLower(Byval str)
	    
	    Dim oRegEx, cMatches
	    bStr_CheckIsAllLower = True
	    Set oRegEx = New RegExp
	    oRegEx.Pattern = "[^a-z]"
	    oRegEx.IgnoreCase = False
	    Set cMatches = oRegEx.Execute(str)
	    If cMatches.Count > 0 Then
	        bStr_CheckIsAllLower = False
	    End If
	    Set cMatches = Nothing
	    Set oRegEx = Nothing
	    
	End Function
	
	''' <summary>
    ''' Convert a string to lower case
    ''' </summary>
    ''' <param name="strSource" type="string">The string to be converted</param>
    ''' <return>upper case string</return>
    ''' <remarks></remarks>
	Public Function LowerCaseString(ByVal strSource)
		
		LowerCaseString = UCase(strSource)
		
	End Function
	
	''' <summary>
    ''' Check if a string is ENTIRELY in upper case
    ''' </summary>
    ''' <param name="strSource" type="string">The string to be checked</param>
    ''' <return>true/false</return>
    ''' <remarks></remarks>
	Public Function bStr_CheckIsAllUpper(ByVal strSource)
	    
	    Dim oRegEx, cMatches
	    bStr_CheckIsAllUpper = True
	    Set oRegEx = New RegExp
	    oRegEx.Pattern = "[^A-Z]"
	    oRegEx.IgnoreCase = False
	    Set cMatches = oRegEx.Execute(strSource)
	    If cMatches.Count > 0 Then
	        bStr_CheckIsAllUpper = False
	    End If
	    Set cMatches = Nothing
	    Set oRegEx = Nothing
	    
	End Function
	
	''' <summary>
    ''' Convert a string to upper case
    ''' </summary>
    ''' <param name="strSource" type="string">The string to be converted</param>
    ''' <return>upper case string</return>
    ''' <remarks></remarks>
	Public Function UpperCaseString(ByVal strSource)
		
		UpperCaseString = UCase(strSource)
		
	End Function
	
	''' <summary>
    ''' Convert a string to sentence case
    ''' </summary>
    ''' <param name="strSource" type="string">The string to be converted</param>
    ''' <return>The converted string</return>
    ''' <remarks></remarks>
	Public Function CapFirstLowerRest(ByVal strSource)
		
		' Chr(65) = A, Chr(90) = Z
		' Therefore if outside of the range [65,90] must UCase the first character
		Dim strArray, firstChar, remainString, retVal
	
		' Convert the first character if its not a capital
		firstChar = Left(strSource, 1)
		remainString = Right(strSource, Len(strSource) - 1)
		If Asc(firstChar) < 65 Or Asc(firstChar) > 90 Then
			firstChar = UCase(firstChar)
		End If
		remainString = LCase(remainString)
	
		' Rebuild the string
		CapFirstLowerRest = firstChar & remainString
		
	End Function
	
	''' <summary>
    ''' Convert a string to sentence case
    ''' </summary>
    ''' <param name="strSource" type="string">The string to be converted</param>
    ''' <return>The converted string</return>
    ''' <remarks></remarks>
	Public Function SentenceCaseString(ByVal strSource)
	
		SentenceCaseString = CapFirstLowerRest(strSource)

	End Function

	''' <summary>
    ''' Convert a string to title case
    ''' </summary>
    ''' <param name="strSource" type="string">The string to be converted</param>
    ''' <return>The converted string</return>
    ''' <remarks></remarks>
	Public Function TitleCaseString(ByVal strSource)
		
		Dim temp, arrString, i
		arrString = Split(strSource, " ")
		temp = ""
		For i = 0 To UBound(arrString)
			If temp = "" Then
				temp = capFirstLowerRest(arrString(i))
			Else
				temp = temp & " " & CapFirstLowerRest(arrString(i))
			End If
		Next
		titleCaseString = temp
		
	End Function

	''' <summary>
    ''' Check if a string is a number
    ''' </summary>
    ''' <param name="str" type="string">The str to be checked</param>
    ''' <return>true/false</return>
    ''' <remarks></remarks>
	Public Function bIsNumber(ByVal str)
	
		Dim myRegExp
		Set myRegExp = New RegExp
		myRegExp.Pattern = "^\d+?$"
		If myRegExp.Test(str) Then
			bIsNumber = True
		Else
			bIsNumber = False
		End If
	    Set myRegExp = Nothing
	    
	End Function


End Class

Public Function StringLib()
	
	Set StringLib = New ClsStringLib

End Function	