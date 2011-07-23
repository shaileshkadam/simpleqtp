Option Explicit

''' #########################################################
''' <summary>
''' A Library to work with array
''' </summary>
''' <remarks></remarks>	 
''' #########################################################

Class ClsArrayLib

	''' <summary>
    ''' Return the length of an array at a specified dimension
    ''' </summary>
    ''' <param name="arrValue" type="array">Array to be evaluated</param>
    ''' <param name="dimension" type="array">The dimension to be evaluated</param>
    ''' <return></return>
    ''' <remarks></remarks>
	Public Function GetArrDimLength(ByVal arrValue, ByVal dimension)
	
		GetArrDimLength = UBound(arrValue, dimension) + 1
		
	End Function

	''' <summary>
    ''' Get the dimension of the array
    ''' </summary>
    ''' <param name="arrValue" type="array">Array to be evaluated</param>
    ''' <return>The dimension of the array</return>
    ''' <remarks></remarks>
	Public Function GetArrayDimension(ByVal arrValue)
	
		Dim tmp
		Dim haveDim : haveDim = False
		Dim i : i = 1
	
		On Error Resume Next
			While Not haveDim
				tmp = UBound(arrValue, i)
				If Err.Number <> 0 Then
					haveDim = True
					i = i - 1
				Else
					i = i + 1
				End If
			Wend
			Err.Clear()
			GetArrayDimension = i
		On Error GoTo 0
		
	End Function
	
	''' <summary>
    ''' Return the length of an array
    ''' </summary>
    ''' <param name="arrValue" type="array">Array to be evaluated</param>
    ''' <return>Array containing the lengths of each dimension in order</return>
    ''' <remarks></remarks>
	Public Function getArrLength(ByVal arrValue)
	
		Dim arrLen(), maxDim, i
		maxDim = getArrayDimension(arrValue)
		If maxDim = 0 Then
			ReDim arrLen(0)
			arrLen(0) = 0
		Else
			ReDim arrLen(maxDim - 1)

			For i = 0 To maxDim - 1
				arrLen(i) = getArrDimLength(arrValue, i+1)
			Next
		End If
		getArrLength = arrLen
		
	End Function

	''' <summary>
    ''' Determine whether or not a given entry exists in the array
    ''' </summary>
    ''' <param name="arrValue" type="array">Array to be evaluated</param>
    ''' <param name="strEntry" type="string">Array to be evaluated</param>
    ''' <param name="useCase" type="bool">Indicator on whether or not to use case when performing the comparison</param>
    ''' <return>True/False</return>
    ''' <remarks></remarks>
	Public Function isEntryInArray(ByVal arrValue, ByVal strEntry, ByVal useCase )
	
		Dim arrLen : arrLen = getArrLength(arrValue)
		Select Case UBound(arrLen)
			Case 0
				Dim i
				isEntryInArray = False
				For i = 0 To arrLen(0) - 1
					If useCase Then
						If arrValue(i) = strEntry Then
							isEntryInArray = True
							i = arrLen(0)
						End If
					Else
						If UCase(arrValue(i)) = UCase(strEntry) Then
							isEntryInArray = True
							i = arrLen(0)
						End If
					End If
				Next
			Case 1
				Dim j
				isEntryInArray = False
				For i = 0 To arrLen(0) - 1
					For j = 0 To arrLen(1) - 1
						If useCase Then
							If arrValue(i, j) = strEntry Then
								isEntryInArray = True
								i = arrLen(0)
								j = arrLen(1)
							End If
						Else
							If UCase(arrValue(i, j)) = UCase(strEntry) Then
								isEntryInArray = True
								i = arrLen(0)
								j = arrLen(1)
							End If
						End If
					Next
				Next
			Case Else
				isEntryInArray = -1
		End Select
		
	End Function

	''' <summary>
    ''' Convert array to string
    ''' </summary>
    ''' <param name="arrValue" type="array">Array to be evaluated</param>
    ''' <param name="delimiter" type="string">Separate the entries with delimiter</param>
    ''' <return>The converted array as a string</return>
    ''' <remarks></remarks>
	Public Function ConvertArrToString(ByVal arrValue, ByVal delimiter)
	
		Dim retStr : retStr = ""
		Dim i, j, k, arrDim, arrLen
	
		arrDim = getArrayDimension(arrValue)
		arrLen = getArrLength(arrValue)
	
		Select Case arrDim
			Case 1
				For i = 0 To arrLen(0) - 1
					If i = 0 Then
						retStr = retStr & arrValue(i)
					Else
						retStr = retStr & "," & arrValue(i)
					End If
				Next
			Case 2
				For i = 0 To arrLen(0) - 1
					For j = 0 To arrLen(1) - 1
						If j = 0 Then
							retStr = retStr & arrValue(i, j)
						Else
							retStr = retStr & "," & arrValue(i, j)
						End If
					Next
					If i = arrLen(0)-1 Then
						retStr = retStr & ""
					else
						retStr = retStr & ","
					End If
				Next
			Case Else
				MsgBox "Unhandled array dimension (" & arrDim & ")", vbCritical, "Exception in ConvertArrToString"
		End Select
	
		ConvertArrToString = retStr
		
	End Function
	
	''' <summary>
    ''' Convert array to Dictionary
    ''' </summary>
    ''' <param name="arrValue" type="array">Array to be evaluated</param>
    ''' <return>The converted array as a dictionary</return>
    ''' <remarks></remarks>
	Public Function ConvertArrToDictionary(ByVal arrValue)
		
		Dim oDict
		Dim i, j, arrDim, arrLen
		Set oDict = CreateObject("Scripting.Dictionary")
		arrDim = getArrayDimension(arrValue)
		arrLen = getArrLength(arrValue)
		Select Case arrDim
			Case 0
				' Nothing to do
			Case 1
				' Assume keys are to be numbers starting from 0
				For i = 0 To arrLen(0) - 1
					oDict.Add i, arrValue(i)
				Next
			Case 2
				For i = 0 To arrLen(0) - 1
					For j = 0 To arrLen(1) - 1
						oDict.Add arrValue(i, j), arrValue(i, j)
					Next
				Next
			Case 3
				' Don't think that this makes sense.
				' Leaving as TODO until determine what makes most sense.
			Case Else
				MsgBox "Unhandled array dimension (" & arrDim & ")", vbCritical, "Exception in ConvertArrToDictionary"
		End Select
	
		If oDict.Count > 0 Then
			Set ConvertArrToDictionary = oDict
		Else
			ConvertArrToDictionary = ""
		End If
		
	End Function

	
	''' <summary>
    ''' Removing Duplicate Values From A List Of Values
    ''' </summary>
    ''' <param name="arrValue" type="array">The arr to be removed</param>
    ''' <return>removed arr</return>
    ''' <remarks></remarks>
	Public Function RemoveDupValueFromArr(ByVal arrValue)
		
		Dim arrUnique
		'Create a dictionary object
		Set oDict = CreateObject("Scripting.Dictionary")
		'CompareMode allows to set how the values need to be compared. 
		'Setting it to vbTextCompare mode means that the keys are incasesensitive. 
		'If we want to treat values as case sensitive then we need to use vbTextBinary for CompareMode
		oDict.CompareMode = vbTextCompare
		 
		'Loop through each value and add the items to dictionary
		'all the duplicate values only overwrite the existing values and we are left with only the unique values
		For each sValue in arrValue
		    oDict(sValue) = sValue
		Next
		
		'access these values using the items property
		RemoveDupValueFromArr = oDict.Items
		
	End Function
	
	''' <summary>
    ''' Array sort in alphabetical order
    ''' </summary>
    ''' <param name="arrSortIn" type="array">The array to be sorted</param>
    ''' <return></return>
    ''' <remarks></remarks>
	Public Function SortArrInAlph(Byval arrSortIn)
		
		Dim i, j, temp
		for i = UBound(arrSortIn) - 1 To 0 Step -1
		    for j= 0 to i
		        if arrSortIn(j)>arrSortIn(j+1) then
		            temp=arrSortIn(j+1)
		            arrSortIn(j+1)=arrSortIn(j)
		            arrSortIn(j)=temp
		        end if
		    Next
		Next
		SortArrInAlph = arrSortIn
		
	End Function
	
	''' <summary>
    ''' Convert a string to an array
    ''' </summary>
    ''' <param name="str" type="string">String to be converted to the array</param>
    ''' <param name="delimiter" type="string">string delimiter</param>
    ''' <return>Array</return>
    ''' <remarks></remarks>
	Public Function ConvertStringToArray(ByVal str, ByVal delimiter)
		
		ConvertStringToArray = Split(str, delimiter)
		
	End Function

End Class

Public Function ArrayLib()
	
	Set ArrayLib = New ClsArrayLib

End Function	
	
	