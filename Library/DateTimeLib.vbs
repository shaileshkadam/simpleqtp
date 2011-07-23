Option Explicit

''' #########################################################
''' <summary>
''' A library to work with DateTime
''' </summary>
''' <remarks></remarks>	 
''' #########################################################

Class ClsDateTimeLib
	
	''' <summary>
    ''' Format date and time
    ''' </summary>
    ''' <param name="sDateTime" type="string">DateTime</param>
    ''' <param name="sFormat" type="string">
	''' D - Single digit date (if possible), DD - Double digit date,
	''' DDD - Abreviated day name(Ex - Mon), DDDD - Complete Day Name (Ex - Sunday)
	''' DDDDD - short format date m/d/yyyy (Ex - 3/1/2007)
	''' DDDDDD - Long format date dddd, mmmm dd, yyyy (Ex - Thursday, March 12, 2007)
	''' M - Single digiti month (if possible), 
	''' MM - double digit month, MMM - abbreviated month (Ex - Jan), 
	''' MMMM - Complete Month Name (Ex - January)
	''' YY - 2 digit year, YYYY - Complete year
	''' H - Single digit hour( if possible), HH - two digit hours
	''' M - Single digit minute (if possible), MM - double digit minute. M/MM is only treated as Minute in case it just next to a H/HH tag
	''' S - Single digit Second, SS - double digit seconds
	''' AM/PM - Display time in 12 hrs format and display AM/PM whichever is applicable, AMPM - Same as AM/PM
	''' A/P - Display time in 12 hrs format and display A/P  whichever is applicable instead of AM/PM
    ''' </param>
    ''' <returns>formated date</returns>
    ''' <remarks></remarks>
    ''' <example>
	''' Displays Wednesday, 10 December 2008 00:43 AM
	''' MsgBox ReFormatDateTime(Now, "dddd, dd mmmm yyyy, hh:mm AMPM")
	''' Generate a unique number using currrent date and time
	''' MsgBox ReFormatDateTime(Now, "DDMMYYYYHHMMSS")
	''' Displays 01-Mar-2007
	''' MsgBox ReFormatDateTime("3/1/2007", "dd-MMM-yyyy") 
	''' Displays 03/1/2007
	''' MsgBox ReFormatDateTime("3/1/2007", "mm/d/yyyy") 
	''' Displays 1:31 PM
	''' MsgBox ReFormatDateTime("13:31", "H:MM AM/PM") 
	''' Displays 1:31 P
	''' MsgBox ReFormatDateTime("13:31", "H:MM A/P") 
    ''' </example>
	Function ReFormatDateTime(ByVal sDateTime, ByVal sFormat)
		'Boolean tag to identify if MM tag is to be interpreted as for Month or Minute.
		'Minute should only in case the last tag was an H i.e. hour tag
		Dim isMMTime
		
		'Array for full day and month names
		Dim sDays, sMonths
		
		'Array for storing values of various tags. valD 
		'stores values for D, DD, DDD, DDDD and DDDDD.
		Dim valD, valH, valTM, valDM, valY, valS
		
		'AM/PM String in case time is to be displayed in a 12 hour format
		Dim timeAMPM
		
		'Length of the current tag
		Dim iLenFormat
		'Loop control for current Tag
		Dim iTag
		'Current Tag
		Dim curTag
		'Tag value
		Dim sTagValue
		'Operation array. Contains array which needs to be picked from 
		'valS, valY, valDM, valTM, valH and valD depending on the Tag
		Dim optArray
		
		'Always assume MM as Month until unless a H or HH tag appears
		isMMTime = False
		'Convert the date time to standard date time format
		sDateTime = CDate(sDateTime)
		'Convert the format to upper case
		sFormat = UCase(sFormat)
		
		ReFormatDateTime = ""
		'Populate the array with name of days and months
		'Note: The first element of all the arrays will be kept blank 
		'so that index starts from 1 for referal instead of 0
		sDays = Array("", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
		sMonths = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
		
		'Values of H, HH
		valH = Array("", _
		                 Hour(sDateTime), _
		                 Hour(sDateTime) _
		             )
		'Values for M, MM in case of time
		valTM = Array("", _
		                Minute(sDateTime), _
		                Minute(sDateTime) _
		              )
		
		'Values for S, SS
		valS = Array("", _
		               Second(sDateTime), _
		               Second(sDateTime) _
		             )
		timeAMPM = ""
		
		'Values for D, DD, DDD, DDDD, DDDDD, DDDDDD
		'in Case of DDDDD display the date in short format which is equivalent to m/d/yyyy
		'in Case of DDDDDD display date in long format which is equivalent to dddd, mmmm dd, yyyy
		valD = Array("", _
		               Day(sDateTime), _
		               Day(sDateTime), _
		               Left(sDays(Weekday(sDateTime)), 3), _
		               sDays(Weekday(sDateTime)), _
		               FormatDateTime(sDateTime, vbShortDate), _
		               FormatDateTime(sDateTime, vbLongDate) _
		             )
		
		'Values for M, MM, MMM and MMMM
		valDM = Array("", _
		                 Month(sDateTime), _
		                 Month(sDateTime), _
		                 Left(sMonths(Month(sDateTime)), 3), _
		                 sMonths(Month(sDateTime)) _
		              )
		
		'Values for Y, YY, YYY and YYYY. Note: we dont want to touch Y and YYYtags
		' so we keep the same value as the tag itself.
		valY = Array("", "Y", _
		                 Right(Year(sDateTime), 2), _
		                 "YYY", _
		                 Year(sDateTime) _
		             )
		
		'Check if AM/PM or A/P is contained the format
		If InStr(sFormat, "AMPM") Or InStr(sFormat, "AM/PM") Or InStr(sFormat, "A/P") Then
		    valH = Array("", Hour(sDateTime), Hour(sDateTime))
		    If valH(1) >= 12 Then
		        valH(1) = valH(1) - 12
		        If valH(1) = 0 Then valH(1) = 12
		        valH(2) = valH(1)
		        timeAMPM = "PM"
		    Else
		        timeAMPM = "AM"
		    End If
		    'Replace AM/PM with AAPP or A/P with AP.
		    'this is necessary because if we dont change it then 
		    'the M tag in AM/PM would be processed
		    'Replace AAPP/AP tags once processing for all other tags is done
		    sFormat = Replace(sFormat, "AM/PM", "AAPP")
		    sFormat = Replace(sFormat, "AMPM", "AAPP")
		    sFormat = Replace(sFormat, "A/P", "AP")
		End If
		
		'Make HH, MM, DD, SS all 2 digits in they are single digit
		If Len(valH(2)) = 1 Then valH(2) = "0" & valH(2)
		If Len(valTM(2)) = 1 Then valTM(2) = "0" & valTM(2)
		If Len(valS(2)) = 1 Then valS(2) = "0" & valS(2)
		If Len(valD(2)) = 1 Then valD(2) = "0" & valD(2)
		If Len(valDM(2)) = 1 Then valDM(2) = "0" & valDM(2)
		
		iLenFormat = Len(sFormat)
		Dim curChar, iLenTag
		
		'Process the format string
		For iTag = 1 To iLenFormat
		    iLenTag = 1
		    'The current tag
		    curChar = Mid(sFormat, iTag, 1)
		    curTag = curChar
		    'Loop while the current tag repeats and stop if the tag changes 
		    'or format length is exhausted
		    While iTag < iLenFormat And curChar = Mid(sFormat, iTag + iLenTag, 1)
		        iLenTag = iLenTag + 1
		    Wend
		    'Increase the loop contorl value and -1 to compensate for the +1 from the loop
		    iTag = iTag + iLenTag - 1
		    'Create the complete tag
		    sTagValue = String(iLenTag, curTag)
		
		    'Convert Tag to its actual value
		    Select Case curTag
		        Case "D", "M", "Y", "H", "S"
		            'Select the array to be operated from
		            Select Case curTag
		                Case "H"
		                    'Set the flag for next immidiate M/MM to be interpreted as Time
		                    isMMTime = True
		                    optArray = valH
		                Case "M"
		                    'M/MM is to be interpreted as Time. 
		                    'In case it is MMM or higher then it has to 
		                    'be interpreted as month only
		                    If isMMTime And iLenTag <= UBound(valTM) Then
		                        optArray = valTM
		                        isMMTime = False
		                    Else
		                        optArray = valDM
		                    End If
		                Case "Y"
		                        isMMTime = False
		                        optArray = valY
		                Case "S"
		                    isMMTime = False
		                    optArray = valS
		                Case "D"
		                    isMMTime = False
		                    optArray = valD
		            End Select
		
		            'If length of tag is greater than max available tag value 
		            'then do the replacement from higher to lower order
		            'i.e. in case of YYYYY tag replace values in order of YYYY 
		            ', then YYY, so on...
		            If iLenTag > UBound(optArray) Then
		                iCount = UBound(optArray)
		                For iIndex = iCount To 1 Step -1
		                    sTagValue = Replace(sTagValue, _
		                                   String(i, curTag), _
		                                   optArray(iIndex))
		                Next
		            Else
		                'Replace the value directly from the val array
		                sTagValue = Replace(sTagValue, _
		                                    String(iLenTag, curTag), _
		                                    optArray(iLenTag))
		            End If
		        Case Else
		            'Do Nothing
		    End Select
		
		    'Append the tag value the current formated date till now
		    ReFormatDateTime = ReFormatDateTime & sTagValue
		Next
		
		'Update any AAPP or AP tag with the actual value
		ReFormatDateTime = Replace(ReFormatDateTime, "AAPP", timeAMPM)
		ReFormatDateTime = Replace(ReFormatDateTime, "AP", Left(timeAMPM, 1))
	
	End Function
	
	''' <summary>
    ''' Get relative date 
    ''' </summary>
    ''' <param name="sDateTime" type="string">DateTime</param>
    ''' <param name="sDateAdd" type="string">Interval</param>
    ''' <returns>new date calculated</returns>
    ''' <remarks>VBScript provided a DateAdd method which can be used to add any type of parameter to the give date</remarks>
    ''' <example>
	''' Msgbox AddDateTime(now, "1 day 2 hours 3 min 4 sec") 
	''' Msgbox AddDateTime("15-Mar-11", "1 year -1 hr")
    ''' </example>
	Function AddDateTime(ByVal sDateTime, ByVal sDateAdd)
		
		'Pattern to scan for our date
		sPattern = "(\d+) *(day)s?|(\d+) *(month)s?|(\d+) *(year)s?|(\d+) *(hr|hour)s?|(\d+) *(min)(?:utes?)?|(\d+) *(sec)(?:onds?)?"
	 
		Dim oRegEx
		Set oRegEx = New RegExp
		oRegEx.Global = True
		oRegEx.IgnoreCase = True
		oRegEx.Pattern = sPattern
		sDateAdd = Trim(sDateAdd)
	 
		'Convert the existing format to date
		sDateTime = CDate(sDateTime)
	 
		Dim iMultiplier
	 
		'Check if we need to add ot subtract
		Select Case Left(sDateAdd, 1)
			Case "-"
				iMultiplier = -1
			Case Else
				iMultiplier = 1
		End Select
	 
		'Find all pattern to scan for keywords
		Set oMatches = oRegEx.Execute(sDateAdd)
	 
		For each oMatch in oMatches
			Dim i: i = 0
	 
			'Since we are matching multiple groups, empty groups will get created
			'in case the group is not found. Ignore these groups
			While IsEmpty(oMatch.Submatches(i))
				i = i + 1
			Wend
	 
			'Keyword are we processing and interval to be used in DateAdd
			Dim sInterval
	        Select Case oMatch.Submatches(i+1)
				Case "day"
					sInterval = "d"
				Case "month"
					sInterval = "m"
				Case "year"
					sInterval = "yyyy"
				Case "hr", "hour"
					sInterval = "h"
				Case "min"
					sInterval = "n"
				Case "sec"
					sInterval = "s"
			End Select
	 
			'Add or subtract the date
			sDateTime = DateAdd(sInterval, iMultiplier * Cint(oMatch.Submatches(i)),sDateTime)
		Next
	 
		'Return the new date calculated
		AddDateTime = sDateTime
		
	End Function

End Class

Public Function DateTimeLib()
	
	Set DateTimeLib = New ClsDateTimeLib

End Function