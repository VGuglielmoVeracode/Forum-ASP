<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Forums(TM)
'**  http://www.webwizforums.com
'**                            
'**  Copyright (C)2001-2019 Web Wiz Ltd. All Rights Reserved.
'**  
'**  THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS UNDER LICENSE FROM WEB WIZ LTD.
'**  
'**  IF YOU DO NOT AGREE TO THE LICENSE AGREEMENT THEN WEB WIZ LTD. IS UNWILLING TO LICENSE 
'**  THE SOFTWARE TO YOU, AND YOU SHOULD DESTROY ALL COPIES YOU HOLD OF 'WEB WIZ' SOFTWARE
'**  AND DERIVATIVE WORKS IMMEDIATELY.
'**  
'**  If you have not received a copy of the license with this work then a copy of the latest
'**  license contract can be found at:-
'**
'**  https://www.webwiz.net/license
'**
'**  For more information about this software and for licensing information please contact
'**  'Web Wiz' at the address and website below:-
'**
'**  Web Wiz Ltd, Unit 18, The Glenmore Centre, Fancy Road, Poole, Dorset, BH12 4FB, England
'**  https://www.webwiz.net
'**
'**  Removal or modification of this copyright notice will violate the license contract.
'**
'****************************************************************************************



'Dimension global variables
Dim rsDateTimeFormat		'Holds the date a time data
Dim saryDateTimeData		'Holds the info for converting the date and time
Dim intLoopCounter		'loop counter


'******************************************
'***   	Initialise  array		***
'******************************************

'The date and time formatting data is feed into an application array as this data won't change 
'between users and pages so cuts done on un-necessary calls to the database
	
'Initialise  the array from the application veriable
If isArray(Application(strAppPrefix & "saryAppDateTimeFormatData")) AND blnUseApplicationVariables Then
	
	saryDateTimeData = Application(strAppPrefix & "saryAppDateTimeFormatData")


'Else the application level array holding the date and time data is not created so create it
Else
	'Craete a recordset to get the date and time format data
	Set rsDateTimeFormat = Server.CreateObject("ADODB.Recordset")
	
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "DateTimeFormat.ID, " & strDbTable & "DateTimeFormat.Date_format, " & strDbTable & "DateTimeFormat.Year_format, " & strDbTable & "DateTimeFormat.Seporator, " & strDbTable & "DateTimeFormat.Month1, " & strDbTable & "DateTimeFormat.Month2, " & strDbTable & "DateTimeFormat.Month3, " & strDbTable & "DateTimeFormat.Month4, " & strDbTable & "DateTimeFormat.Month5, " & strDbTable & "DateTimeFormat.Month6, " & strDbTable & "DateTimeFormat.Month7, " & strDbTable & "DateTimeFormat.Month8, " & strDbTable & "DateTimeFormat.Month9, " & strDbTable & "DateTimeFormat.Month10, " & strDbTable & "DateTimeFormat.Month11, " & strDbTable & "DateTimeFormat.Month12, " & strDbTable & "DateTimeFormat.Time_format, " & strDbTable & "DateTimeFormat.am, " & strDbTable & "DateTimeFormat.pm, " & strDbTable & "DateTimeFormat.Time_offset, " & strDbTable & "DateTimeFormat.Time_offset_hours " & _
	"FROM " & strDbTable & "DateTimeFormat" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "DateTimeFormat.ID = 1;"

	'Query the database
	rsDateTimeFormat.Open strSQL, adoCon
	
	'If there are records returned then enter the data returned into an array
	If NOT rsDateTimeFormat.EOF Then
		'Place the date time data into an array
		saryDateTimeData = rsDateTimeFormat.GetRows()
	End If
	
	'Relese server objects
	rsDateTimeFormat.Close
	Set rsDateTimeFormat = Nothing
	
	'Update the application level variable holding the the time and date formatting (better performance)
	If blnUseApplicationVariables Then
		'Lock the application so that no other user can try and update the application level variable at the same time
		Application.Lock
		
		'Update the application level variable
		Application(strAppPrefix & "saryAppDateTimeFormatData") = saryDateTimeData
		
		'Unlock the application
		Application.UnLock
	End If
End If





'******************************************
'***          Date Format   	      *****
'******************************************

'Function to format date
Private Function DateFormat(ByVal dtmDate)

	Dim strNewDate		'Holds the new date format
	Dim intDay		'Holds the integer number for the day
	Dim intMonth		'Holds a integer number from 1 to 12 for the month
	Dim strMonth		'Holds the month in it's final format which may be a number or a string so it is set to a sring value
	Dim intYear		'Holds the year
	Dim dtmNow		'Holds the present date
	Dim dtmTempDate		'Temprary storage area for date
	Dim intHour		'Holds the integer number for the hours
	Dim intMinute		'Holds a integer number for the mintes
	Dim intTempMinutes
	Dim intTempHours
	
	
	'If the array is empty set the date as server default
	If isNull(saryDateTimeData) Then
		
		'Set the date as orginal
		DateFormat = dtmDate
		
	'If there is a data in the array then format the date
	Else
		
		'Get the date now from the server
		dtmNow = Now()
		
		'Place the global forum time off-set onto the recorded database time
		If saryDateTimeData(19,0) = "+" Then
			dtmTempDate = DateAdd("h", + saryDateTimeData(20,0), dtmDate)
			dtmNow = DateAdd("h", + saryDateTimeData(20,0), dtmNow)
		ElseIf saryDateTimeData(19,0) = "-" Then
			dtmTempDate = DateAdd("h", -  saryDateTimeData(20,0), dtmDate)
			dtmNow = DateAdd("h", - saryDateTimeData(20,0), dtmNow)
		End If
		
		'Place the users time off set onto the recorded database time
		If strTimeOffSet = "+" Then
			dtmTempDate = DateAdd("h", + intTimeOffSet, dtmTempDate)
			dtmNow = DateAdd("h", + intTimeOffSet, dtmNow)
		ElseIf strTimeOffSet = "-" Then
			dtmTempDate = DateAdd("h", - intTimeOffSet, dtmTempDate)
			dtmNow = DateAdd("h", - intTimeOffSet, dtmNow)
		End If
		
		
		'Seprate the date into differnet strings
		intDay = CInt(Day(dtmTempDate))
		intMonth = CInt(Month(dtmTempDate))
		intYear = CInt(Year(dtmTempDate))
		intHour = CInt(Hour(dtmTempDate))
		intMinute = CInt(Minute(dtmTempDate))
		
		
		'If not a search engine format date in more human readable format
		If DateDiff("h", dtmTempDate, dtmNow) < 24 AND NOT strOSType = "Search Robot" Then
			
			intTempHours = 0 
			intTempMinutes = DateDiff("n", dtmTempDate, dtmNow)
			
			'If less than 1 minute
			If DateDiff("s", dtmTempDate, dtmNow) < 60 Then
				
				'If 1 minute then use singler for wording
				If DateDiff("s", dtmTempDate, dtmNow) < 1 Then
					DateFormat =  "1 " & strTxtSecond
				'Else use minutes for wording
				Else
					DateFormat = DateDiff("s", dtmTempDate, dtmNow) & " " & strTxtSecounds
				End If

			'If less than 60 minutes
			ElseIf intTempMinutes < 60 Then
				
				'If 1 minute then use singler for wording
				If intTempMinutes = 1 Then
					DateFormat =  "1 " & strTxtMinute
				'Else use minutes for wording
				Else
					DateFormat = intTempMinutes & " " & strTxtMinutes
				End If
				
			'If less than 24 hours
			Else 
				'Loop through calculating the number of hours and removing them from the minutes
				Do while NOT intTempMinutes <= 60
					
					'Add 1 to the hour
					intTempHours = intTempHours + 1 
					
					'Remove 60 minutes from toltal minutes, this will remove whole hors from the overal minutes
					intTempMinutes = intTempMinutes - 60
				Loop
				
				'Return number of hours and minutes
				If intTempHours = 1 Then
					DateFormat = intTempHours  & " " & strTxtHour & " " & intTempMinutes & " " & strTxtMinutes
				Else
					DateFormat = intTempHours  & " " & strTxtHours & " " & intTempMinutes & " " & strTxtMinutes
				End If
	
			End If
			
			
			'If todays date shown in bold 
			If blnBoldToday Then 
				DateFormat = "<strong>" & DateFormat & "</strong> " & strTxtAgo
			Else
				DateFormat = DateFormat & " " & strTxtAgo
			End If
		
			
	
		'Else if the date was yesterday return yeserday as the date
		ElseIf intDay = (CInt(Day(dtmNow))-1) AND intMonth = CInt(Month(dtmNow)) AND intYear = CInt(Year(dtmNow)) AND NOT strOSType = "Search Robot" Then
				
			DateFormat = strTxtYesterday

		
		'Else if the date is 30 december 1899 then date is unknown
		ElseIf intDay = 30 AND intMonth = 12 AND intYear = 1899 Then
			
			DateFormat = strTxtUnknown
		
		'Else format the date	
		Else
		
		
			'Place 0 infront of days under 10
			If intDay < 10 then intDay = "0" & intDay
		
			'If the year is two digits then sorten the year string
			If saryDateTimeData(2,0) = "short" Then intYear = Right(intYear, 2)
			
			'Format the month
			strMonth = saryDateTimeData((intMonth + 3),0)
			
			'If the user has entered their own date format get that
			If NOT strDateFormat = "" Then saryDateTimeData(1,0) = strDateFormat

			'Format the date
			Select Case saryDateTimeData(1,0)
				
				'Format dd/mm/yy
				Case "dd/mm/yy"
					DateFormat = intDay & saryDateTimeData(3,0) & strMonth & saryDateTimeData(3,0) & intYear
				
				'Format mm/dd/yy
				Case "mm/dd/yy"
					DateFormat = strMonth & saryDateTimeData(3,0) & intDay & saryDateTimeData(3,0) & intYear	
			
				'Format yy/dd/mm
				Case "yy/dd/mm"
					DateFormat = intYear & saryDateTimeData(3,0) & intDay & saryDateTimeData(3,0) & strMonth
				
				'Format yy/mm/dd
				Case "yy/mm/dd"
					DateFormat = intYear & saryDateTimeData(3,0) & strMonth & saryDateTimeData(3,0) & intDay		
			End Select
		
		End If	
	
	End If
	
End Function






'******************************************
'***          Time Format   	      *****
'******************************************

'Function to format time
Function TimeFormat(ByVal dtmTime)

	Dim strNewDate		'Holds the new date format
	Dim intHour		'Holds the integer number for the hours
	Dim intMinute		'Holds a integer number for the mintes
	Dim strPeriod		'Holds wether it is am or pm
	Dim dtmTempTime		'Temporary storage area for the time

	
	'If the array is empty then return tyhe original time
	If isNull(saryDateTimeData) Then
		
		'Set the date as server default
		TimeFormat = dtmTime
		
	'If there is a data in the array then format the date
	Else
	
		'Place the global forum time off-set onto the recorded database time
		If saryDateTimeData(19,0) = "+" Then
			dtmTempTime = DateAdd("h", + saryDateTimeData(20,0), dtmTime)
		ElseIf saryDateTimeData(19,0) = "-" Then
			dtmTempTime = DateAdd("h", -  saryDateTimeData(20,0), dtmTime)
		End If
		
		'Place the users time off-set onto the recorded database time
		If strTimeOffSet = "+" Then
			dtmTempTime = DateAdd("h", + intTimeOffSet, dtmTempTime)
		ElseIf strTimeOffSet = "-" Then
			dtmTempTime = DateAdd("h", - intTimeOffSet, dtmTempTime)
		End If
		
		'Seprate the time into differnet strings
		intHour = CInt(Hour(dtmTempTime))
		intMinute = CInt(Minute(dtmTempTime))
		
		'Place 0 infront of minutes under 10
		If intMinute < 10 then intMinute = "0" & intMinute
	
		'If the time is 12 hours then change the time to 12 hour clock
		If CInt(saryDateTimeData(16,0)) = 12 Then
			
			'Set the time period
			If intHour >= 12 then 
				strPeriod = saryDateTimeData(18,0)
			Else 
				strPeriod = saryDateTimeData(17,0)
			End If
			
			
			'Change the hour to 12 hour clock time
			Select Case intHour
				Case 00
					intHour = 12
				Case 01
					intHour = 1
				Case 02
					intHour = 2
				Case 03
					intHour = 3
				Case 04
					intHour = 4
				Case 05
					intHour = 5					
				Case 06
					intHour = 6					
				Case 07
					intHour = 7					
				Case 08
					intHour = 8					
				Case 09
					intHour = 9
				Case 13
					intHour = 1
				Case 14
					intHour = 2					
				Case 15
					intHour = 3					
				Case 16
					intHour = 4					
				Case 17
					intHour = 5					
				Case 18
					intHour = 6					
				Case 19
					intHour = 7					
				Case 20
					intHour = 8					
				Case 21
					intHour = 9					
				Case 22
					intHour = 10					
				Case 23
					intHour = 11	
						
			End Select
		
		'ElseIf it is 24 hour clock place another 0 infront of anything below 10 hours
		ElseIf intHour < 10 Then 
			intHour = "0" & intHour
		End If
		
		'Return the Formated time
		TimeFormat = intHour & ":" & intMinute & strPeriod
	
	End If		
End Function




'**************************************************************************
'***   Date Format for without 'Today' and no date alteration option  *****
'**************************************************************************

'Function to format date that doesn't use 'Today' or 'Yesterday' in dates
Private Function stdDateFormat(ByVal dtmDate, ByVal blnDateOffSet)

	Dim strNewDate		'Holds the new date format
	Dim intDay		'Holds the integer number for the day
	Dim intMonth		'Holds a integer number from 1 to 12 for the month
	Dim strMonth		'Holds the month in it's final format which may be a number or a string so it is set to a sring value
	Dim intYear		'Holds the year
	Dim dtmTempDate		'Temporary storage area for date
	
	
	'If the array is empty set the date as server default
	If isNull(saryDateTimeData) Then
		
		'Set the date as orginal
		stdDateFormat = dtmDate
		
	'If there is a data in the array then format the date
	Else
		
		'If date time off set is included the calaculate new date
		If blnDateOffSet Then
			'Place the global forum time off-set onto the recorded database time
			If saryDateTimeData(19,0) = "+" Then
				dtmTempDate = DateAdd("h", + saryDateTimeData(20,0), dtmDate)
			ElseIf saryDateTimeData(19,0) = "-" Then
				dtmTempDate = DateAdd("h", -  saryDateTimeData(20,0), dtmDate)
			End If
			
			'Place the users time off set onto the recorded database time
			If strTimeOffSet = "+" Then
				dtmTempDate = DateAdd("h", + intTimeOffSet, dtmTempDate)
			ElseIf strTimeOffSet = "-" Then
				dtmTempDate = DateAdd("h", - intTimeOffSet, dtmTempDate)
			End If
		
		'Else just process the date 'as is'
		Else
			dtmTempDate = dtmDate
		End If
		
		
		'Seprate the date into differnet strings
		intDay = CInt(Day(dtmTempDate))
		intMonth = CInt(Month(dtmTempDate))
		intYear = CInt(Year(dtmTempDate))
		
		'Place 0 infront of days under 10
		If intDay < 10 then intDay = "0" & intDay
	
		'If the year is two digits then sorten the year string
		If saryDateTimeData(2,0) = "short" Then intYear = Right(intYear, 2)
		
		'Format the month
		strMonth = saryDateTimeData((intMonth + 3),0)
		
		'If the user has entered their own date format get that
		If NOT strDateFormat = "" Then saryDateTimeData(1,0) = strDateFormat

		'Format the date
		Select Case saryDateTimeData(1,0)
			
			'Format dd/mm/yy
			Case "dd/mm/yy"
				stdDateFormat = intDay & saryDateTimeData(3,0) & strMonth & saryDateTimeData(3,0) & intYear
			
			'Format mm/dd/yy
			Case "mm/dd/yy"
				stdDateFormat = strMonth & saryDateTimeData(3,0) & intDay & saryDateTimeData(3,0) & intYear	
		
			'Format yy/dd/mm
			Case "yy/dd/mm"
				stdDateFormat = intYear & saryDateTimeData(3,0) & intDay & saryDateTimeData(3,0) & strMonth
			
			'Format yy/mm/dd
			Case "yy/mm/dd"
				stdDateFormat = intYear & saryDateTimeData(3,0) & strMonth & saryDateTimeData(3,0) & intDay		
		End Select
	End If
End Function




'******************************************
'***          Date/Time Number 	      *****
'******************************************

'Function to format time
Function DateTimeNum(ByVal strElement)

	Dim strDateElement

	'Get the date/time element required
	Select Case strElement
		Case "Year"
			strDateElement = CInt(Year(Now()))
		Case "Month"
			strDateElement = CInt(Month(Now()))
		Case "Day"
			strDateElement = CInt(Day(Now()))
		Case "Hour"
			strDateElement = CInt(Hour(Now()))
		Case "Minute"
			strDateElement = CInt(Minute(Now()))
		Case "Second"
			strDateElement = CInt(Second(Now()))
	End Select
	
	'If below 10 then place a 0 in front of te returned string
	If strDateElement < 10 then strDateElement = "0" & strDateElement
	
	'Return function
	DateTimeNum = strDateElement
End Function









'******************************************
'***         RSS Date Format   	      *****
'******************************************

'Function to format date for RSS feeds
Private Function RssDateFormat(ByVal dtmDate, ByVal strTimeZone)

	Dim strNewDate		'Holds the new date format
	Dim intDay		'Holds the integer number for the day
	Dim intWeekDay		'Holds the weekday in interget format
	Dim strWeekDay		'Holds the day in string format
	Dim intMonth		'Holds a integer number from 1 to 12 for the month
	Dim strMonth		'Holds the month in it's final format which may be a number or a string so it is set to a sring value
	Dim intYear		'Holds the year
	Dim dtmNow		'Holds the present date
	Dim dtmTempDate		'Temprary storage area for date
	Dim intHour		'Holds the integer number for the hours
	Dim intMinute		'Holds a integer number for the mintes
	Dim intSeconds		'Holds the secounds
	
	
	'If the array is empty set the date as server default
	If isNull(saryDateTimeData) Then
		
		'Set the date as orginal
		RssDateFormat = dtmDate
		
	'If there is a data in the array then format the date
	Else
		
		'Get the date now from the server
		dtmNow = Now()
		
		'Place the global forum time off-set onto the recorded database time
		If saryDateTimeData(19,0) = "+" Then
			dtmTempDate = DateAdd("h", + saryDateTimeData(20,0), dtmDate)
			dtmNow = DateAdd("h", + saryDateTimeData(20,0), dtmNow)
		ElseIf saryDateTimeData(19,0) = "-" Then
			dtmTempDate = DateAdd("h", -  saryDateTimeData(20,0), dtmDate)
			dtmNow = DateAdd("h", - saryDateTimeData(20,0), dtmNow)
		End If
		
		'Place the users time off set onto the recorded database time
		If strTimeOffSet = "+" Then
			dtmTempDate = DateAdd("h", + intTimeOffSet, dtmTempDate)
			dtmNow = DateAdd("h", + intTimeOffSet, dtmNow)
		ElseIf strTimeOffSet = "-" Then
			dtmTempDate = DateAdd("h", - intTimeOffSet, dtmTempDate)
			dtmNow = DateAdd("h", - intTimeOffSet, dtmNow)
		End If
		
		
		'Seprate the date into differnet strings
		intDay = CInt(Day(dtmTempDate))
		intWeekDay = CInt(WeekDay(dtmTempDate))
		intMonth = CInt(Month(dtmTempDate))
		intYear = CInt(Year(dtmTempDate))
		intHour = CInt(Hour(dtmTempDate))
		intMinute = CInt(Minute(dtmTempDate))
		intSeconds = CInt(Second(dtmTempDate))
		
		
		'Place 0 infront of days under 10
		If intDay < 10 then intDay = "0" & intDay
			
		'Place 0 infront of hours under 10
		If intHour < 10 then intHour = "0" & intHour
				
		'Place 0 infront of minutes under 10
		If intMinute < 10 then intMinute = "0" & intMinute
			
		'Place 0 infront of hours under 10
		If intSeconds < 10 then intSeconds = "0" & intSeconds
		
		'Format the month
		Select Case intMonth
			Case 1
				strMonth = "Jan"
			Case 2
				strMonth = "Feb"
			Case 3
				strMonth = "Mar"
			Case 4
				strMonth = "Apr"
			Case 5
				strMonth = "May"
			Case 6
				strMonth = "Jun"
			Case 7
				strMonth = "Jul"
			Case 8
				strMonth = "Aug"
			Case 9
				strMonth = "Sep"
			Case 10
				strMonth = "Oct"
			Case 11
				strMonth = "Nov"
			Case 12
				strMonth = "Dec"
		End Select
		
		
		'Format the day
		Select Case intWeekDay
			Case 1
				strWeekDay = "Sun"
			Case 2
				strWeekDay = "Mon"
			Case 3
				strWeekDay = "Tue"
			Case 4
				strWeekDay = "Wed"
			Case 5
				strWeekDay = "Thu"
			Case 6
				strWeekDay = "Fri"
			Case 7
				strWeekDay = "Sat"
		End Select	
		

		'Format the date
		RssDateFormat = strWeekDay & ", " & intDay & " " & strMonth & " " & intYear & " " & intHour & ":" & intMinute & ":" & intSeconds & " " & strTimeZone
	End If
	
End Function








'******************************************
'***  	  Now() date with off-set     *****
'******************************************

'Calculate the now() date according to any server time off-set
Private Function getNowDate()
	
	Dim dtmNow
	
	'Get the date now from the server
	dtmNow = Now()
		
	'Place the global forum time off-set onto the recorded database time
	If saryDateTimeData(19,0) = "+" Then
		dtmNow = DateAdd("h", + saryDateTimeData(20,0), dtmNow)
	ElseIf saryDateTimeData(19,0) = "-" Then
		dtmNow = DateAdd("h", - saryDateTimeData(20,0), dtmNow)
	End If
	
	'Place the users time off set onto the recorded database time
	If strTimeOffSet = "+" Then
		dtmNow = DateAdd("h", + intTimeOffSet, dtmNow)
	ElseIf strTimeOffSet = "-" Then
		dtmNow = DateAdd("h", - intTimeOffSet, dtmNow)
	End If
	
	'Return date
	getNowDate = dtmNow
End Function


%>