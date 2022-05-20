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





'******************************************
'***  	   No. of days in month	      *****
'******************************************

'Calculate the number of days in a month
Private Function getMonthDayNo(ByVal intMonth, ByVal intYear)

	'The date formats below are in internation ISO date format yyyy-mm-dd for server compatibility
	
	'If the month has 31 days then return 31
	If IsDate(intYear & "-" & intMonth & "-" & 31) Then 
		getMonthDayNo = 31
		
	'If the month has 30 days then return 30
	ElseIf IsDate(intYear & "-" & intMonth & "-" & 30) Then 
		getMonthDayNo = 30
	
	'If the month has 29 days then return 29 (Leap Year)
	ElseIf IsDate(intYear & "-" & intMonth & "-" & 29) Then 
		getMonthDayNo = 29
	
	'If the month has 28 days then return 28 (February (non leap year))
	ElseIf IsDate(intYear & "-" & intMonth & "-" & 28) Then 
		getMonthDayNo = 28
	End If
End Function





'******************************************
'***  	 Convert int. to month name   *****
'******************************************

'Get the month in name format as set in language file
Private Function getMonthName(ByVal intMonth)
	
	Select Case intMonth
		Case 1
			getMonthName = strTxtJanuary
		Case 2
			getMonthName =  strTxtFebruary
		Case 3
			getMonthName =  strTxtMarch
		Case 4
			getMonthName =  strTxtApril
		Case 5
			getMonthName =  strTxtMay
		Case 6
			getMonthName =  strTxtJune
		Case 7
			getMonthName =  strTxtJuly
		Case 8
			getMonthName =  strTxtAugust
		Case 9
			getMonthName =  strTxtSeptember
		Case 10
			getMonthName =  strTxtOctober
		Case 11
			getMonthName =  strTxtNovember
		Case 12
			getMonthName =  strTxtDecember	
	End Select
End Function






'******************************************
'***  	   Create Small Calendar      *****
'******************************************

'Function to create small calendar for months
Private Sub displayMonth(ByVal intMonth, ByVal intYear, ByVal intViewType)

	Dim intDay
	Dim intFistDayOfMonth
	Dim intMaxNoMonthDays
	Dim strMonth
	Dim intWeekLoop
	Dim intDayLoopCounter
	Dim intNoOfEvents
	Dim lngTopicID
	Dim strSeoSubject

	'Get the first day of the month (use internation ISO date fomat (yyyy-mm-dd) for server compatibility)
	intFistDayOfMonth = WeekDay(intYear & "-" & intMonth & "-01")
	
	'Get the number of days in the month
	intMaxNoMonthDays = getMonthDayNo(intMonth, intYear)
	
	'Get the month in name format
	strMonth = getMonthName(intMonth)

	'Create table for small month calendar
	Response.Write(vbCrLf & "<table cellspacing=""1"" cellpadding=""3"" class=""tableBorder"" style=""width:98%;"" align=""center"">" & _
	vbCrLf & " <tr class=""tableLedger"">" & _
	vbCrLf & "  <td width=""100%"" colspan=""8"" align=""left""><a href=""calendar.asp?M=" & intMonth & "&Y=" & intYear  & strQsSID2 &  """ title=""" & strTxtViewMonthInDetail & """>" & strMonth & " " & intYear & "</a></td>" & _
	vbCrLf & " </tr>" & _
	vbCrLf & " <tr align=""center"" class=""calLedger"">" & _
	vbCrLf & "  <td width=""12.5%"">&nbsp;</td>" & _
	vbCrLf & "  <td width=""12.5%"">" & strTxtSu & "</td>" & _
	vbCrLf & "  <td width=""12.5%"">" & strTxtMo & "</td>" & _
	vbCrLf & "  <td width=""12.5%"">" & strTxtTu & "</td>" & _
	vbCrLf & "  <td width=""12.5%"">" & strTxtWe & "</td>" & _
	vbCrLf & "  <td width=""12.5%"">" & strTxtTh & "</td>" & _
	vbCrLf & "  <td width=""12.5%"">" & strTxtFr & "</td>" & _
	vbCrLf & "  <td width=""12.5%"">" & strTxtSa & "</td>" & _
	vbCrLf & " </tr>")
	
	'Loop through to display the weeks of the month (we use 6 as this is the most required to cover an entire month)
	For intWeekLoop = 1 TO 6
	   	
	   	'Display ledger info
	   	Response.Write(vbCrLf & " <tr align=""center"">" & _
	   	vbCrLf & "  <td class=""calLedger""><a href=""calendar_week.asp?M=" & intMonth & "&Y=" & intYear & "&W=" & intWeekLoop  & strQsSID2 &  """ title=""" & strTxtViewWeekInDetail & """>&gt;</a></a></td>")
	   	
	   	'Loop through the 7 days of the week
	   	For intDayLoopCounter = 1 TO 7 
	   	
	   		'Increment the day by 1
	   		If intDay > 0 Then intDay = intDay + 1
	   			
	   		'See if this is the first day of the month
			If intFistDayOfMonth = intDayLoopCounter AND intDay = 0 Then intDay = 1
	
			'Write the table cell
			Response.Write(vbCrLf & "  <td")
			
			'Calculate the class for the table cell
			If intDay => 1 AND intDay <= intMaxNoMonthDays Then
				
				'If today place a red border around the day
				If intMonth = Month(dtmNow) AND intDay = Day(dtmNow) AND intYear = Year(dtmNow) Then 
					Response.Write(" class=""calTodayCell"">")
				Else
					Response.Write(" class=""calDateCell"">")
				End If
			
			'Else the day is not a date in this month
			Else
				Response.Write(" class=""calEmptyDateCell"">")
			End If
			
			
			
			'If this is a day in the month display day number
			If intDay => 1 AND intDay <= intMaxNoMonthDays Then 
				
				'If there is an event on this date display a link to it
				If isArray(sarryEvents) Then
			
					'Initlise the loop array
					intDbArrayLoop = 0
					intNoOfEvents = 0
			
					'Loop through the events array
					Do While intDbArrayLoop <= Ubound(sarryEvents,2)
			
						'If there isn't an end date set, set the end date as the event start date to prevent errors
						If isDate(sarryEvents(3,intDbArrayLoop)) = False Then sarryEvents(3,intDbArrayLoop) = sarryEvents(2,intDbArrayLoop)
			
						'If an event is found for this date display it
						If CDate(intYear & "-" & intMonth & "-" & intDay) >= CDate(sarryEvents(2,intDbArrayLoop)) AND CDate(intYear & "-" & intMonth & "-" & intDay) <= CDate(sarryEvents(3,intDbArrayLoop)) Then
							
							'Read the event details
							strSubject = sarryEvents(1,intDbArrayLoop)
							strSeoSubject = sarryEvents(1,intDbArrayLoop)
							lngTopicID = sarryEvents(0,intDbArrayLoop)
							
							'Trim the subject
							strSubject = TrimString(strSubject, 25)
					
							'Clean up input to prevent XXS hack
							strSubject = formatInput(strSubject)
							
							'Increment the number of events
							intNoOfEvents = intNoOfEvents + 1
						End If
			
						'Move to next array position
						intDbArrayLoop = intDbArrayLoop + 1
					Loop
				End If
				
				
				'Write the HTML for the date
				'If 1 event use the topic as the title and link to event
				If intNoOfEvents = 1 Then
					Response.Write("<a href=""forum_posts.asp?TID=" & lngTopicID  & strQsSID2 & SeoUrlTitle(strSeoSubject, "&title=") &  """ title=""" & strSubject & """>" & intDay & "</a>")
				'If more than 1 event link to week calendar view to see the events
				ElseIf intNoOfEvents > 1 Then
					Response.Write("<a href=""calendar_week.asp?M=" & intMonth & "&Y=" & intYear & "&W=" & intWeekLoop  & strQsSID2 &  """ title=""" & intNoOfEvents & " " & strTxtEvents & """>" & intDay & "</a>")
				'Else just show the date
				Else
					Response.Write(intDay)
				End If
			End If
			
			Response.Write("</td>")
		Next
	   
		Response.Write(vbCrLf & " </tr>")
	 
		'If we have run out of weeks in this month exit loop
		If intViewType = 1 AND intMaxNoMonthDays =< intDay Then Exit For
	Next
	
	Response.Write(vbCrLf & "</table>")
End Sub







'******************************************
'***  	Get events from database      *****
'******************************************

'Sub procedure for running the SQL to get the events from the database
Private Sub getEvents()

	'Initalise SQL query
	strSQL = "" & _
	"SELECT " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Event_date, " & strDbTable & "Topic.Event_date_end, " & strDbTable & "Thread.Message " & _
	"FROM " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Topic.Start_Thread_ID = " & strDbTable & "Thread.Thread_ID " & _
		"AND       ((" & strDbTable & "Topic.Event_date BETWEEN " & dtmDbStartDate & " AND " & dtmDbEndDate & ") " & _
			"OR (" & strDbTable & "Topic.Event_date_end BETWEEN " & dtmDbStartDate & " AND " & dtmDbEndDate & ") " & _
			"OR (" & strDbTable & "Topic.Event_date < " & dtmDbStartDate & " AND " & strDbTable & "Topic.Event_date_end > " & dtmDbEndDate & ")) " & _
		"AND (" & strDbTable & "Topic.Forum_ID " & _
			"IN (" & _
				"SELECT " & strDbTable & "Permissions.Forum_ID " & _
				"FROM " & strDbTable & "Permissions " & strDBNoLock & " " & _
				"WHERE (" & strDbTable & "Permissions.Group_ID = " & intGroupID & " OR " & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & ") " & _
					"AND " & strDbTable & "Permissions.View_Forum=" & strDBTrue & _
			")" & _
		")"
	'If this isn't a moderator only display hidden events if the user posted them
	If blnModerator = false AND blnAdmin = false Then
		strSQL = strSQL & "AND (" & strDbTable & "Topic.Hide=" & strDBFalse & " "
		'Don't display hidden posts if guest
		If intGroupID <> 2 Then strSQL = strSQL & "OR " & strDbTable & "Topic.Start_Thread_ID = " & lngLoggedInUserID
		strSQL = strSQL & ") "
	End If
	strSQL = strSQL & "ORDER BY " & strDbTable & "Topic.Event_date ASC;"
	
	
	
	'Set error trapping
	On Error Resume Next
		
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'If an error has occurred write an error to the page
	If Err.Number <> 0 AND  strDatabaseType = "mySQL" Then	
		Call errorMsg("An error has occurred while executing SQL query on database.<br />Please check that the MySQL Server version is 4.1 or above.", "getEvents()_get_events_data", "functions_calendar.asp")
	ElseIf Err.Number <> 0 Then	
		Call errorMsg("An error has occurred while executing SQL query on database.", "getEvents()_get_events_data", "functions_calendar.asp")
	End If
				
	'Disable error trapping
	On Error goto 0
	
	
	'Read in the event recordset into an array
	If NOT rsCommon.EOF Then sarryEvents = rsCommon.GetRows()
		
	'Close the recordset
	rsCommon.Close
		
	'SQL Query Array Look Up table
	'0 = Topic_ID
	'1 = Subject
	'2 = Event_date
	'3 = Event_date_end
	'4 = Message
End Sub







'******************************************
'***  	Get birthdays from database   *****
'******************************************

'Sub procedure for running the SQL to get the birthdays from the database
Private Sub getBirthdays(intMonth, intWeek)

	Dim intAltMonth
	
	
	'Calulate if we need to get birthdays from a previous or next month (for week view)
	If intWeek = 1 AND intMonth = 1 Then 
		intAltMonth = 12
	ElseIf intWeek = 1 Then
		intAltMonth = intMonth - 1
	ElseIf intWeek > 3 AND intMonth = 12 Then 
		intAltMonth = 1
	ElseIf intWeek > 3 Then
		intAltMonth = intMonth + 1
	End If
	

	'Initalise SQL query
	strSQL = "" & _
	"SELECT " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.DOB " & _
	"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
	"WHERE MONTH(" & strDbTable & "Author.DOB) = " & intMonth & " "
	
	'Get birthdays from an adjacent month
	If intAltMonth <> "" Then strSQL = strSQL & "OR MONTH(" & strDbTable & "Author.DOB) = " & intAltMonth & " "
	
	strSQL = strSQL & _
	"ORDER BY " & strDbTable & "Author.Username ASC;"
	
	'Set error trapping
	On Error Resume Next
		
	'Query the database
	rsCommon.Open strSQL, adoCon

	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then Call errorMsg("An error has occurred while executing SQL query on database.", "getBirthdays()_get_birthdays_data", "functions_calendar.asp")
	
				
	'Disable error trapping
	On Error goto 0
	
	
	'Read in the event recordset into an array
	If NOT rsCommon.EOF Then saryBirthdays = rsCommon.GetRows()
		
	'Close the recordset
	rsCommon.Close
	
	'SQL Query Array Look Up table
	'0 = Author_ID
	'1 = Username
	'2 = DOB
End Sub






'******************************************
'***  	Hide/Show Birthday Function    *****
'******************************************

'Sub procedure for running the SQL to get the birthdays from the database
Private Function showBirthdays()

	'Get what date to show topics till from querystring
	If isNumeric(Request.QueryString("DB")) AND Request.QueryString("DB") <> "" Then
	
		Call saveSessionItem("DB", Request.QueryString("DB"))
		showBirthdays = IntC(Request.QueryString("DB"))
	
	'Get what date to show topics
	ElseIf getSessionItem("DB") <> "" Then
	
		showBirthdays = IntC(getSessionItem("DB"))
	
	'Else if there is no cookie use the default set by the forum
	Else
		showBirthdays = 0
	End If
End Function
%>