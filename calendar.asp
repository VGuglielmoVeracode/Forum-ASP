<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="language_files/calendar_language_file_inc.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
<!--#include file="functions/functions_calendar.asp" -->
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




'Set the response buffer to true as we maybe redirecting
Response.Buffer = True



'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"


'Dimension variables
Dim sarryEvents		'Holds the event recordset in an array
Dim saryBirthdays	'Holds the birthdays recordset in an array
Dim strSubject		'Holds the event subject
Dim strMessage		'Holds the event message
Dim strMonth		'Holds the month name
Dim intMonth		'Holds the integer value for the month
Dim intYear		'Holds the year
Dim intFistDayOfMonth	'Holds the first day of the month as an interger
Dim intWeekLoop		'Week Loop counter
Dim intDayLoopCounter	'Day loop counter
Dim intDay		'Holds the day
Dim intMaxNoMonthDays	'Holds total number of days for the month
Dim dtmNow		'Holds the present date according to any server off-set
Dim intView		'Holds the type of view (1=monh, 2=week, 3=year)
Dim dtmDbStartDate	'Holds the database search start date
Dim dtmDbEndDate	'Holds the database search end date
Dim intDbArrayLoop	'Database array loop counter
Dim intNoOfBirthdays	'Holds the number of birthdays
Dim intAge		'Holds the age for birthdays
Dim strMemBirthdays	'Holds the members birthdays
Dim strUserName		'Holds the username of the member with a birthday
Dim lngUserProfile	'Holds the user profile number
Dim blnShowBirthdays	'Holds if birthdays are hidden or not
Dim strPageRedirect	'Holds the page redirect details
Dim intWeek





'Test querystrings for any SQL Injection keywords
Call SqlInjectionTest(Request.QueryString())



'If the calendar system is disabled send the user away
If blnCalendar = False Then 
	
	'Clean up
	Call closeDatabase()
	
	'Send to default page
	Response.Redirect("default.asp" & strQsSID1)
End If




'If this is a year view redirect to the year calendar view page
If Request.QueryString("V") = "3" Then
	strPageRedirect = "calendar_year.asp?Y=" & Request.QueryString("Y")
	If Request.QueryString("DB") <> "" Then strPageRedirect = strPageRedirect & "&DB=" & Request.QueryString("DB")
	Response.Redirect(strPageRedirect & strQsSID3)
'Else if week view
ElseIf Request.QueryString("V") = "2" Then
	If isNumeric(Request.QueryString("W")) Then intWeek = IntC(Request.QueryString("W")) Else intWeek = 1
	strPageRedirect = "calendar_week.asp?M=" & Request.QueryString("M") & "&Y=" & Request.QueryString("Y") & "&W=" & intWeek
	If Request.QueryString("DB") <> "" Then strPageRedirect = strPageRedirect & "&DB=" & Request.QueryString("DB")
	Response.Redirect(strPageRedirect & strQsSID3)
End If


'Intilise variables
dtmNow = getNowDate()
blnShowBirthdays = CBool(showBirthdays())
intDay = 0


'Read in the year to view
If IntC(Request.QueryString("Y")) => 2001 AND  IntC(Request.QueryString("Y")) =< Cint(Year(dtmNow)+5) Then
	intYear = IntC(Request.QueryString("Y"))
'Else out of date range so relaod the page with this year
Else
	Response.Redirect("calendar.asp?M=" & IntC(Request.QueryString("M")) & "&Y=" & Year(dtmNow) & strQsSID3)
End If


'Read in the month to view
If IntC(Request.QueryString("M")) => 1 AND  IntC(Request.QueryString("M")) =< 12 Then
	intMonth = IntC(Request.QueryString("M"))
'Else use this month as the month to view
Else
	Response.Redirect("calendar.asp?M=" & Month(dtmNow) & "&Y=" & IntC(Request.QueryString("Y")) & strQsSID3)
End If


'Get the first day of the month (use internation ISO date fomat (yyyy-mm-dd) for server compatibility)
intFistDayOfMonth = WeekDay(intYear & "-" & intMonth & "-01")

'Get the number of days in the month
intMaxNoMonthDays = getMonthDayNo(intMonth, intYear)

'Get the month in name format
strMonth = getMonthName(intMonth)








'Calculate the db serach start date
If intMonth = 1 Then
	dtmDbStartDate = internationalDateTime(intYear-1 & "-12-1")
Else
	dtmDbStartDate = internationalDateTime(intYear & "-" & intMonth-1 & "-1")
End If

'Calculate the db serach end date
If intMonth = 12 Then
	dtmDbEndDate = internationalDateTime(intYear+1 & "-01-" & getMonthDayNo(1, intYear+1))
Else
	dtmDbEndDate = internationalDateTime(intYear & "-" & intMonth+1 & "-" & getMonthDayNo(intMonth+1, intYear))
End If

'SQL Server doesn't like ISO dates with '-' in them, so remove the '-' part
If strDatabaseType = "SQLServer" Then
	dtmDbStartDate = Replace(dtmDbStartDate, "-", "", 1, -1, 1)
	dtmDbEndDate = Replace(dtmDbEndDate, "-", "", 1, -1, 1)
End If

'Place the date in SQL safe # or '
If strDatabaseType = "Access" Then
	dtmDbStartDate = "#" & dtmDbStartDate & "#"
	dtmDbEndDate = "#" & dtmDbEndDate & "#"
Else
	dtmDbStartDate = "'" & dtmDbStartDate & "'"
	dtmDbEndDate = "'" & dtmDbEndDate & "'"
End If





'Call the sub procedure to get the events from the database
Call getEvents()
		
'SQL Query Array Look Up table
'0 = Topic_ID
'1 = Subject
'2 = Hide
'3 = Event_date
'4 = Message



'Call sub to get birthdays from the database
If blnDisplayBirthdays AND blnShowBirthdays Then Call getBirthdays(intMonth, 0)

'SQL Query Array Look Up table
'0 = Author_ID
'1 = Username
'2 = DOB


Dim lngTotalRecordsPages
lngTotalRecordsPages = 12
Dim intRecordPositionPageNum
Dim intPageLinkLoopCounter




'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers(strTxtViewing & " " & strTxtEvents, strMonth & " " & intYear, "calendar.asp?Y=" & intYear & "&M=" & intMonth, 0)
End If


'Clean up
Call closeDatabase()


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""calendar.asp" & strQsSID1 & """>" & strTxtEvents & "</a>" & strNavSpacer & strMonth & " " & intYear

'Status bar tools
strStatusBarTools = strStatusBarTools & "&nbsp;<a href=""RSS_calendar_feed.asp" & strQsSID1 & """ target=""_blank""><img src=""" & strImagePath & "rss." & strForumImageType & """ border=""0"" alt=""" & strTxtRSS & " - " & strTxtLatestEventFeed & """ title=""" & strTxtRSS & " - " & strTxtLatestEventFeed & """ /></a>"

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strMainForumName & " " & strTxtEvents & " - " & strMonth & " " & intYear %></title>
<meta name="generator" content="Web Wiz Forums" />
<meta name="description" content="<% = strBoardMetaDescription %>" />
<meta name="keywords" content="<% = strBoardMetaKeywords %>" />

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

'If RSS feed is enabled then have an alt link to it for browers that support RSS Feeds
If blnRSS Then Response.Write(vbCrLf & "<link rel=""alternate"" type=""application/rss+xml"" title=""RSS 2.0"" href=""RSS_calendar_feed.asp" & strQsSID1 & """ />")
%>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtEvents & " - " & strMonth & " " & intYear %></h1></td>
 </tr>
</table>
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td nowrap><!--#include file="includes/calendar_jump_inc.asp" --></td>
  <td align="right" nowrap><a href="calendar.asp?<% 
   
   	'Calculate if we need to chnage years
   	If intMonth = 1 Then 
		Response.Write("M=12&Y=" & intYear-1) 
	Else 
		Response.Write("M=" & intMonth-1 & "&Y=" & intYear) 
	End If
	
	%><% = strQsSID2 %>">&lt;&lt; <% = strTxtPrevious %></a>&nbsp;
   <select onchange="linkURL(this)" name="calSelect" id="calSelect"><%
    
    	Dim intJumpLoop
    	
    
    	'Setup list options
    	For intJumpLoop = 1 TO 12
    	
    		Response.Write(vbCrLf & "     <option value=""calendar.asp?M=" & intJumpLoop & "&Y=" & intYear & strQsSID2&"""")
    		If intJumpLoop = intMonth Then Response.Write(" selected")
    		Response.Write(">" & getMonthName(intJumpLoop) & "</option>")
    		
    	Next
    
     
%>
   </select>
   &nbsp;<a href="calendar.asp?<% 
   
   	'Calculate if we need to chnage years
   	If intMonth = 12 Then 
   		Response.Write("M=1&Y=" & intYear+1) 
   	Else 
   		Response.Write("M=" & intMonth+1 & "&Y=" & intYear) 
   	End If
   	
   	%><% = strQsSID2 %>"><% = strTxtNext %> &gt;&gt;</a></td>
 </tr>
</table>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td width="100%" colspan="8" align="left"><% = strMonth & " " & intYear %></td>
 </tr><%
 

'Loop through to display the weeks of the month (we use 6 as this is the most required to cover an entire month)
For intWeekLoop = 1 TO 6

%>
 <tr valign="top">
  <td align="center" height="80" width="2%" class="calLedger" valign="middle"><a href="calendar_week.asp?M=<% = intMonth %>&Y=<% = intYear %>&W=<% = intWeekLoop %><% = strQsSID2 %>" title="<% = strTxtViewWeekInDetail %>">&gt;<br />&gt;<br />&gt;<br />&gt;</a></td><%
   	
   	'Display ledger info
   	
   	'Loop through the 7 days of the week
   	For intDayLoopCounter = 1 TO 7 
   	
   		'Increment the day by 1
   		If intDay > 0 Then intDay = intDay + 1
   			
   		'See if this is the first day of the month
		If intFistDayOfMonth = intDayLoopCounter AND intDay = 0 Then intDay = 1

		'Write the table cell
		Response.Write(vbCrLf & "  <td width=""14%""")
		
		'Calculate the class for the table cell
		If intDay => 1 AND intDay <= intMaxNoMonthDays Then
			
			'If today place a red border around the day
			If intMonth = Month(dtmNow) AND intDay = Day(dtmNow) AND intYear = Year(dtmNow) Then 
				Response.Write(" class=""calTodayCell""")
			Else
				Response.Write(" class=""calDateCell""")
			End If
		
		'Else the day is not a date in this month
		Else
			Response.Write(" class=""calEmptyDateCell""")
		End If
		
		
		Response.Write(" style=""padding:0px"">")
		Response.Write("<div class=""calLedger"">")
		
			
		'If this is a day in the month display day number
		If intDay => 1 AND intDay <= intMaxNoMonthDays Then Response.Write("<span style=""float:right"">" & intDay & "</span>")
		
		'Display the day
		Select Case intDayLoopCounter
			Case 1
				Response.Write(strTxtSun)
			Case 2
				Response.Write(strTxtMon)
			Case 3
				Response.Write(strTxtTue)
			Case 4
				Response.Write(strTxtWed)
			Case 5
				Response.Write(strTxtThu)
			Case 6
				Response.Write(strTxtFri)
			Case 7
				Response.Write(strTxtSat)
		End Select
		
		Response.Write("</div>")
		
		Response.Write("<div style=""padding:4px"">")
		
		
		
		'See if we have an event to display
		If isArray(saryBirthdays) AND intDay => 1 AND intDay <= intMaxNoMonthDays Then
			
			'Initlise the loop array
			intDbArrayLoop = 0
			intNoOfBirthdays = 0
			intAge = 0
			strMemBirthdays = ""
			
			'Loop through the events array
			Do While intDbArrayLoop <= Ubound(saryBirthdays,2)
			
				'If bitrhday is found for this date display it
				If intMonth = Month(saryBirthdays(2,intDbArrayLoop)) AND intDay = Day(saryBirthdays(2,intDbArrayLoop)) Then 
					
					'If we have been around once before then place a comma between the entries
					If intAge <> 0 Then strMemBirthdays = strMemBirthdays & ", "
					
					'Calculate the age (use months / 12 as counting years is not accurate) (use FIX to get the whole number)
					intAge = Fix(DateDiff("m", saryBirthdays(2,intDbArrayLoop), intYear & "-" & intMonth & "-" & intDay)/12)
					
					'Place the bitrhdays into a string to show as as title for the link
					strMemBirthdays = strMemBirthdays & saryBirthdays(1,intDbArrayLoop) & "(" & intAge & ")"
					
					'Initilise variables for the first birthday, it only 1 birthday found we display this data
					strUserName = saryBirthdays(1,intDbArrayLoop)
					lngUserProfile = saryBirthdays(0,intDbArrayLoop)
					
					'Increment the number of birthdays
					intNoOfBirthdays = intNoOfBirthdays + 1
				End If
			
				'Move to next array position
				intDbArrayLoop = intDbArrayLoop + 1
			Loop
			
			'Write the HTML for the date
			'If 1 birthday display the user birthday
			If intNoOfBirthdays = 1 Then
				Response.Write("<em><img src=""" & strImagePath & "calendar_birthday." & strForumImageType & """ alt=""" & strTxtBirthdays & """ title=""" & strTxtBirthdays & """ /> <a href=""member_profile.asp?PF=" & lngUserProfile & strQsSID2 & """>" & strUserName & "</a> (" & intAge & ")</em>")
			'If more than 1 birhday display the number
			ElseIf intNoOfBirthdays > 1 Then
				Response.Write("<em><img src=""" & strImagePath & "calendar_birthday." & strForumImageType & """ alt=""" & strTxtBirthdays & """ title=""" & strTxtBirthdays & """ /> <a href=""calendar_week.asp?M=" & intMonth & "&Y=" & intYear & "&W=" & intWeekLoop & strQsSID2 & """ title=""" & strMemBirthdays & """>" & intNoOfBirthdays & " " & strTxtBirthdays & "</a></em>")
			End If
		End If
		
		
		
		
		'See if we have an event to display
		If isArray(sarryEvents) AND intDay => 1 AND intDay <= intMaxNoMonthDays Then
			
			'Initlise the loop array
			intDbArrayLoop = 0
			
			'Loop through the events array
			Do While intDbArrayLoop <= Ubound(sarryEvents,2)
			
				'If there isn't an end date set, set the end date as the event start date to prevent errors
				If isDate(sarryEvents(3,intDbArrayLoop)) = False Then sarryEvents(3,intDbArrayLoop) = sarryEvents(2,intDbArrayLoop)
			
				'If an event is found for this date display it
				If CDate(intYear & "-" & intMonth & "-" & intDay) >= CDate(sarryEvents(2,intDbArrayLoop)) AND CDate(intYear & "-" & intMonth & "-" & intDay) <= CDate(sarryEvents(3,intDbArrayLoop)) Then
					
					'Read the event details
					strSubject = sarryEvents(1,intDbArrayLoop)
					strMessage = sarryEvents(4,intDbArrayLoop)
					
					'Remove HTML from message for subject link title
					strMessage = removeHTML(strMessage, 100, true)
					
					'Trim the subject
					strSubject = TrimString(strSubject, 25)
			
					'Clean up input to prevent XXS hack
					strMessage = formatInput(strMessage)
					strSubject = formatInput(strSubject)
					
					'Place in a <br /> if birthdays are displayed
					If intNoOfBirthdays > 0 Then Response.Write("<br />")
					
					'Write the HTML for the event
					Response.Write("<img src=""" & strImagePath & "calendar_event." & strForumImageType & """ alt=""" & strTxtEvent & """ title=""" & strTxtEvent & """ /> <a href=""forum_posts.asp?TID=" & sarryEvents(0,intDbArrayLoop) & strQsSID2 & SeoUrlTitle(strSubject, "&title=") & """ title=""" & strMessage & """>" & strSubject & "</a><br />")
				End If
			
				'Move to next array position
				intDbArrayLoop = intDbArrayLoop + 1
			Loop
		End If
		
		Response.Write("</div>")
	Next
   
%></td>
 </tr><%
 
	'If we have run out of weeks in this month exit loop
	If intMaxNoMonthDays =< intDay Then Exit For
Next

%>
</table>
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
   <td align="right" nowrap><a href="calendar.asp?<% 
   
   	'Calculate if we need to chnage years
   	If intMonth = 1 Then 
		Response.Write("M=12&Y=" & intYear-1) 
	Else 
		Response.Write("M=" & intMonth-1 & "&Y=" & intYear) 
	End If
	
	%><% = strQsSID2 %>">&lt;&lt; <% = strTxtPrevious %></a>&nbsp;
    <select onchange="linkURL(this)" name="calSelect" id="calSelect"><%
    
    	'Setup list options
    	For intJumpLoop = 1 TO 12
    	
    		Response.Write(vbCrLf & "     <option value=""calendar.asp?M=" & intJumpLoop & "&Y=" & intYear & strQsSID2&"""")
    		If intJumpLoop = intMonth Then Response.Write(" selected")
    		Response.Write(">" & getMonthName(intJumpLoop) & "</option>")
    		
    	Next
    
     
%>
    </select>
   &nbsp;<a href="calendar.asp?<% 
   
   	'Calculate if we need to chnage years
   	If intMonth = 12 Then 
   		Response.Write("M=1&Y=" & intYear+1) 
   	Else 
   		Response.Write("M=" & intMonth+1 & "&Y=" & intYear) 
   	End If
   	
   	%><% = strQsSID2 %>"><% = strTxtNext %> &gt;&gt;</a></td>
  </tr>
</table>
<table cellspacing="3" cellpadding="0" align="center" class="basicTable">
 <tr valign="top">
  <td width="25%"><%

'Call the sub procedure to write the calendar
If intMonth = 1 Then
	Call displayMonth(12, intYear-1, 1)
Else
	Call displayMonth(intMonth-1, intYear, 1)
End If

%>
  </td>
  <td width="25%"><%

'Call the sub procedure to write the calendar
If intMonth = 12 Then
	Call displayMonth(1, intYear+1, 1)
Else
	Call displayMonth(intMonth+1, intYear, 1)
End If

%>
  </td>
  <td width="50%">&nbsp;</td>
 </tr>
</table>
<br />
<div align="center"><%

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	
	If blnTextLinks = True Then
		Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
	Else
  		Response.Write("<a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank"" ><img src=""webwizforums_image.asp"" border=""0"" title=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ alt=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ /></a>")
	End If

	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")
End If
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

'Display the process time
If blnShowProcessTime Then Response.Write "<span class=""smText""><br /><br />" & strTxtThisPageWasGeneratedIn & " " & FormatNumber(Timer() - dblStartTime, 3) & " " & strTxtSeconds & "</span>"
%>
</div>
<!-- #include file="includes/footer.asp" -->
