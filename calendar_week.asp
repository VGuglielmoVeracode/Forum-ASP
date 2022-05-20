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
Dim intWeek		'Holds the week to view
Dim intFistDayOfMonth	'Holds the first day of the month as an interger
Dim intWeekLoop		'Week Loop counter
Dim intDayLoopCounter	'Day loop counter
Dim intDay		'Holds the day
Dim dtmNow		'Holds the present date according to any server off-set
Dim intMonthSmCalendar	'Holds the month for the small calendars
Dim intView		'Holds the type of view (1=monh, 2=week, 3=year)
Dim dtmDbStartDate	'Holds the database search start date
Dim dtmDbEndDate	'Holds the database search end date
Dim intDbArrayLoop	'Database array loop counter
Dim intAge		'Holds the age for birthdays
Dim blnShowBirthdays	'Holds if birthdays are hidden or not




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
If Request.QueryString("V") = "2" Then
	Response.Redirect("calendar_year.asp?Y=" & Request.QueryString("Y") & strQsSID3)
End If


'Intilise variables
blnShowBirthdays = CBool(showBirthdays())
dtmNow = getNowDate()
intDay = 1
intView = 2


'Read in the year to view
If IntC(Request.QueryString("Y")) => 2001 AND  IntC(Request.QueryString("Y")) =< Cint(Year(dtmNow)+5) Then
	intYear = IntC(Request.QueryString("Y"))
'Else use this year as the month to view
Else
	Response.Redirect("calendar_week.asp?M=" & IntC(Request.QueryString("M")) & "&Y=" & Year(dtmNow) & "&W=" & IntC(Request.QueryString("W")) & strQsSID3)
End If


'Read in the month to view
If IntC(Request.QueryString("M")) => 1 AND  IntC(Request.QueryString("M")) =< 12 Then
	intMonth = IntC(Request.QueryString("M"))
'Else use this month as the month to view
Else
	Response.Redirect("calendar_week.asp?M=" &  Month(dtmNow) & "&Y=" & IntC(Request.QueryString("Y")) & "&W=" & IntC(Request.QueryString("W")) & strQsSID3)
End If

'Read in the week to view
If IntC(Request.QueryString("W")) => 1 AND  IntC(Request.QueryString("W")) =< 6 Then
	intWeek = IntC(Request.QueryString("W"))
'Else use this month as the month to view
Else
	Response.Redirect("calendar_week.asp?M=" &  IntC(Request.QueryString("M")) & "&Y=" & IntC(Request.QueryString("Y")) & "&W=1" & strQsSID3)
End If





'Get the first day of the month (use internation ISO date fomat (yyyy-mm-dd) for server compatibility)
intFistDayOfMonth = WeekDay(intYear & "-" & intMonth & "-01")

'Initilise the month variable for the small calendars
intMonthSmCalendar = intMonth

'Calulate the start date for the week
Select Case intWeek
	Case 2 
		intDay = 8 - (intFistDayOfMonth - 1)
	Case 3 
		intDay = 15 - (intFistDayOfMonth - 1)
	Case 4 
		intDay = 22 - (intFistDayOfMonth - 1)
	Case 5 
		intDay = 29 - (intFistDayOfMonth - 1)
	Case 6 
		intDay = 36 - (intFistDayOfMonth - 1)
End Select



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
If blnDisplayBirthdays AND blnShowBirthdays Then Call getBirthdays(intMonth, intWeek)

'SQL Query Array Look Up table
'0 = Author_ID
'1 = Username
'2 = DOB
	


'Clean up
Call closeDatabase()


'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers(strTxtViewing & " " & strTxtEvents, strMonth & " " & intYear, "calendar_week.asp?Y=" & intYear & "&M=" & intMonth & "&W=" & intWeek, 0)
End If

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
<table cellspacing="5" cellpadding="0" align="center" class="basicTable">
 <tr>
  <td nowrap><!--#include file="includes/calendar_jump_inc.asp" --></td>
 </tr>
</table>
<table cellspacing="2" cellpadding="0" class="basicTable" align="center">
  <tr>
    <td width="25%" height="10"><%

'Display calendar of previous month
If intMonthSmCalendar = 1 Then
	Call displayMonth(12, intYear-1, 1)
Else
	Call displayMonth(intMonthSmCalendar-1, intYear, 1)
End If

%></td>
    <td width="75%" rowspan="3" valign="top">
     <table cellspacing="1" cellpadding="3" class="tableBorder" align="center"><%

'Loop through the days of the week
For intDayLoopCounter = 1 TO 7 


	'If the first day on the month is not on a sunday, the we also need to show the last days of the previous month
	If intFistDayOfMonth > 1 AND intDayLoopCounter = 1 AND intWeek = 1 Then
		
		'Go back to last month (if not January (month=1))
		If intMonth = 1 Then
			intMonth = 12
			intYear = intYear - 1
		'Else just go to the next month
		Else 
			intMonth = intMonth - 1
		End If

		'Calculate the day on the previous month to show from
		intDay = getMonthDayNo(intMonth, intYear) - (intFistDayOfMonth - 2)
	End If


	'If we have reached the end of the month, jump to next month
	If intDay > getMonthDayNo(intMonth, intYear) Then
		
		'See if we need to jump to next year
		If intMonth = 12 Then
			intMonth = 1
			intYear = intYear + 1
		'Else just go to the next month
		Else 
			intMonth = intMonth + 1
		End If
		
		'Reset day
		intDay = 1
	End If
	
	
	'If this is the first day in the month display a ledger
	If intDay = 1 OR intDayLoopCounter = 1 Then
		'Display the ledger with the month
		Response.Write(vbCrLf & "      <tr class=""tableLedger"">" & _
		vbCrLf & "      <td colspan=""2""><a href=""calendar.asp?M=" & intMonth & "&Y=" & intYear & strQsSID2 & """>" & getMonthName(intMonth) & " " & intYear & "</a></td>" & _
		vbCrLf & "       </tr>")
	End If
	
	
	'Write the day
	Response.Write(vbCrLf & "      <tr class=""calLedger"">" & _
	vbCrLf & "       <td colspan=""2"">")
	Select Case intDayLoopCounter
		Case 1
			Response.Write(strTxtSunday)
		Case 2
			Response.Write(strTxtMonday)
		Case 3
			Response.Write(strTxtTuesday)
		Case 4
			Response.Write(strTxtWednesday)
		Case 5
			Response.Write(strTxtThursday)
		Case 6
			Response.Write(strTxtFriday)
		Case 7
			Response.Write(strTxtSaturday)
	End Select	
	Response.Write("</td>" & _
	vbCrLf & "      </tr>")	
	
	'Write the date and event details
	Response.Write(vbCrLf & "      <tr height=""41"" class=""calDateCell"">" & _
	vbCrLf & "       <td width=""41"" ")
	
	'If today place a red border around the day
	If intMonth = Month(dtmNow) AND intDay = Day(dtmNow) AND intYear = Year(dtmNow) Then 
		Response.Write(" class=""calTodayCell""")
	End If
	
	Response.Write(" align=""center"" style=""font-size:17px"">" & intDay & "</td>" & _
	vbCrLf & "       <td valign=""top"">")
	
	
	'See if we have any birthdays to display
	If isArray(saryBirthdays) AND intDay => 1 Then
		
		'Initlise the loop array
		intDbArrayLoop = 0
		intAge = 0
		
		'Loop through the birthdays array
		Do While intDbArrayLoop <= Ubound(saryBirthdays,2)
		
			'If an bitrhday is found for this date display it
			If intMonth = Month(saryBirthdays(2,intDbArrayLoop)) AND intDay = Day(saryBirthdays(2,intDbArrayLoop)) Then 
				
				'If we have been around once before then place a comma between the entries
				If intAge = 0 Then 
					Response.Write("<img src=""" & strImagePath & "calendar_birthday." & strForumImageType & """ alt=""" & strTxtBirthdays & """ title=""" & strTxtBirthdays & """ /> ")
				Else
					Response.Write(", ")
				End If
				
				'Calculate the age (use months / 12 as counting years is not accurate) (use FIX to get the whole number)
				intAge = Fix(DateDiff("m", saryBirthdays(2,intDbArrayLoop), intYear & "-" & intMonth & "-" & intDay)/12)
				
				'Write the HTML for the birthday
				Response.Write("<em><a href=""member_profile.asp?PF=" & saryBirthdays(0,intDbArrayLoop) & strQsSID2 &  """>" & saryBirthdays(1,intDbArrayLoop) & "</a> (" & intAge & ")</em>")
			End If
		
			'Move to next array position
			intDbArrayLoop = intDbArrayLoop + 1
		Loop
	End If
	
	
	'See if we have an event to display
	If isArray(sarryEvents) AND intDay => 1 Then
		
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
				
				'Clean up input to prevent XXS hack
				strMessage = formatInput(strMessage)
				strSubject = formatInput(strSubject)
				
				'Place in a <br /> if birthdays are displayed
				If intAge > 0 Then Response.Write("<br />")
				
				'Write the HTML for the event
				Response.Write("<img src=""" & strImagePath & "calendar_event." & strForumImageType & """ alt=""" & strTxtEvent & """ title=""" & strTxtEvent & """ /> <a href=""forum_posts.asp?TID=" & sarryEvents(0,intDbArrayLoop) & strQsSID2 & SeoUrlTitle(strSubject, "&title=") &  """ title=""" & strMessage & """>" & strSubject & "</a><br />")
			End If
		
			'Move to next array position
			intDbArrayLoop = intDbArrayLoop + 1
		Loop
	End If
	
	Response.Write("</td>" & _
	vbCrLf & "      </tr>")	
	
	
	'Increment the day by 1	
	intDay = intDay + 1
Next

%>

    </table>
   </td>
  </tr>
  <tr>
    <td height="10"><%

'Display calendar for this month
Call displayMonth(intMonthSmCalendar, intYear, 1)

%></td>
  </tr>
  <tr>
    <td valign="top"><%

'Display calendar of next month
If intMonthSmCalendar = 12 Then
	Call displayMonth(1, intYear + 1, 1)
Else
	Call displayMonth(intMonthSmCalendar + 1, intYear, 1)
End If

%></td>
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
