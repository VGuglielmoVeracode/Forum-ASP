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


Rem Dimension variables
Dim sarryEvents		'Holds the event recordset in an array
Dim strSubject		'Holds the event subject
Dim intYear		'Holds the year
Dim dtmNow		'Holds the present date according to any server off-set
Dim intLoop		'Loop counter
Dim intRowLoop		'Loop counter to build table
Dim intCellLoop		'Loop counter to build cells
Dim intMonth		'Holds the month to show
Dim intView		'Holds the type of view (1=monh, 2=week, 3=year)
Dim dtmDbStartDate	'Holds the database search start date
Dim dtmDbEndDate	'Holds the database search end date
Dim intDbArrayLoop	'Database array loop counter
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



'Intilise variables
dtmNow = getNowDate()
blnShowBirthdays = CBool(showBirthdays())
intMonth = 1
intView = 3


'Read in the year to view
If IntC(Request.QueryString("Y")) => 2001 AND  IntC(Request.QueryString("Y")) =< Cint(Year(dtmNow)+5) Then
	intYear = IntC(Request.QueryString("Y"))
'Else use this year as the month to view
Else
	Response.Redirect("calendar_year.asp?Y=" & Year(dtmNow) & "&DB=" & Request.QueryString("DB") & strQsSID3)
End If







'Calculate the db serach start date
dtmDbStartDate = internationalDateTime(intYear & "-01-01")

'Calculate the db serach end date
dtmDbEndDate = internationalDateTime(intYear & "-12-31")

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

'Read in any events from the database and place them into an array to display later
'Call the sub procedure to get the events from the database
Call getEvents()
		
'SQL Query Array Look Up table
'0 = Topic_ID
'1 = Subject
'2 = Hide
'3 = Event_date
'4 = Message


'Clean up
Call closeDatabase()




'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers(strTxtViewing & " " & strTxtEvents, intYear, "calendar_year.asp?Y=" & intYear, 0)
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""calendar.asp" & strQsSID1 & """>" & strTxtEvents & "</a>" & strNavSpacer & intYear

'Status bar tools
strStatusBarTools = strStatusBarTools & "&nbsp;<a href=""RSS_calendar_feed.asp" & strQsSID1 & """ target=""_blank""><img src=""" & strImagePath & "rss." & strForumImageType & """ border=""0"" alt=""" & strTxtRSS & " - " & strTxtLatestEventFeed & """ title=""" & strTxtRSS & " - " & strTxtLatestEventFeed & """ /></a>"



%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strMainForumName & " " & strTxtEvents & " - " & intYear %></title>
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
vbCrLf & "//-->" & vbCrLf & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

'If RSS feed is enabled then have an alt link to it for browers that support RSS Feeds
If blnRSS Then Response.Write(vbCrLf & "<link rel=""alternate"" type=""application/rss+xml"" title=""RSS 2.0"" href=""RSS_calendar_feed.asp" & strQsSID1 & """ />")
%>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtEvents & " - " & intYear %></h1></td>
 </tr>
</table>
<table cellspacing="5" cellpadding="0" align="center" class="basicTable">
 <tr >
  <td><!--#include file="includes/calendar_jump_inc.asp" --></td>
 </tr>
</table>
<table cellspacing="3" cellpadding="0" align="center" class="basicTable"><%

'Loop through rows in the table
For intRowLoop = 1 TO 3

	'Create the table row
	Response.Write(vbCrLf & " <tr>")
	
	'Loop through the table cells
	For intCellLoop = 1 TO 4
	
		'Create the table cell
		Response.Write(vbCrLf & "  <td width=""25%"">")
		
		'Create the calendar
		Call displayMonth(intMonth, intYear, 0)

		'Close table cell
		Response.Write(vbCrLf & "  </td>")

		'Increment month number by 1
		intMonth = intMonth + 1
	Next
	
	'Close table row
	Response.Write(vbCrLf & " </tr>")
Next

%>
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
