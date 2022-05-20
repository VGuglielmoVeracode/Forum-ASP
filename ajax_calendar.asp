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

Response.ContentType = "text/html"

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-store"


'Dimension variables
Dim sarryEvents		'Holds the event recordset in an array
Dim strSubject		'Holds the event subject
Dim strMessage		'Holds the event message
Dim intMonth		'Holds the integer value for the month
Dim intYear		'Holds the year
Dim intDay		'Holds the day
Dim dtmNow		'Holds the present date according to any server off-set
Dim dtmDbStartDate	'Holds the database search start date
Dim dtmDbEndDate	'Holds the database search end date
Dim intDbArrayLoop	'Database array loop counter



'If the calendar system is disabled send the user away
If blnCalendar = false Then 
	
	'Clean up
	Call closeDatabase()
	
	'Send to default page
	Response.Redirect("default.asp" & strQsSID1)
End If



'Intilise variables
dtmNow = getNowDate()
intYear = Year(dtmNow)
intMonth = Month(dtmNow)



'Calculate the db serach start date
dtmDbStartDate = internationalDateTime(intYear & "-" & intMonth & "-1")

'Calculate the db serach end date
dtmDbEndDate = internationalDateTime(intYear & "-" & intMonth & "-" & getMonthDayNo(intMonth, intYear))


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
		
	
'Clean up
Call closeDatabase()


'Display calendar for this month
Call displayMonth(intMonth, intYear, 1)

%>