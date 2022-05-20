<% Option Explicit %>
<!--#include file="admin_common.asp" -->
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




Session.Timeout = 90

'Set the response buffer to true as we maybe redirecting
Response.Buffer = True


'If in demo mode redirect
If blnDemoMode Then
	Call closeDatabase()
	Response.Redirect("admin_web_wiz_forums_premium.asp" & strQsSID1)
End If



'Dimension variables
Dim intNoOfDays			'Holds the number of days to delete posts from
Dim dtmSelectedDate


'Check the form ID to prevent XCSRF
Call checkFormID(Request.Form("formID"))


'get teh number of days to delte from
intNoOfDays = IntC(Request.Form("days"))


'set up db dates
dtmSelectedDate = internationalDateTime(DateAdd("d", -intNoOfDays, now()))

'If SQL server remove dash (-) from the ISO international date to make SQL Server safe
If strDatabaseType = "SQLServer" Then dtmSelectedDate = Replace(dtmSelectedDate, "-", "", 1, -1, 1)

'If acess used # around dates
If strDatabaseType = "Access" Then
	dtmSelectedDate = "#" & dtmSelectedDate & "#"
Else
	dtmSelectedDate = "'" & dtmSelectedDate & "'"
End If




'Initalise the strSQL variable with an SQL statement to delete the private messages from the database
strSQL = "DELETE FROM " & strDbTable & "PMMessage "
strSQL = strSQL & "WHERE " & strDbTable & "PMMessage.PM_Message_date < " & dtmSelectedDate & ";"

'Delete the topics
adoCon.Execute(strSQL)

'Reset Server Objects
Call closeDatabase()

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Batch Delete Private Messages</title>

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
%>
  
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center">
 <h1>Batch Delete Private Messages</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br>
  <br>
  <br>
  <br>
  <br />
  Private Messages have been Deleted.<br />
 </p>
</div>
<!-- #include file="includes/admin_footer_inc.asp" -->
