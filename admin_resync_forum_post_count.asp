<% @ Language=VBScript %>
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




'Set the response buffer to true
Response.Buffer = True 


	
'Initalise the strSQL variable with an SQL statement to get the all the forum id's
strSQL = "SELECT " & strDbTable & "Forum.Forum_ID FROM " & strDbTable & "Forum;"
	
'Query the database
rsCommon.Open strSQL, adoCon

'Loop through all the forums to re-sync them
Do While NOT rsCommon.Eof

	'Update topic and post count
	updateForumStats(CInt(rsCommon("Forum_ID")))
	
	'Move to the next record
	rsCommon.MoveNext
Loop


'Reset server objects
rsCommon.Close
Call closeDatabase()
%>  
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Re-sync the Topic and Post Count</title>
<meta name="generator" content="Web Wiz Forums" />

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
     	
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center"> 
 <h1>Re-sync the Topic and Post Count</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
 </p>
 <br>
 <span class="lgText">The Topic and Post Count for all Forums have been re-synchronised.</span></div>
<!-- #include file="includes/admin_footer_inc.asp" -->