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


'Reset Server Objects
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Web Wiz Forums Premium Edition</title>
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
<style type="text/css">
<!--

.AdTable {
	width: 720px;
	border: 1px solid #E7E7E7;
	background-color: #FFFFFF;
	font-family: Verdana, Arial, Helvetica, sans-serif;
	color : #000000;
	font-weight: normal;
	font-size: 12px;
	text-align: left;
	
}

.AdHeading {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 18px;
	margin-bottom: 3px;
	margin-top: 0px;
}

-->
</style>

<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center">
  <h1>Web Wiz Forums Premium Edition</h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br /><br />
   <table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
   <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong>ALERT</strong>
    <br /><strong><%
    
    'Dirrent message if in demo mode   
If blnDemoMode Then
	Response.Write("Sorry this option is not available in Demo Mode")
Else	
	Response.Write("Sorry this option is only available in the Web Wiz Forums Premium Edition")
End If

%></strong>
  </tr>
</table>
    <br /><br /><br /><br />
</div>
<!--#include file="includes/webwizforums_premium_edition_inc.asp" -->
<br />
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
