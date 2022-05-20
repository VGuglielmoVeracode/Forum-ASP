<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="language_files/chat_room_language_file_inc.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
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





'Set the buffer to true
Response.Buffer = True

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"



Dim strUsername
Dim strPassword


'If a login is displayed then switch to HTTPS	
If intGroupID = 2 Then  blnSslEnabledPage = True


'Save the author ID to the database so it can be read in by the chat server
Call saveSessionItem("AID", lngLoggedInUserID)


	



'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers(strTxtChatRoom, "", "chat.asp", 0)
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""chat.asp" & strQsSID1 & """>" & strTxtChatRoom & "</a>"


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtChatRoom %></title>

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

<script language="javascript" src="includes/default_javascript_v9.js" type="text/javascript"></script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body style="margin:0px;" OnLoad="self.focus();">
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtChatRoom %></h1></td>
</tr>
</table>
<br /><%

'Clear server objects
Call closeDatabase()

'If chat room not enabled
If blnChatRoom = False Then
%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><% = strTxtYouAreNotPermittedToUseThisChatRoom %></td>
 </tr>
</table><%

'If not logged in then show login form
ElseIf intGroupID = 2 Then
	
	%><!--#include file="includes/login_form_inc.asp" --><%

'Else display the message
Else

%>
<table cellspacing="1" cellpadding="3" class="tableBorder" align="center" style="table-layout: fixed;">
 <tr class="tableLedger">
  <td width="75%"><% = strTxtChatRoom %></td>
  <td><% = strTxtOnlineMembers %></td>
 </tr>
 <tr class="ChatTableRow">
  <td valign="top">
   <!-- Start Chat Room -->
    <div class="ChatRoomBox" id="ChatRoomBox"><img alt="" src="<% = strImagePath %>wait16.gif" style="vertical-align:text-top" width="16" height="16" /> <% = strTxtConnecting %></div>
   <!-- End Chat Room -->
  </td>
  <td valign="top">
   <!-- Start Chat Members -->
    <div class="ChatMembersBox" id="ChatMembersBox"><img alt="" src="<% = strImagePath %>wait16.gif" style="vertical-align:text-top" width="16" height="16" /> <% = strTxtConnecting %></div>
   <!-- End Chat Members -->
  </td>
 </tr>
 <tr class="tableBottomRow">
  <td colspan="2">
   <input id="SID" type="hidden" value="<% = strSessionID %>" />
   <input id="message" type="text" size="90" maxlength="250" onkeyup="keyup(event);" />
   <input type="button" value="Submit" onclick="chatWrite();" />
  </td>
 </tr>
</table>
<script language="javascript" src="includes/chat_javascript.js" type="text/javascript"></script>
<div id="ajaxPingAlive">&nbsp;</div>
<script  language="JavaScript">pingAlive(); function pingAlive(){getAjaxData('ajax_session_alive.asp?Ping=alive<% = strQsSID3 %>','ajaxPingAlive'); p = setTimeout('pingAlive()',120000);}</script><%

End If

%>
<br />
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="center"><a href="#" onclick="javascript:window.close();" ><% = strTxtCloseWindow %></a><br />
  </td>
 </tr>
</table>

<div align="center"><%


'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

'Display the process time
If blnShowProcessTime Then Response.Write "<span class=""smText""><br /><br />" & strTxtThisPageWasGeneratedIn & " " & FormatNumber(Timer() - dblStartTime, 3) & " " & strTxtSeconds & "</span>"
%>
</div>
</body>