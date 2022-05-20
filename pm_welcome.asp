<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="language_files/pm_language_file_inc.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
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

'Declare variables
Dim saryPmMessage		'db recordset holding any new pm's since last vist
Dim intNewPMs

blnSslEnabledPage = True


'If Priavte messages are not on then send them away
If blnPrivateMessages = False Then 
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
End If


'If the user is not allowed then send them away
If intGroupID = 2 OR blnActiveMember = False OR blnBanned Then 
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If


'Get the number of new pm's this user has
intNewPMs = intNoOfPms





'Now get the date of the last PM
strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "PMMessage.PM_Message_date, " & strDbTable & "Author.Username " & _
"FROM " & strDbTable & "Author" & strDBNoLock & ", " & strDbTable & "PMMessage" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Author.Author_ID=" & strDbTable & "PMMessage.From_ID " & _
	"AND " & strDbTable & "PMMessage.Read_Post=" & strDBFalse & " " & _
	"AND " & strDbTable & "PMMessage.Author_ID=" & lngLoggedInUserID & " " & _
"ORDER BY " & strDbTable & "PMMessage.PM_Message_date DESC" & strDBLimit1 & ";"

		
'Query the database
rsCommon.Open strSQL, adoCon

'Place the recordset into an array
If NOT rsCommon.EOF Then
	
	'Read in PM data from recordset
	saryPmMessage = rsCommon.GetRows()
	
End If

'Close the recordset
rsCommon.Close



'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers("", strTxtPrivateMessenger & " " & strTxtWelcome, "pm_welcome.asp", 0)
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""pm_welcome.asp" & strQsSID1 & """>" & strTxtPrivateMessenger & "</a>" & strNavSpacer & strTxtWelcome

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtPrivateMessenger & " - " & strTxtWelcome %></title>

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

<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />   	
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtPrivateMessenger %></h1></td>
</tr>
</table>
<br />
<table class="basicTable" cellspacing="0" cellpadding="0" align="center"> 
 <tr> 
  <td class="tabTable">
   <a href="pm_welcome.asp<% = strQsSID1 %>" title="<% = strTxtPrivateMessenger %>" class="tabButtonActive">&nbsp;<img src="<% = strImagePath %>messenger.<% = strForumImageType %>" border="0" alt="<% = strTxtPrivateMessenger %>" /> <% = strTxtMessenger %></a>
   <a href="pm_inbox.asp<% = strQsSID1 %>" title="<% = strTxtPrivateMessenger & " " & strTxtInbox %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>inbox_messages.<% = strForumImageType %>" border="0" alt="<% = strTxtPrivateMessenger & " " & strTxtInbox %>" /> <% = strTxtInbox %></a>
   <a href="pm_outbox.asp<% = strQsSID1 %>" title="<% = strTxtPrivateMessenger & " " & strTxtOutbox %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>sent_messages.<% = strForumImageType %>" border="0" alt="<% = strTxtPrivateMessenger & " " & strTxtOutbox %>" /> <% = strTxtOutbox %></a>
   <a href="pm_new_message_form.asp<% = strQsSID1 %>" title="<% = strTxtNewPrivateMessage %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>new_message.<% = strForumImageType %>" border="0" alt="<% = strTxtNewPrivateMessage %>" /> <% = strTxtNewMessage %></a>
  </td>
 </tr>
</table>
<br />
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td><% = strTxtWelcome & " " & strLoggedInUsername & " " & strTxtToYourPrivateMessenger %></td>
 </tr>
 <tr class="tableRow">
  <td><% = strTxtPmIntroduction %><br /><br /></td>
 </tr>
 <tr class="tableSubLedger">
  <td><% = strTxtInboxStatus %></td>
 </tr> 
 <tr class="tableRow">
  <td><%
'If there are pm's display full inbox icon
If isArray(saryPmMessage) Then
	Response.Write("<img src=""" & strImagePath & "inbox_new_mail.png"" align=""left"" alt=""" & strTxtInboxStatus & """ hspace=""7"" />")
Else
	Response.Write("<img src=""" & strImagePath & "inbox_empty.png"" align=""left"" alt=""" & strTxtInboxStatus & """ hspace=""7"" />")
End If

'If there are pm's display the last pm details
If isArray(saryPmMessage) Then
	Response.Write("<strong>" & strTxtYouHave & " " & intNewPMs & " " & strTxtNewMsgsInYourInbox & "</strong> <a href=""pm_inbox.asp" & strQsSID1 & """>" & strTxtGoToYourInbox & "</a>.<br />")
	Response.Write(strTxtYourLatestPrivateMessageIsFrom & " " & saryPmMessage(1,0) & " " & strTxtSent & " " & DateFormat(saryPmMessage(0,0)) & " " & strTxtAt & " " & TimeFormat(saryPmMessage(0,0)) & "<br />")
Else	
	Response.Write("<br />" & strTxtNoNewMsgsInYourInbox & " <a href=""pm_inbox.asp" & strQsSID1 & """>" & strTxtGoToYourInbox & "</a>")
End If
         %></td>
 </tr>
</table>
<br />    
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td><% = strTxtPrivateMessengerOverview %></td>
 </tr>
 <tr class="tableSubLedger">
  <td><% = strTxtInbox %></td>
 </tr>
 <tr class="tableRow">
  <td><% = strTxtInboxOverview %><br /><br /></td>
 </tr>
 <tr class="tableSubLedger">
  <td><% = strTxtOutbox %></td>
 </tr>
 <tr class="tableRow">
  <td><% = strTxtOutboxOverview %><br /><br /></td>
 </tr>
 <tr class="tableSubLedger">
  <td><% = strTxtNewPrivateMessage %></td>
 </tr>
 <tr  class="tableRow">
  <td><% = strTxtNewMsgOverview %></td>
 </tr>
</table>
<br />
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
   <td>
    <!-- #include file="includes/forum_jump_inc.asp" -->
   </td>
 </tr>
</table>
<div align="center"><br />
<% 
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
<%
'Clean up
Call closeDatabase()
%>