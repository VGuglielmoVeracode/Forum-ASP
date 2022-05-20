<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
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




'Set the response buffer to true as we maybe redirecting and setting a cookie
Response.Buffer = True

Dim strMode

'Initialial varaibles
blnSslEnabledPage = True

'read in the mode of the page
strMode = Request.QueryString("TP")


'Release server objects
Call closeDatabase()


'Set bread crumb trail
If strMode = "UPD" OR strMode = "DEL" Then 
	strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtEditProfile 
Else 
	strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtRegisterNewUser
End If


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% If strMode = "UPD" OR strMode = "DEL" Then Response.Write(strTxtEditProfile) Else Response.Write(strTxtRegisterNewUser) %></title>

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

%>

<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% If strMode = "UPD" OR strMode = "DEL" Then Response.Write(strTxtEditProfile) Else Response.Write(strTxtRegisterNewUser) %></h1></td>
 </tr>
</table><%

'If this is an update and email notify is on show link to email subcriptions
If strMode = "UPD" OR strMode = "DEL" Then 

%>
<table class="basicTable" cellspacing="0" cellpadding="0" align="center"> 
 <tr> 
  <td class="tabTable">
   <a href="member_control_panel.asp<% If strMode = "A" Then Response.Write("?PF=" & lngEmailUserID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtControlPanel %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>member_control_panel.<% = strForumImageType %>" border="0" alt="<% = strTxtControlPanel %>" /> <% = strTxtControlPanel %></a>
   <a href="register.asp<% If strMode = "A" Then Response.Write("?PF=" & lngEmailUserID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtProfile2 %>" class="tabButtonActive">&nbsp;<img src="<% = strImagePath %>profile.<% = strForumImageType %>" border="0" alt="<% = strTxtProfile2 %>" /> <% = strTxtProfile2 %></a><%
 
	If blnEmail Then

%>
   <a href="email_notify_subscriptions.asp<% If strMode = "A" Then Response.Write("?PF=" & lngEmailUserID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtSubscriptions %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>subscriptions.<% = strForumImageType %>" border="0" alt="<% = strTxtSubscriptions %>" /> <% = strTxtSubscriptions %></a><%
	End If


	'Only disply other links if not in admin mode
	If strMode <> "A" AND blnActiveMember AND blnPrivateMessages Then

%>
   <a href="pm_buddy_list.asp<% = strQsSID1 %>" title="<% = strTxtBuddyList %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>buddy_list.<% = strForumImageType %>" border="0" alt="<% = strTxtBuddyList %>" /> <% = strTxtBuddyList %></a><%

	End If
	

	'If file/image uploads
	If blnAttachments OR blnImageUpload Then

%>
   <a href="file_manager.asp<% If strMode = "A" Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtFileManager %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>file_manager.<% = strForumImageType %>" border="0" alt="<% = strTxtFileManager %>" /> <% = strTxtFileManager %></a><%

	End If


%>
  </td>
 </tr>
</table>
<br /><%

End If

%>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
   <td colspan="2" align="left"><% If strMode = "UPD" OR strMode = "DEL" Then Response.Write(strTxtEditProfile) Else Response.Write(strTxtRegisterNewUser) %></td>
  </tr>
  <tr class="tableRow">
   <td><%

'If this is a re-activation then tell the member
If strMode = "REACT" Then

	Response.Write(strTxtYourEmailAddressHasBeenChanged)

'Else member updating profile
ElseIf strMode = "UPD" Then
	
	%>
    <strong><% = strTxtYourProfileHasBeenUpdated %></strong>
    <br />
    <br /><% =  strTxtYouCanAccessCP %> <a href="member_control_panel.asp<% = strQsSID1 %>"><% =  strTxtMemberCPMenu %></a><%

'Else if the admin has deleted the member disply the delete msg
ElseIf strMode = "DEL" Then
	
	%>
    <strong><% = strTxtTheMemberHasBeenDleted %></strong><%

'Else welcome new user
Else
	%>
    <strong><% = strTxtThankYouForRegistering & " " & strMainForumName %></strong>
    <br />
    <br /><% =  strTxtYouCanAccessCP %> <a href="member_control_panel.asp<% = strQsSID1 %>"><% =  strTxtMemberCPMenu %></a><%
	
End If 


'If this is an admin activation then tell the member
If strMode = "MACT" Then

 	Response.Write(vbCrLf & "    <br /><br />" & strTxtYouAdminNeedsToActivateYourMembership)

ElseIf strMode = "REACT" Then

 	Response.Write(vbCrLf & "    <br /><br />" & strTxtYouShouldReceiveAReactivateEmail)

'Else welcome the new member
ElseIf strMode = "ACT" Then

 	Response.Write(vbCrLf & "    <br /><br />" & strTxtYouShouldReceiveAnEmail)
End If
%>
    <br />
    <br /><a href="default.asp<% = strQsSID1 %>"><% = strTxtReturnToDiscussionForum %></a><%

'If this person needs to activate the account
If strMode = "REACT" OR strMode = "ACT" Then
 %>
    <br /><br /><% = strTxtIfErrorActvatingMembership & " " & strMainForumName & " " & " <a href=""mailto:" & strForumEmailAddress & """>" & strTxtForumAdministrator & "</a>." %><br /><%
End If
%>
    <br />
   </td>
  </tr>
</table>
<br />
<div align="center">
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