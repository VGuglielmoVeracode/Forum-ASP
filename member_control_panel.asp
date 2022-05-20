<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="language_files/cp_language_file_inc.asp" -->
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
Dim lngUserProfileID            'Holds the users ID of the profile to get
Dim strMode			'Holds the mode of the page


blnSslEnabledPage = True


'If the user his not activated their mem
If blnActiveMember = False OR blnBanned Then

        'clean up before redirecting
        Call closeDatabase()

        'redirect to insufficient permissions page
        Response.Redirect("insufficient_permission.asp?M=ACT" & strQsSID3)
End If

'If the user has not logged in then redirect them to the insufficient permissions page
If intGroupID = 2 Then

        'clean up before redirecting
        Call closeDatabase()

        'redirect to insufficient permissions page
        Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If



'Read in the mode of the page
strMode = Trim(Mid(Request.QueryString("M"), 1, 1))


'If this is not an admin but in admin mode then see if the user is a moderator
If blnAdmin = False AND strMode = "A" AND blnModeratorProfileEdit Then
	
	'Initalise the strSQL variable with an SQL statement to query the database
        strSQL = "SELECT " & strDbTable & "Permissions.Moderate " & _
        "FROM " & strDbTable & "Permissions " & _
        "WHERE (" & strDbTable & "Permissions.Group_ID = " & intGroupID & " OR " & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & ") AND  " & strDbTable & "Permissions.Moderate=" & strDBTrue & ";"
               

        'Query the database
         rsCommon.Open strSQL, adoCon

        'If a record is returned then the user is a moderator in one of the forums
        If NOT rsCommon.EOF Then blnModerator = True

        'Clean up
        rsCommon.Close
End If


'Get the user ID of the member being edited if in admin mode
If (blnAdmin OR (blnModerator AND LngC(Request.QueryString("PF")) > 2)) AND strMode = "A" Then
	
	lngUserProfileID = LngC(Request.QueryString("PF"))

'Get the logged in ID number
Else
	lngUserProfileID = lngLoggedInUserID
End If

'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers("", strTxtMemberCPMenu, "member_control_panel.asp", 0)
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtMemberCPMenu

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtMemberCPMenu %></title>

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

<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />  	
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtMemberCPMenu %></h1></td>
</tr>
</table>
<br />
<table class="basicTable" cellspacing="0" cellpadding="0" align="center"> 
 <tr> 
  <td class="tabTable">
   <a href="member_control_panel.asp<% If strMode = "A" Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtControlPanel %>" class="tabButtonActive">&nbsp;<img src="<% = strImagePath %>member_control_panel.<% = strForumImageType %>" border="0" alt="<% = strTxtControlPanel %>" /> <% = strTxtControlPanel %></a>
   <a href="register.asp<% If strMode = "A" Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtProfile2 %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>profile.<% = strForumImageType %>" border="0" alt="<% = strTxtProfile2 %>" /> <% = strTxtProfile2 %></a><%
 
If blnEmail Then

%>
   <a href="email_notify_subscriptions.asp<% If strMode = "A" Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtSubscriptions %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>subscriptions.<% = strForumImageType %>" border="0" alt="<% = strTxtSubscriptions %>" /> <% = strTxtSubscriptions %></a><%
End If


'Only disply other links if not in admin mode
If strMode <> "A" AND blnActiveMember AND blnPrivateMessages Then

%>
   <a href="pm_buddy_list.asp<% = strQsSID1 %>" title="<% = strTxtBuddyList %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>buddy_list.<% = strForumImageType %>" border="0" alt="<% = strTxtBuddyList %>" /> <% = strTxtBuddyList %></a><%

End If


'If member is allowed to upload display link to file manager
If blnAttachments OR blnImageUpload Then

%>
   <a href="file_manager.asp<% If strMode = "A" Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtFileManager %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>file_manager.<% = strForumImageType %>" border="0" alt="<% = strTxtFileManager %>" /> <% = strTxtFileManager %></a><%

End If


%>
  </td>
 </tr>
</table>
<br />
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td colspan="2" align="left"><% = strTxtMemberCPMenu %></td>
 </tr>
 <tr class="tableRow">
  <td>
   <a href="register.asp?FPN=0<% If strMode = "A" AND (blnAdmin OR blnModerator) Then Response.Write("&PF=" & lngUserProfileID & "&M=A") %><% = strQsSID2 %>"><% = strTxtEditProfile %></a><br /><% = strTxtChangeProfile %>.
   <br /><br />
   <a href="register.asp?FPN=1<% If strMode = "A" AND (blnAdmin OR blnModerator) Then Response.Write("&PF=" & lngUserProfileID & "&M=A") %><% = strQsSID2 %>"><% = strTxtRegistrationDetails %></a><br /><% = strTxtChangePassAndEmail %>.
   <br /><br />
   <a href="register.asp?FPN=2<% If strMode = "A" AND (blnAdmin OR blnModerator) Then Response.Write("&PF=" & lngUserProfileID & "&M=A") %><% = strQsSID2 %>"><% = strTxtProfileInfo %></a><br /><% = strTxtChangeProfileInfo %>.
   <br /><br />
   <a href="register.asp?FPN=3<% If strMode = "A" AND (blnAdmin OR blnModerator) Then Response.Write("&PF=" & lngUserProfileID & "&M=A") %><% = strQsSID2 %>"><% = strTxtForumPreferences %></a><br /><% = strTxtChangeForumPreferences %>.<%
          
'email notify is on show link to email subcriptions
If blnEmail Then
%>	
   <br /><br />
   <a href="email_notify_subscriptions.asp<% If strMode = "A" AND (blnAdmin OR blnModerator) Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>"><% = strTxtSubscriptions %></a><br /><% = strTxtAlterEmailSubscriptions %>.<%
End If
         
'If PM is on then show links to PM functions
If blnPrivateMessages AND strMode <> "A" Then
	%>
   <br /><br />
   <a href="pm_buddy_list.asp<% = strQsSID1 %>"><% = strTxtBuddyList %></a><br /><% = strTxtListOfYourForumBuddies %>
   <br /><br />
   <a href="pm_welcome.asp<% = strQsSID1 %>"><% = strTxtPrivateMessenger %></a><br /><% = strTxtReadandSendPMs %><%
End If

'If the user is user is using a banned IP redirect to an error page
If blnAttachments OR blnImageUpload Then
%>
   <br /><br />
   <a href="file_manager.asp<% If strMode = "A" Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>"><% = strTxtFileManager %></a><br /><% = strTxtFileManagerDescription %><%

End If



'If admin/mod mode have link to admin functions
If strMode = "A" AND (blnAdmin OR blnModerator) Then
%>
   <br /><br />
   <a href="register.asp?PF=<% = lngUserProfileID %>&M=A<% = strQsSID2 %>#admin"><% = strTxtAdminAndModFunc %></a><br /><% = strTxtAdminFunctionsTo %>.<%
End If

%>
   <br /><br />
   <a href="member_data.asp?PF=<% = lngUserProfileID %><% = strQsSID2 %>" target="_blank"><% = strTxtViewYourPersonalData %></a><br /><% = strTxtHaveALookAtThePersonalDataTheForumHasOnYou %>.
   <br /><br />
   <a href="help.asp<% = strQsSID1 %>"><% = strTxtForumHelp %></a><br /><% = strTxtForumHelpFilesandFAQsToHelpYou %>.
  </td>
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
<div align="center"><br /><%

'Clean up
Call closeDatabase()

	
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