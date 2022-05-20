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



'Make sure that the forum ID number has a number in it otherwise set it to 0
If isEmpty(intForumID) OR intForumID = "" Then intForumID = 0


'Don't display normal status bar if forum is locked as it will course a loop with the AJAX
If blnForumClosed Then

%>
<table cellspacing="1" cellpadding="3" align="center" class="tableBorder">
 <tr class="tableStatusBar"> 
  <td>
   <div style="float:left;"><% = strBreadCrumbTrail %></div>
   <div style="float:right;"><%

'If the user is admin display a link to the admin menu
If intGroupID = 1 Then Response.Write("&nbsp;&nbsp;<img src=""" & strImagePath & "admin_control_panel." & strForumImageType & """ title=""" & strTxtAdminControlPanel & """ alt=""" & strTxtAdminControlPanel & """ style=""vertical-align: text-bottom"" />&nbsp;<a href=""admin.asp" & strQsSID1 & """>" & strTxtAdmin & "</a>")
	

%></div>
  </td>
 </tr>
</table><%



'Display different status bar for mobile browsers
ElseIf blnMobileBrowser Then
		
%>
<table cellspacing="1" cellpadding="3" align="center" class="tableBorder">
 <tr>
   <td><%
   	
   	'Forum home link for mobile browsers
   	Response.Write ("<img src=""" & strImagePath & "forum_home.png"" alt=""" & strTxtForumHome & """ title=""" & strTxtForumHome & """ /> <a href=""default.asp" & strQsSID1 & """>" & strTxtForumHome & "</a>")
   	
   	'Display link to forum topics home for mobile browsers
   	If intForumID <> 0 Then Response.Write (" <img src=""" & strImagePath & "topic_sm.png"" title=""" & strTxtForum & " " & strTxtTopics & """ alt=""" & strTxtForum & " " & strTxtTopics & """ />  <a href=""forum_topics.asp?FID=" & intForumID & strQsSID2 & SeoUrlTitle(strForumName, "&title=") & """>" & strTxtForum & " " & strTxtTopics & "</a>")

	'News Posts
	Response.Write (" <img src=""" & strImagePath & "active_topics.png"" alt=""" & strTxtActiveTopics & """ title=""" & strTxtActiveTopics & """ /> <a href=""active_topics.asp" & strQsSID1 & """>" & strTxtActiveTopics & "</a>")
	
	'Search
	Response.Write (" <img src=""" & strImagePath & "search.png"" alt=""" & strTxtSearchTheForum & """ title=""" & strTxtSearchTheForum & """ /> <a href=""search_form.asp" & strQsSID1 & """>" & strTxtSearch & "</a>")
	
	
	
	'Member control panel
   	If lngLoggedInUserID <> 0 AND lngLoggedInUserID <> 2 AND blnBanned = False Then Response.Write ("<br /><img src=""" & strImagePath & "member_control_panel.png"" title=""" & strTxtMemberCPMenu & """ alt=""" & strTxtMemberCPMenu & """ /> <a href=""member_control_panel.asp" & strQsSID1 & """>" & strTxtControlPanel & "</a>")
   	
   	'If the user is logged in and there account is active display if they have private messages
	If intGroupID <> 2 AND blnActiveMember AND blnPrivateMessages Then
	
		'Display the number of new pm's
		If intNoOfPms > 0 Then
			Response.Write(" <img src=""" & strImagePath & "new_private_message.png"" title=""" & intNoOfPms & " " & strTxtNewMessages & """ alt=""" & intNoOfPms & " " & strTxtNewMessages & """ /> <a href=""pm_inbox.asp" & strQsSID1 & """>" & strTxtPrivateMessenger & "</a> [" & intNoOfPms & " " & strTxtNew & "]")
		Else
			Response.Write(" <img src=""" & strImagePath & "private_message.png"" title=""0 " & strTxtNewMessages & """ alt=""0 " & strTxtNewMessages & """ /> <a href=""pm_welcome.asp" & strQsSID1 & """>" & strTxtPrivateMessenger & "</a>")
		End If
	End If
	
	'If the user has logged in then the Logged In User ID number will not be 0 and not 2 for the guest account
	If lngLoggedInUserID <> 0 AND lngLoggedInUserID <> 2 Then

		'Don't display logout if windows authentication is enabled or member API
		If blnWindowsAuthentication = False AND (blnMemberAPIDisableAccountControl = False OR strMemberAPILogoutURL <> "") Then 
			Response.Write (" <img src=""" & strImagePath & "logout.png"" alt=""" & strTxtLogOff & """ title=""" & strTxtLogOff & """ /> <a href=""log_off_user.asp?XID=" & getSessionItem("KEY") & strQsSID2 & """>" & strTxtLogOff & "</a>")
		End If

	'Else if the member API is enabled and there are links to the websites on registration and login pages display the links to them
	ElseIf blnMemberAPI AND (strMemberAPIRegistrationURL <> "" OR strMemberAPILoginURL <> "") Then
		
		If strMemberAPIRegistrationURL <> "" Then Response.Write (" <br /><img src=""" & strImagePath & "register.png"" alt=""" & strTxtRegister & """ title=""" & strTxtRegister & """ /> <a href=""" & strMemberAPIRegistrationURL & """>" & strTxtRegister & "</a>")
		If strMemberAPILoginURL <> "" Then Response.Write (" <img src=""" & strImagePath & "login.png"" alt=""" & strTxtLogin & """ title=""" & strTxtLogin & """ /> <a href=""" & strMemberAPILoginURL & """>" & strTxtLogin & "</a>")	
	
	'Else the user is not logged (Don't display logout if windows authentication is enabled or member API)
	ElseIf blnWindowsAuthentication = False AND (blnMemberAPI = False OR blnMemberAPIDisableAccountControl = False) Then
		
		Response.Write (" <br /><img src=""" & strImagePath & "register.png"" alt=""" & strTxtRegister & """ title=""" & strTxtRegister & """ /> <a href=""forum_terms.asp?M=reg&FID=" & intForumID & strQsSID2 & """>" & strTxtRegister & "</a>")
	    	Response.Write (" <img src=""" & strImagePath & "login.png"" alt=""" & strTxtLogin & """ title=""" & strTxtLogin & """ /> <a href=""login_user.asp?returnURL=" & strLinkPage & strQsSID2 & """>" & strTxtLogin & "</a>")
	End If
	
	
	'If the user is admin display a link to the admin menu
	If intGroupID = 1 Then Response.Write(" <img src=""" & strImagePath & "admin_control_panel.png"" title=""" & strTxtAdminControlPanel & """ alt=""" & strTxtAdminControlPanel & """ /> <a href=""admin.asp" & strQsSID1 & """>" & strTxtAdmin & "</a>")
	
	%>
  </td>
 </tr>
</table><%	



'Else display the desktop status bar
Else

%>
<iframe id="dropDownSearch" src="quick_search.asp?FID=<% = intForumID & "&SSL=" & blnSslEnabledPage & strQsSID2 %>" class="dropDownSearch" frameborder="0" scrolling="no"></iframe>
<table cellspacing="1" cellpadding="3" align="center" class="tableBorder">
 <tr class="tableStatusBar"> 
  <td>
   <div style="float:left;"><% = strBreadCrumbTrail %></div>
   <div style="float:right;"><%

	'If the user is admin display a link to the admin menu
	If intGroupID = 1 Then Response.Write("&nbsp;&nbsp;<img src=""" & strImagePath & "admin_control_panel." & strForumImageType & """ title=""" & strTxtAdminControlPanel & """ alt=""" & strTxtAdminControlPanel & """ style=""vertical-align: text-bottom"" />&nbsp;<a href=""admin.asp" & strQsSID1 & """>" & strTxtAdmin & "</a>")
		
	'Display any status bar graphics or menus for this page
	Response.Write(strStatusBarTools)

%></div>
  </td>
 </tr>
 <tr class="tableStatusBar">
  <td><%
     
	'If the user has logged in then the Logged In User ID number will not be 0 and not 2 for the guest account
	If lngLoggedInUserID <> 0 AND lngLoggedInUserID <> 2 Then
			
		'Link to user cp
		Response.Write(vbCrLf & "   <div style=""float:left;"">")
		If blnBanned = False Then Response.Write ("<img src=""" & strImagePath & "member_control_panel." & strForumImageType & """ title=""" & strTxtMemberCPMenu & """ alt=""" & strTxtMemberCPMenu & """ style=""vertical-align: text-bottom"" />&nbsp;<a href=""member_control_panel.asp" & strQsSID1 & """>" & strTxtMemberCPMenu & "</a>")
		%><!-- #include file="pm_check_inc.asp" --><%
		Response.Write("</div>")
	End If
%>
   <div style="float:right;"><%
   
	'Display the other common buttons
	Response.Write ("&nbsp;&nbsp;<img src=""" & strImagePath & "FAQ." & strForumImageType & """ alt=""" & strTxtFAQ & """ title=""" & strTxtFAQ & """ style=""vertical-align: text-bottom"" /> <a href=""help.asp" & strQsSID1 & """>" & strTxtFAQ & "</a>")
	Response.Write ("&nbsp;&nbsp;<span id=""SearchLink"" onclick=""showDropDown('SearchLink', 'dropDownSearch', 230, 0);"" class=""dropDownPointer""><img src=""" & strImagePath & "search." & strForumImageType & """ alt=""" & strTxtSearchTheForum & """ title=""" & strTxtSearchTheForum & """ style=""vertical-align: text-bottom"" /> <script language=""JavaScript"" type=""text/javascript"">document.write('" & strTxtSearch & "')</script><noscript><a href=""search_form.asp" & strQsSID1 & """>" & strTxtSearch & "</a></noscript></span>")
	If blnChatRoom AND lngLoggedInUserID <> 0 AND lngLoggedInUserID <> 2 AND blnBanned = False Then Response.Write ("&nbsp;&nbsp;<img src=""" & strImagePath & "chat_room." & strForumImageType & """ alt=""" & strTxtChat & """ title=""" & strTxtChat & """ style=""vertical-align: text-bottom"" /> <a href=""javascript:winOpener('chat.asp" & strQsSID1 & "','chat',1,1,700,580)"">" & strTxtChat & "</a>")
	If blnCalendar Then Response.Write ("&nbsp;&nbsp;<span id=""CalLink"" onclick=""getAjaxData('ajax_calendar.asp" & strQsSID1 & "', 'showCalendar');showDropDown('CalLink', 'dropDownCalendar', 210, 0);"" class=""dropDownPointer""><img src=""" & strImagePath & "calendar." & strForumImageType & """ alt=""" & strTxtEvents & """ title=""" & strTxtEvents & """ style=""vertical-align: text-bottom"" /> <script language=""JavaScript"" type=""text/javascript"">document.write('" & strTxtEvents & "')</script><noscript><a href=""calendar.asp" & strQsSID1 & """>" & strTxtEvents & "</a></noscript></span>")
	
	     
	'If the user has logged in then the Logged In User ID number will not be 0 and not 2 for the guest account
	If lngLoggedInUserID <> 0 AND lngLoggedInUserID <> 2 Then
	
		If blnDisplayMemberList Then Response.Write ("&nbsp;&nbsp;<img src=""" & strImagePath & "members." & strForumImageType & """ alt=""" & strTxtMembersList & """ title=""" & strTxtMembersList & """ style=""vertical-align: text-bottom"" /> <a href=""members.asp" & strQsSID1 & """>" & strTxtMemberlist & "</a>")
		'Don't display logout if windows authentication is enabled or member API
		If blnWindowsAuthentication = False AND (blnMemberAPIDisableAccountControl = False OR strMemberAPILogoutURL <> "") Then 
			Response.Write ("&nbsp;&nbsp;<img src=""" & strImagePath & "logout." & strForumImageType & """ alt=""" & strTxtLogOff & """ title=""" & strTxtLogOff & """ style=""vertical-align: text-bottom"" /> <a href=""log_off_user.asp?XID=" & getSessionItem("KEY") & strQsSID2 & """>" & strTxtLogOff & " [" & strLoggedInUsername & "]</a>")
		End If

	'Else if the member API is enabled and there are links to the websites on registration and login pages display the links to them
	ElseIf blnMemberAPI AND (strMemberAPIRegistrationURL <> "" OR strMemberAPILoginURL <> "") Then
		
		If strMemberAPIRegistrationURL <> "" Then Response.Write ("&nbsp;&nbsp;<img src=""" & strImagePath & "register." & strForumImageType & """ alt=""" & strTxtRegister & """ title=""" & strTxtRegister & """ style=""vertical-align: text-bottom"" /> <a href=""" & strMemberAPIRegistrationURL & """>" & strTxtRegister & "</a>")
		If strMemberAPILoginURL <> "" Then Response.Write ("&nbsp;&nbsp;<img src=""" & strImagePath & "login." & strForumImageType & """ alt=""" & strTxtLogin & """ title=""" & strTxtLogin & """ style=""vertical-align: text-bottom"" /> <a href=""" & strMemberAPILoginURL & """>" & strTxtLogin & "</a>")	
	
	'Else the user is not logged (Don't display logout if windows authentication is enabled or member API)
	ElseIf blnWindowsAuthentication = False AND (blnMemberAPI = False OR blnMemberAPIDisableAccountControl = False) Then
		
		Response.Write ("&nbsp;&nbsp;<img src=""" & strImagePath & "register." & strForumImageType & """ alt=""" & strTxtRegister & """ title=""" & strTxtRegister & """ style=""vertical-align: text-bottom"" /> <a href=""forum_terms.asp?M=reg&FID=" & intForumID & strQsSID2 & """>" & strTxtRegister & "</a>")
	    	Response.Write ("&nbsp;&nbsp;<img src=""" & strImagePath & "login." & strForumImageType & """ alt=""" & strTxtLogin & """ title=""" & strTxtLogin & """ style=""vertical-align: text-bottom"" /> <a href=""login_user.asp?returnURL=" & strLinkPage & strQsSID2 & """>" & strTxtLogin & "</a>")
	End If

%></div>
  </td>
 </tr>
</table><%

	'Display hidden div for calandar
	'(AJAX used for this one so that the extra database hit is only performed if the user wants to view the calendar)
	If blnCalendar Then

%>
<div id="dropDownCalendar" class="dropDownCalendar"><span id="showCalendar"></span></div><%
	
	End If
End If

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If NOT strForumHeaderAd = "" Then 
	If blnMobileBrowser = False Then 
		Response.Write("<div align=""center"" style=""margin:5px;""><br />" & strForumHeaderAd & "</div>")
	Else
		Response.Write("<div align=""center"" style=""margin:5px;max-width:620px;overflow:auto;overflow-x:hidden;"">" & strForumHeaderAd & "</div>")
	End If
End If
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

%>
<br /><%

'If a message for all forums display it
If strForumsMessage <> "" Then
	
%>
<table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
 <tr class="tableRow"> 
  <td><% = strForumsMessage %></td>
 </tr>
</table>
<br /><%

End If
%>