<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="language_files/pm_language_file_inc.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
<!--#include file="functions/functions_send_mail.asp" -->
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

'Declare variables
Dim lngPmMessageID		'Private message id
Dim strPmSubject 		'Holds the subject of the private message
Dim strMessage			'Holds the message body of the thread
Dim lngMessageID		'Holds the message ID number
Dim lngFromUserID		'Holds the from user ID
Dim lngToUserID			'Holds the to user ID
Dim dtmTopicDate		'Holds the date the thread was made
Dim strUsername 		'Holds the Username of the thread
Dim strAuthorHomepage		'Holds the homepage of the Username if it is given
Dim strAuthorLocation		'Holds the location of the user if given
Dim strAuthorAvatar		'Holds the authors avatar	
Dim lngAuthorNumOfPosts		'Holds the number of posts the user has made to the forum
Dim dtmAuthorRegistration	'Holds the registration date of the user
Dim intRecordLoopCounter	'Holds the loop counter numeber
Dim intTopicPageLoopCounter	'Holds the number of pages there are of pm messages
Dim strEmailBody		'Holds the body of the e-mail message
Dim strEmailSubject		'Holds the subject of the e-mail
Dim blnEmailSent		'set to true if an e-mail is sent
Dim strGroupName		'Holds the authors group name
Dim intRankStars		'Holds the number of stars for the group
Dim strMemberTitle		'Holds the members title
Dim strRankCustomStars		'Holds custom stars for the user group
Dim strXID			'Holds the xid


blnSslEnabledPage = True



'Read in the pm mesage number to display
If isNumeric(Request.QueryString("ID")) Then
	lngPmMessageID = LngC(Request.QueryString("ID"))
Else
	'Clean up
	Call closeDatabase()
	
	Response.Redirect("pm_inbox.asp" & strQsSID1)
End If



'If Priavte messages are not on then send them away
If blnPrivateMessages = False Then 
	
	'Clean up
	Call closeDatabase()
	
	Response.Redirect("default.asp" & strQsSID1)
End If



'If the user is not allowed then send them away
If intGroupID = 2 OR blnActiveMember = False OR blnBanned Then 
	
	'Clean up
	Call closeDatabase()
	
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If


	
'Initlise the sql statement
strSQL = "SELECT " & strDbTable & "PMMessage.*, " & strDbTable & "Author.Username, " & strDbTable & "Author.Author_email, " & strDbTable & "Author.Join_date, " & strDbTable & "Author.Signature, " & strDbTable & "Author.Active, " & strDbTable & "Author.Avatar, " & strDbTable & "Group.Name, " & strDbTable & "Group.Stars, " & strDbTable & "Group.Custom_stars " & _
"FROM " & strDbTable & "Author" & strDBNoLock & ", " & strDbTable & "PMMessage" & strDBNoLock & ", " & strDbTable & "Group" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Author.Author_ID = " & strDbTable & "PMMessage.From_ID " & _
	"AND " & strDbTable & "Author.Group_ID = " & strDbTable & "Group.Group_ID " & _
	"AND " & strDbTable & "PMMessage.PM_ID = " & lngPmMessageID & " "
	
'If this is a link from the out box then check the from author ID to check the user can view the message
If Request.QueryString("M") = "OB" Then
	strSQL = strSQL & " AND " & strDbTable & "PMMessage.From_ID = " & lngLoggedInUserID & ";"
'Else use the to author ID to check the user can view the message
Else
	strSQL = strSQL & " AND " & strDbTable & "PMMessage.Author_ID = " & lngLoggedInUserID & ";"
End If



'Query the database
rsCommon.Open strSQL, adoCon


'If a mesage is found then send a mail if the sender wants notifying
If NOT rsCommon.EOF Then 
	
	'Read in some of the details
	strPmSubject = rsCommon("PM_Tittle")
	strUsername = rsCommon("Username")
	strMessage = rsCommon("PM_Message")
	lngFromUserID = CLng(rsCommon("From_ID"))
	lngToUserID = CLng(rsCommon("Author_ID"))
	dtmTopicDate = CDate(rsCommon("PM_Message_date")) 
	strAuthorAvatar = rsCommon("Avatar")
	strGroupName = rsCommon("Name")
	intRankStars = CInt(rsCommon("Stars"))
	strRankCustomStars = rsCommon("Custom_stars")
	
	'If the sender wants notifying then send a mail as long as e-mail notify is on and the message hasn't already been read
	If CBool(rsCommon("Email_notify")) AND rsCommon("Author_email") <> "" AND blnEmail AND CBool(rsCommon("Read_Post")) = False AND Request.QueryString("M") <> "OB" Then
		
		'Set the subject
		strEmailSubject = strMainForumName & " " & strTxtNotificationPM
	
		'Initailise the e-mail body variable with the body of the e-mail
		strEmailBody = strTxtHi & " " & decodeString(strUsername) & "," & _
		vbCrLf & vbCrLf & strTxtThisIsToNotifyYouThat & " '" & strLoggedInUsername & "' " & strTxtHasReadPM & ", '" & decodeString(strPmSubject) & "', " & strTxtYouSentToThemOn & " " & strMainForumName & "." & _
		vbCrLf & vbCrLf & strTxtToViewThePrivateMessage & " " & strTxtForumClickOnTheLinkBelow & ": -" & _
                vbCrLf & vbCrLf & strForumPath & "pm_message.asp?ID=" & lngPmMessageID & "&M=OB"
		
		'Call the function to send the e-mail
		blnEmailSent = SendMail(strEmailBody, decodeString(strUsername), decodeString(rsCommon("Author_email")), strWebsiteName, decodeString(strForumEmailAddress), strEmailSubject, strMailComponent, false)
	
	End If
	
	'Filter for CSS hacks
	strPmSubject = formatInput(strPmSubject)
	
	'If the pm contains a quote or code block then format it
	If InStr(1, strMessage, "[QUOTE=", 1) > 0 AND InStr(1, strMessage, "[/QUOTE]", 1) > 0 Then strMessage = formatUserQuote(strMessage)
	If InStr(1, strMessage, "[QUOTE]", 1) > 0 AND InStr(1, strMessage, "[/QUOTE]", 1) > 0 Then strMessage = formatQuote(strMessage)
	If InStr(1, strMessage, "[CODE]", 1) > 0 AND InStr(1, strMessage, "[/CODE]", 1) > 0 Then strMessage = formatCode(strMessage)
	If InStr(1, strMessage, "[SPOILER]", 1) > 0 AND InStr(1, strMessage, "[/SPOILER]", 1) > 0 Then strMessage = formatSpoiler(strMessage)

	
	'If the Post or signature contains Flash BBcodes then display them
	If blnPmFlashFiles Then
		If InStr(1, strMessage, "[FLASH", 1) > 0 AND InStr(1, strMessage, "[/FLASH]", 1) > 0 Then strMessage = formatFlash(strMessage)
	End If
		
	'If YouTube
	If blnPmYouTube Then
		If InStr(1, strMessage, "[TUBE]", 1) > 0 AND InStr(1, strMessage, "[/TUBE]", 1) > 0 Then strMessage = formatYouTube(strMessage)
	End If
End If



'Close recordset
rsCommon.Close



'If this is not from the outbox then update the read field
If Request.QueryString("M") <> "OB" Then
	
	'Inittilise the sql veriable to update the database
	strSQL = "UPDATE " & strDbTable & "PMMessage " & strRowLock & " " & _
	"SET " & strDbTable & "PMMessage.Read_Post = " & strDBTrue & " " & _
	"WHERE " & strDbTable & "PMMessage.PM_ID = " & lngPmMessageID & ";"

	'Execute the sql statement to set the pm to read
	adoCon.Execute(strSQL)
	
	
	'Update the number of unread PM's 
	Call updateUnreadPM(lngLoggedInUserID)
	
	
	'Update the notified PM session variable
	If intNoOfPms = 0 Then
		Call saveSessionItem("PMN", "")
	Else
		Call saveSessionItem("PMN", intNoOfPms)
	End If
End If


'Get session key
strXID = getSessionItem("KEY")

	

'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers(strTxtViewingPrivateMessage, "", "", 0)
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""pm_welcome.asp" & strQsSID1 & """>" & strTxtPrivateMessenger & "</a>" & strNavSpacer & strPmSubject


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtPrivateMessenger & " - " & strPmSubject %></title>

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
  <td align="left"><h1><% = strTxtPrivateMessenger & " - " & strPmSubject %></h1></td>
</tr>
</table>
<br />
<table class="basicTable" cellspacing="0" cellpadding="0" align="center"> 
 <tr> 
  <td class="tabTable">
   <a href="pm_welcome.asp<% = strQsSID1 %>" title="<% = strTxtPrivateMessenger %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>messenger.<% = strForumImageType %>" border="0" alt="<% = strTxtPrivateMessenger %>" /> <% = strTxtMessenger %></a>
   <a href="pm_inbox.asp<% = strQsSID1 %>" title="<% = strTxtPrivateMessenger & " " & strTxtInbox %>" class="tabButtonActive">&nbsp;<img src="<% = strImagePath %>inbox_messages.<% = strForumImageType %>" border="0" alt="<% = strTxtPrivateMessenger & " " & strTxtInbox %>" /> <% = strTxtInbox %></a>
   <a href="pm_outbox.asp<% = strQsSID1 %>" title="<% = strTxtPrivateMessenger & " " & strTxtOutbox %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>sent_messages.<% = strForumImageType %>" border="0" alt="<% = strTxtPrivateMessenger & " " & strTxtOutbox %>" /> <% = strTxtOutbox %></a>
   <a href="pm_new_message_form.asp<% = strQsSID1 %>" title="<% = strTxtNewPrivateMessage %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>new_message.<% = strForumImageType %>" border="0" alt="<% = strTxtNewPrivateMessage %>" /> <% = strTxtNewMessage %></a>
  </td>
 </tr>
</table>
<br /><%

'If no private message display an error
IF strPmSubject = "" Then
%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><% = strTxtYouDoNotHavePermissionViewPM %></td>
 </tr>
</table><%

'Else display the message
Else

%>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center" style="table-layout: fixed;">
 <tr class="tableLedger">
  <td><% = strTxtPrivateMessage %></td>
 </tr>
 <tr class="tableRow">
  <td><%
	'If the user has an avatar then display it
	If blnAvatar = True Then 
		
		'If the avatar is not done then show the blank avatar
		If strAuthorAvatar = "" Then strAuthorAvatar = "avatars/blank_avatar.jpg"
			
		'If SSL is enabled, but the avatar is not HTTPS then remove avatar
		If strSslEnabled = "Enabled" AND InStr(strAuthorAvatar, "http://") Then strAuthorAvatar = "avatars/blank_avatar.jpg"
		
		%><img src="<% = strAuthorAvatar %>" id="avatar" alt="<% = strTxtAvatar %>" align="left" onError="document.getElementById('avatar').src='avatars/blank_avatar.jpg'" /><%
	  	
	End If
%>
      <% = strTxtSubject & " - " &  strPmSubject %>
      <br />
      <% = strTxtSent & " - " & DateFormat(dtmTopicDate) & " " & strTxtAt & " " & TimeFormat(dtmTopicDate) %>
      <br />
      <% = strTxtSentBy %>: <a href="member_profile.asp?PF=<% =lngFromUserID & strQsSID2 %>" rel="nofollow"><% = strUsername %></a>
      <br /><% 
      
      	'Display author details
      	Response.Write(strTxtGroup & " - " & strGroupName & " <img src=""")
	If strRankCustomStars <> "" Then Response.Write(strRankCustomStars) Else Response.Write(strImagePath & intRankStars & "_star_rating.png")
	Response.Write(""" alt=""" & strGroupName & """> ")
      	
      	
      	
      	%>
  </td>
 </tr>
  <tr class="tableTopRow">
  <td>
   <div style="float:left;"><strong><% = strPmSubject %></strong></div>
   <div style="float:right;">
    <img src="<% = strImagePath %>add_buddy.png" border="0" alt="<% = strTxtAddToBuddyList %>" alt="<% = strTxtAddToBuddyList %>" /> <a href="pm_buddy_list.asp?name=<% = Server.URLEncode(strUsername) & strQsSID2 %>" class="PMsmLink"><% = strTxtAddBuddy %></a><%
  
	'If the person reading the pm is the recepient disply delete and reply buttons
	If lngToUserID = lngLoggedInUserID Then
		%>
    &nbsp;<img src="<% = strImagePath %>delete_message.png" border="0" alt="<% = strTxtDelete %>" title="<% = strTxtDelete %>" /> <a href="pm_delete_message.asp?pm_id=<% = lngPmMessageID %>&XID=<% = strXID & strQsSID2 %>" OnClick="return confirm('<% = strTxtDeletePrivateMessageAlert %>')" class="PMsmLink"><% = strTxtDelete %></a> 
    &nbsp;<img src="<% = strImagePath %>message_reply.png" border="0" alt="<% = strTxtReplyToPrivateMessage %>" title="<% = strTxtReplyToPrivateMessage %>" /> <a href="pm_new_message_form.asp?code=reply&pm=<% = lngPmMessageID & strQsSID2 %>" class="PMsmLink"><% = strTxtReply %></a><%
	
	End If

%>
   </div>
  </td>
 </tr>
 <tr class="PMtableRow">
  <td>
   <!-- Start Private Message -->
    <div class="PMmsgBody">
     <% = strMessage %>
    </div>
   <!-- End Private Message -->
  </td>
 </tr>
 <tr class="tableBottomRow">
  <td>
   <div style="float:left;"><% 
   
   	'If the user has an email address and emailing is enabled then allow user to receive this pm by email
	If blnLoggedInUserEmail AND blnEmail Then
%>
    <a href="pm_email_pm.asp?ID=<% = lngPmMessageID %><% If Request.QueryString("M") = "OB" Then Response.Write("&M=OB")%><% = strQsSID2 %>" class="PMsmLink"><% = strTxtEmailThisPMToMe %></a><%
    	
	End If
%>		
   </div>		
   <div style="float:right;">
    <img src="<% = strImagePath %>add_buddy.png" border="0" alt="<% = strTxtAddToBuddyList %>" alt="<% = strTxtAddToBuddyList %>" /> <a href="pm_buddy_list.asp?name=<% = Server.URLEncode(strUsername) & strQsSID2 %>" class="PMsmLink"><% = strTxtAddBuddy %></a><%
  
	'If the person reading the pm is the recepient disply delete and reply buttons
	If lngToUserID = lngLoggedInUserID Then
		%>
    &nbsp;<img src="<% = strImagePath %>delete_message.png" border="0" alt="<% = strTxtDelete %>" title="<% = strTxtDelete %>" /> <a href="pm_delete_message.asp?pm_id=<% = lngPmMessageID %>&XID=<% = strXID & strQsSID2 %>" OnClick="return confirm('<% = strTxtDeletePrivateMessageAlert %>')" class="PMsmLink"><% = strTxtDelete %></a> 
    &nbsp;<img src="<% = strImagePath %>message_reply.png" border="0" alt="<% = strTxtReplyToPrivateMessage %>" title="<% = strTxtReplyToPrivateMessage %>" /> <a href="pm_new_message_form.asp?code=reply&pm=<% = lngPmMessageID & strQsSID2 %>" class="PMsmLink"><% = strTxtReply %></a><%
	
	End If

%>
   </div>
  </td>
 </tr>
</table><%

End If

%>
<br />
<div align="center"><%
	

'If a mobile browser display an option to switch to and from mobile view
If blnMobileBrowser Then 
	Response.Write (strTxtViewIn & ": <strong>" & strTxtMoble & "</strong> | <a href=""pm_message.asp?ID=" & LngC(lngPmMessageID) & "&MobileView=off" & strQsSID2 & """ rel=""nofollow"">" & strTxtClassic & "</a><br /><br />")
ElseIf blnMobileClassicView Then
	Response.Write (strTxtViewIn & ": <a href=""pm_message.asp?ID=" & LngC(lngPmMessageID) & "&MobileView=on" & strQsSID2 & """ rel=""nofollow"">" & strTxtMoble & "</a> | <strong>" & strTxtClassic & "</strong><br /><br />")
End If


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
<%


If strVigLinkKey = "" Then 
	Response.Write(" <script type=""text/javascript"" src=""https://syndication.webwiz.net/exped/?SKU=WWF10""></script>" & vbCrLf )


ElseIf NOT strVigLinkKey = "Disabled" Then 
	
	Response.Write(vbCrLf & "<script type=""text/javascript"">" & _
		vbCrLf & "    var vglnk = { key: '" & strVigLinkKey & "' };" & _
		vbCrLf & "  " & _
		vbCrLf & "    (function(d, t) {" & _
		vbCrLf & "      var s = d.createElement(t); s.type = 'text/javascript'; s.async = true;" & _
		vbCrLf & "      s.src = '//cdn.viglink.com/api/vglnk.js';" & _
		vbCrLf & "      var r = d.getElementsByTagName(t)[0]; r.parentNode.insertBefore(s, r);" & _
		vbCrLf & "    }(document, 'script'));" & _
		vbCrLf & "</script>" & _
		vbCrLf & "")
End If


'Display a msg letting the user know they have been emailed a private message
If Request.QueryString("ES") = "True" Then
	Response.Write("<script  language=""JavaScript"">")
	Response.Write("alert('" & strTxtAnEmailWithPM & " " & strTxtBeenSent & ".');")
	Response.Write("</script>")
ElseIf Request.QueryString("ES") = "False" Then
	Response.Write("<script  language=""JavaScript"">")
	Response.Write("alert('" & strTxtAnEmailWithPM & " " & strTxtNotBeenSent & ".');")
	Response.Write("</script>")
End If

'Clear server objects
Call closeDatabase()
%>
<!-- #include file="includes/footer.asp" -->