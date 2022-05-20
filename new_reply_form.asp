<% @ Language=VBScript %>
<% Option Explicit %>
<% Response.AddHeader "X-UA-Compatible","IE=EmulateIE10" %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
<!--#include file="functions/functions_edit_post.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
<!--#include file="includes/emoticons_inc.asp" -->
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

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"


'Dimension variables
Dim strMode			'Holds the mode of the page
Dim strMasterForumName		'Holds the main forum name
Dim intMasterForumID		'Holds the main forum ID
Dim intCatID			'Holds the cat ID
Dim strCatName			'Holds the cat name
Dim lngTopicID			'Holds the Topic ID number
Dim lngMessageID		'Holds the message ID to be edited
Dim strTopicSubject		'Holds the subject topic being replied to
Dim strReplyUsername		'Holds the repliers username or quoters username
Dim strReplyMessage		'Holds the message to be replied to or quoted
Dim lngPostUserID		'Holds the user ID of the user to post the message
Dim blnForumLocked		'Set to true if the forum is locked
Dim blnEmailNotify		'Set to true if the users want to be notified by e-mail of a post
Dim strPostPage 		'Holds the page the form is posted to
Dim intRecordPositionPageNum	'Holds the recorset page number to show the Threads for
Dim strMessage			'Holds the post message
Dim strForumName		'Holds the name of the forum
Dim intIndexPosition		'Holds the idex poistion in the emiticon array
Dim intNumberOfOuterLoops	'Holds the outer loop number for rows
Dim intLoop			'Holds the loop index position
Dim intInnerLoop		'Holds the inner loop number for columns
Dim blnTopicLocked		'Holds if the topic is locked or not
Dim lngTotalRecords		'Holds the total number of records
Dim strUsername			'For login include
Dim strPassword			'For login include
Dim blnPollNoReply		'Holds if it is a poll only
Dim lngPollID			'Holds the poll ID
Dim strUploadedFiles		'Holds the names of any files or images uploaded
Dim strTopicIcon		'Holds the topic icon
Dim intEventYear		'Holds the year of Calendar event
Dim intEventMonth		'Holds the month of Calendar event
Dim intEventDay			'Holds the day of Calendar event
Dim dtmEventDate		'Holds the date if this is a calendar event
Dim dtmPostDate			'Holds the date the thread was made
Dim intEventYearEnd		'Holds the year of Calendar event
Dim intEventMonthEnd		'Holds the month of Calendar event
Dim intEventDayEnd		'Holds the day of Calendar event
Dim dtmEventDateEnd		'Holds the Calendar event date
Dim strFormID			'Holds the ID for the form
Dim strQuote


'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)
End If




'Read in the message ID number to edit
lngTopicID = LngC(Request.QueryString("TID"))
lngMessageID = LngC(Request.QueryString("PID"))
If isNumeric(Request.QueryString("PN")) Then intRecordPositionPageNum = IntC(Request.QueryString("PN")) Else intRecordPositionPageNum = 1
strQuote = Trim(Mid(Request.QueryString("Quote"), 1, 2))
If isNumeric(Request.QueryString("TR")) Then lngTotalRecords = LngC(Request.QueryString("TR")) Else lngTotalRecords = 1
intForumID = 0

'Set the page mode
If strQuote = "1" Then 
	strMode = "quote"
Else
	strMode = "reply"
End If

	
'Get the message from the database to reply/quote
If lngTopicID > 0 Then	
	
	'Initalise the strSQL variable with an SQL statement to get the message to be quoted
	strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Forum_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Locked, " & strDbTable & "Topic.Poll_ID, " & strDbTable & "Topic.Event_date, " & strDbTable & "Thread.Author_ID, " & strDbTable & "Thread.Message, " & strDbTable & "Thread.Message_date, " & strDbTable & "Author.Username, " & strDbTable & "GuestName.Name " & _
	"FROM (" & strDbTable & "Author" & strDBNoLock & " INNER JOIN (" & strDbTable & "Topic" & strDBNoLock & " INNER JOIN " & strDbTable & "Thread" & strDBNoLock & " ON " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID) ON " & strDbTable & "Author.Author_ID = " & strDbTable & "Thread.Author_ID) LEFT JOIN " & strDbTable & "GuestName" & strDBNoLock & " ON " & strDbTable & "Thread.Thread_ID = " & strDbTable & "GuestName.Thread_ID "  & _
	"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & " " & _
	"AND " & strDbTable & "Topic.Hide = " & strDBFalse & " " & _
	"AND (" & strDbTable & "Thread.Hide = " & strDBFalse & " OR " & strDbTable & "Thread.Author_ID = " & lngLoggedInUserID & ") " & _
	"ORDER BY " & strDbTable & "Thread.Message_date DESC " & strDBLimit1 & ";"
	

'If large reply button clicked get the last thread in topic
Else
	'Initalise the strSQL variable with an SQL statement to get the message to be quoted
	strSQL = "SELECT " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Forum_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Locked, " & strDbTable & "Topic.Poll_ID, " & strDbTable & "Topic.Event_date, " & strDbTable & "Thread.Author_ID, " & strDbTable & "Thread.Message, " & strDbTable & "Thread.Message_date, " & strDbTable & "Author.Username, " & strDbTable & "GuestName.Name " & _
	"FROM (" & strDbTable & "Author" & strDBNoLock & " INNER JOIN (" & strDbTable & "Topic" & strDBNoLock & " INNER JOIN " & strDbTable & "Thread" & strDBNoLock & " ON " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID) ON " & strDbTable & "Author.Author_ID = " & strDbTable & "Thread.Author_ID) LEFT JOIN " & strDbTable & "GuestName" & strDBNoLock & " ON " & strDbTable & "Thread.Thread_ID = " & strDbTable & "GuestName.Thread_ID "  & _
	"WHERE " & strDbTable & "Thread.Thread_ID = " & lngMessageID & " " &_
	"AND " & strDbTable & "Topic.Hide = " & strDBFalse & " " & _
	"AND (" & strDbTable & "Thread.Hide = " & strDBFalse & " OR " & strDbTable & "Thread.Author_ID = " & lngLoggedInUserID & ");"
End If

'Query the database
rsCommon.Open strSQL, adoCon 
	

'Read in the details from the recordset
If NOT rsCommon.EOF Then
	intForumID = CInt(rsCommon("Forum_ID"))
	strTopicSubject = rsCommon("Subject")
	blnTopicLocked = CBool(rsCommon("Locked"))
	lngTopicID = CLng(rsCommon("Topic_ID"))
	strTopicSubject = rsCommon("Subject")
	lngPostUserID = CLng(rsCommon("Author_ID"))
	strReplyUsername = rsCommon("Username")
	strReplyMessage = rsCommon("Message")
	lngPollID = CLng(rsCommon("Poll_ID"))
	If isDate(rsCommon("Event_date")) Then dtmEventDate = CDate(rsCommon("Event_date"))
	dtmPostDate = CDate(rsCommon("Message_date"))
	
	'If the post being quoted is written by a guest see if they have a name
	If lngPostUserID = 2 Then strReplyUsername = rsCommon("Name")
End If

'Clean up
rsCommon.Close


'If this is a quote setup the quote block	
If strMode = "quote" Then
		
	'Build up the quoted thread post
	strMessage = " [QUOTE=" & strReplyUsername & "] " 
	
	'Read in the quoted thread from the recordset
	strMessage = strMessage & strReplyMessage
	
	'Apply forum codes 
	strMessage = EditPostConvertion (strMessage)
	
	'Place the forum code for closing quote at the end	
	strMessage = strMessage & "[/QUOTE] "
End If	
	


'Clean up input to prevent XXS hack
strTopicSubject = formatInput(strTopicSubject)




'Check the forum permissions

'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "" & _
"SELECT" & strDBTop1 & " " & strDbTable & "Category.Cat_ID, " & strDbTable & "Category.Cat_name, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum2.Forum_name AS Main_forum, " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Forum_code, " & strDbTable & "Forum.Locked, " & strDbTable & "Forum.Show_topics, " & strDbTable & "Permissions.* " & _
"FROM " & strDbTable & "Category" & strDBNoLock & ", " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Forum AS " & strDbTable & "Forum2" & strDBNoLock & ", " & strDbTable & "Permissions" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID " & _
 	"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Permissions.Forum_ID " & _
 	"AND (" & strDbTable & "Forum.Sub_ID = " & strDbTable & "Forum2.Forum_ID OR (" & strDbTable & "Forum.Sub_ID = 0 AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Forum2.Forum_ID)) " & _
 	"AND " & strDbTable & "Forum.Forum_ID = " & intForumID & " " & _
 	"AND (" & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intGroupID & ") " & _
"ORDER BY " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_Order, " & strDbTable & "Permissions.Author_ID DESC" & strDBLimit1 & ";"
	

'Query the database
rsCommon.Open strSQL, adoCon

'If there is a record returned by the recordset then check to see if you need a password to enter it
If NOT rsCommon.EOF Then

	'Read in forum details from the database
	intCatID = CInt(rsCommon("Cat_ID"))
	strCatName = rsCommon("Cat_name")
	strForumName = rsCommon("Forum_name")
	strMasterForumName = rsCommon("Main_forum")
	intMasterForumID = CLng(rsCommon("Sub_ID"))
	blnForumLocked = CBool(rsCommon("Locked"))
	
	'Read in the forum permissions
	blnRead = CBool(rsCommon("View_Forum"))
	blnReply = CBool(rsCommon("Reply_posts"))
	blnModerator = CBool(rsCommon("Moderate"))
	blnCheckFirst = CBool(rsCommon("Display_post"))
	blnEvents = CBool(rsCommon("Calendar_event"))

	'If the user has no read writes then kick them
	If blnRead = False Then
		
		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()

		'Redirect to a page asking for the user to enter the forum password
		Response.Redirect("insufficient_permission.asp" & strQsSID1)
	End If


	'If the forum requires a password and a logged in forum code is not found on the users machine then send them to a login page
	If rsCommon("Password") <> "" AND (getCookie("fID", "Forum" & intForumID) <> rsCommon("Forum_code") AND getSessionItem("FP" & intForumID) <> "1") Then

		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()
		
		'Redirect to a page asking for the user to enter the forum password
		Response.Redirect("forum_password_form.asp?FID=" & intForumID & strQsSID3)
	End If
End If

'Clean up
rsCommon.Close


'Check that the poll is not a poll only and replies are allowed
If lngPollID <> 0 Then
	
	strSQL = "SELECT " & strDbTable & "Poll.Reply " & _
	"FROM " & strDbTable & "Poll" & strDBNoLock & " "  & _
	"WHERE " & strDbTable & "Poll.Poll_ID=" & lngPollID & ";"
		
	'Query the database
	rsCommon.Open strSQL, adoCon 
	
	'Read in the details from the recordset
	blnPollNoReply = CBool(rsCommon("Reply"))
	
	'Clean up
	rsCommon.Close
End If


'Check to see if the user has email notification for this topic
If blnEmail AND blnLoggedInUserEmail Then
	strSQL = "SELECT " & strDbTable & "EmailNotify.Author_ID  " & _
	"FROM " & strDbTable & "EmailNotify" & strRowLock & " " & _
	"WHERE " & strDbTable & "EmailNotify.Author_ID = " & lngLoggedInUserID & " AND " & strDbTable & "EmailNotify.Topic_ID = " & lngTopicID & ";"
	
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'If a record is returned then users has email notification enabled
	If NOT rsCommon.EOF Then blnReplyNotify = True
	
	'Close RS	
	rsCommon.Close
End If


'If this is a submission from the quick reply form to the full reply form run extra code
If Request.Form("QR") = "1" Then
	
	'The mode is Quick Reply to Full Reply
	strMode = "QuickToFull"
	
	'Read in the contents of the quick reply form
	strMessage = Request.Form("message")

	'Run some filters to make sure the content is OK, and no dodgy HTML is trying to sneek through
	strMessage = HTMLsafe(strMessage)
End If


'If the forum level for the user on this forum is read only set the forum to be locked
If (blnRead = False AND blnModerator = False AND blnAdmin = False) Then blnForumLocked = True


'If a login is displayed then switch to HTTPS
If blnReply = False AND intGroupID = 2  AND blnActiveMember AND blnForumLocked = false AND blnTopicLocked = false AND blnBanned = False Then blnSslEnabledPage = True 



'If the message has been edited remove who edited the post
If InStr(1, strMessage, "<edited>", 1) Then strMessage = removeEditorAuthor(strMessage)

	
	
'Setup reply bottom of page
If InStr(1, strReplyMessage, "[QUOTE=", 1) > 0 AND InStr(1, strReplyMessage, "[/QUOTE]", 1) > 0 Then strReplyMessage = formatUserQuote(strReplyMessage)
If InStr(1, strReplyMessage, "[QUOTE]", 1) > 0 AND InStr(1, strReplyMessage, "[/QUOTE]", 1) > 0 Then strReplyMessage = formatQuote(strReplyMessage)
If InStr(1, strReplyMessage, "[CODE]", 1) > 0 AND InStr(1, strReplyMessage, "[/CODE]", 1) > 0 Then strReplyMessage = formatCode(strReplyMessage)
If InStr(1, strReplyMessage, "[HIDE]", 1) > 0 AND InStr(1, strReplyMessage, "[/HIDE]", 1) > 0 Then strReplyMessage = formatHide(strReplyMessage)
If InStr(1, strReplyMessage, "[SPOILER]", 1) > 0 AND InStr(1, strReplyMessage, "[/SPOILER]", 1) > 0 Then strReplyMessage = formatSpoiler(strReplyMessage, False)
	

'If the post contains a flash link then format it
If blnFlashFiles Then
	If InStr(1, strReplyMessage, "[FLASH", 1) > 0 AND InStr(1, strReplyMessage, "[/FLASH]", 1) > 0 Then strReplyMessage = formatFlash(strReplyMessage)
End If

'If YouTube
If blnYouTube Then
	If InStr(1, strReplyMessage, "[TUBE]", 1) > 0 AND InStr(1, strReplyMessage, "[/TUBE]", 1) > 0 Then strReplyMessage = formatYouTube(strReplyMessage)
End If

'If the message has been edited parse the 'edited by' XML into HTML for the post
If InStr(1, strReplyMessage, "<edited>", 1) Then strReplyMessage = editedXMLParser(strReplyMessage)





'Use the application session to pass around what forum this user is within
If IntC(getSessionItem("FID")) <> intForumID Then Call saveSessionItem("FID", intForumID)
	
	
'get the session key
strFormID = getSessionItem("KEY")



'Set bread crumb trail
''Display the category name
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""default.asp?C=" & intCatID & strQsSID2 & SeoUrlTitle(strCatName, "&title=") & """>" & strCatName & "</a>" & strNavSpacer

'Display if there is a main forum to the sub forums name
If intMasterForumID <> 0 Then strBreadCrumbTrail = strBreadCrumbTrail & "<a href=""forum_topics.asp?FID=" & intMasterForumID & strQsSID2 & SeoUrlTitle(strMasterForumName, "&title=") & """>" & strMasterForumName & "</a>" & strNavSpacer

'Display forum name
If strForumName = "" Then strBreadCrumbTrail = strBreadCrumbTrail &  strTxtNoForums Else strBreadCrumbTrail = strBreadCrumbTrail & "<a href=""forum_topics.asp?FID=" & intForumID & strQsSID2 & SeoUrlTitle(strForumName, "&title=")  & """>" & strForumName & "</a>"

strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtPostReply



%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strTxtPostReply %></title>
<meta name="generator" content="Web Wiz Forums" />
<meta name="robots" content="noindex, follow" />

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

<!-- Check the from is filled in correctly before submitting -->
<script  language="JavaScript">

//Function to check form is filled in correctly before submitting
function CheckForm () {
	
	var errorMsg = "";
	var formArea = document.getElementById('frmMessageForm');
<%
'If Gecko Madis API (RTE) need to strip default input from the API
If RTEenabled = "Gecko" Then Response.Write("	//For Gecko Madis API (RTE)" & vbCrLf & "	if (formArea.message.value.indexOf('<br>') > -1 && formArea.message.value.length==6) formArea.message.value = '';" & vbCrLf)

'If this is a guest posting check that they have entered their name
If lngLoggedInUserID = 2 Then
%>	
	//Check for a name
	if (formArea.Gname.value==""){
		errorMsg += "\n<% = strTxtNoNameError %>";
	}<%

End If

'If CAPTCHA is displayed check it's been entered
If blnCAPTCHAsecurityImages AND lngLoggedInUserID = 2 Then 
	
	%>
	
	//Check for a security code
        if (formArea.securityCode.value == ''){
                errorMsg += "\n<% = strTxtErrorSecurityCode %>";
        }<%

End If

%>	

	//Check for message
	if (formArea.message.value==""){
		errorMsg += "\n<% = strTxtNoMessageError %>";
	}
	
	//Check session is not expired
        if (formArea.formID.value == ''){
                errorMsg += "\n<% = strTxtWarningYourSessionHasExpiredRefreshPageFormDataWillBeLost %>";
        }
	
	//If there is aproblem with the form then display an error
	if (errorMsg != ""){
		msg = "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine1 %>\n";
		msg += "<% = strTxtErrorDisplayLine2 %>\n";
		msg += "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine3 %>\n";
		
		errorMsg += alert(msg + errorMsg + "\n\n");
		return false;
	}
	return true;
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td><h1><%  = strTxtPostReply & " - " & strTopicSubject %></h1></td>
 </tr>
</table>
<br /><%

 
'If the Post is by the logged in user or the adminstrator/moderator then display a form to reply
If (blnReply OR blnAdmin) AND blnActiveMember AND blnPollNoReply = false AND (blnForumLocked = False OR blnAdmin = True) AND (blnTopicLocked = False Or blnAdmin) Then
	
	'Update active users table
	If blnActiveUsers Then saryActiveUsers = activeUsers(strTxtWritingReply, strTopicSubject, "forum_posts.asp?TID=" & lngTopicID, intForumID)

	%>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td align="left"><% = strTxtPostReply %></td>
 </tr>
 <tr class="tableRow">
  <td align="left">
  <!--#include file="includes/message_form_inc.asp" -->
  </td>
 </tr>
</table>
<br />
<table class="tableBorder" align="center" cellspacing="1" cellpadding="10" style="table-layout: fixed;">
 <tr class="tableLedger">
  <td><% = strTxtMessage %></td>
 </tr>
 <tr class="msgEvenTableTop">
  <td valign="top"><%

     		'If display topic title
     		Response.Write("<strong>")
		'If a calendar event then display so
		If isDate(dtmEventDate) Then
			Response.Write(strTxtCalendarEvent & ": " & strTopicSubject & " - " & strTxtEventDate & ": " & DateFormat(dtmEventDate))
		Else  	
		 	Response.Write(strTxtTopic & " - " & strTopicSubject) 
		End If
		Response.Write("</strong><br />")

		'Display message post date and time
		Response.Write(strTxtPosted & " " & DateFormat(dtmPostDate) & " " & strTxtAt & " " & TimeFormat(dtmPostDate) & " " & strTxtBy & " " & strReplyUsername) 
%>
  </td>
 </tr>
 <tr class="msgEvenTableRow" style="height:150px;min-height:150px;">
  <td valign="top" class="msgLineDevider">
   <!-- Start Member Post -->
   <div class="msgBody">
   <% = strReplyMessage %>
   </div>
   <!-- End Member Post -->
  </td>
 </tr>
</table><%

'Else there is an error so show error table
Else

	'Update active users table
	If blnActiveUsers Then saryActiveUsers = activeUsers(strTxtWritingReply & " [" & strTxtAccessDenied & "]", strTopicSubject, "forum_posts.asp?TID=" & lngTopicID, intForumID)

%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><%

	'If the users account is suspended then let them know
	If blnActiveMember = False OR blnBanned Then
			
		'If mem suspended display message
		If blnBanned Then
			Response.Write(strTxtForumMemberSuspended)
		'Else account not yet active
		Else
			Response.Write("<br />" & strTxtForumMembershipNotAct)
			If blnMemberApprove = False Then Response.Write("<br /><br />" & strTxtToActivateYourForumMem)
		
			'If admin activation is enabled let the user know
			If blnMemberApprove Then
				Response.Write("<br />" & strTxtYouAdminNeedsToActivateYourMembership)
			'If email is on then place a re-send activation email link
			ElseIf blnEmailActivation AND blnLoggedInUserEmail Then 
				Response.Write("<br /><br /><a href=""javascript:winOpener('resend_email_activation.asp" & strQsSID1 & "','actMail',1,1,475,300)"">" & strTxtResendActivationEmail & "</a>")
			End If
		End If

	'Else if the forum is locked display a message telling the user so
	ElseIf blnForumLocked Then
		
		Response.Write(strTxtForumLockedByAdmim)
	
	'Display message if the topic is locked
	ElseIf blnTopicLocked Then
	
		Response.Write(strTxtSorryNoReply & "<br />" & strTxtThisTopicIsLocked)
	
	'Else if the user does not have permision to reply in this forum
	ElseIf blnReply = False AND intGroupID <> 2 Then
		
		Response.Write(strTxtSorryYouDoNotHavePerimssionToReplyToPostsInThisForum & "<br /><br />")
		Response.Write("<a href=""javascript:history.back(1)"">" & strTxtReturnForumTopic & "</a>")
	
	'Display message if this is a poll only
	ElseIf blnPollNoReply Then

		Response.Write(strTxtThisIsAPollOnlyYouCanNotReply & "<br /><br />")
		Response.Write("<a href=""javascript:history.back(1)"">" & strTxtReturnForumTopic & "</a>")

		
	'Else the user doesn't have permission to reply in this forum
	Else
		Response.Write(strTxtSorryYouDoNotHavePerimssionToReplyToPostsInThisForum )
	End If



%></td>
  </tr>
</table><%

	'If the user can needs to login display login box
	If blnReply = False AND intGroupID = 2  AND blnActiveMember AND blnForumLocked = false AND blnTopicLocked = false AND blnBanned = False Then 
		%><!--#include file="includes/login_form_inc.asp" --><%
	End If

	
End If

'Clean up
Call closeDatabase()


%>
<br />
<div align="center"><%


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