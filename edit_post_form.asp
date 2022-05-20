<% @ Language=VBScript %>
<% Option Explicit %>
<% Response.AddHeader "X-UA-Compatible","IE=EmulateIE10" %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_edit_post.asp" -->
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
Dim strTopicSubject		'Holds the title of the topic
Dim intTopicPriority		'Holds the priority of the topic
Dim lngMessageID		'Holds the message ID to be edited
Dim strQuoteUsername		'Holds the quoters username
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
Dim lngPollEditID		'Holds the poll ID to be edited
Dim intPollLoopCounter		'Holds the poll loop counter
Dim strUploadedFiles		'Holds the names of any files or images uploaded
Dim strTopicIcon		'Holds the topic icon
Dim intEventYear		'Holds the year of Calendar event
Dim intEventMonth		'Holds the month of Calendar event
Dim intEventDay			'Holds the day of Calendar event
Dim dtmEventDate		'Holds the Calendar event date
Dim intEventYearEnd		'Holds the year of Calendar event
Dim intEventMonthEnd		'Holds the month of Calendar event
Dim intEventDayEnd		'Holds the day of Calendar event
Dim dtmEventDateEnd		'Holds the Calendar event date
Dim strFormID			'Holds the ID for the form
Dim dtmMessageDateTime		'Holds the date and time the message was created


'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)

End If



'Read in the message ID number to edit
lngMessageID = LngC(Request.QueryString("PID"))
intRecordPositionPageNum = IntC(Request.QueryString("PN"))
lngPollEditID = LngC(Request.QueryString("POLL"))
strMode = "edit"




'Initalise the strSQL variable with an SQL statement to query the database get the message details
strSQL = "SELECT " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Forum_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Icon, " & strDbTable & "Topic.Start_Thread_ID, " & strDbTable & "Topic.Priority, " & strDbTable & "Topic.Locked, " & strDbTable & "Topic.Event_date, " & strDbTable & "Topic.Event_date_end, " & strDbTable & "Thread.Thread_ID, " & strDbTable & "Thread.Author_ID, " & strDbTable & "Thread.Message, " & strDbTable & "Thread.Message_date " & _
"FROM " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & ", " & strDbTable & "Author" & strDBNoLock & " " & _
"WHERE (" & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID AND " & strDbTable & "Thread.Author_ID = " & strDbTable & "Author.Author_ID) AND " & strDbTable & "Thread.Thread_ID=" & lngMessageID & ";"
	
'Query the database
rsCommon.Open strSQL, adoCon 

'If there is no post to edit redirect to homepage
If rsCommon.EOF Then
	
	'Clean up
	rsCommon.Close
	Call closeDatabase()

	'Redirect
	Response.Redirect("default.asp")
	
End If

	
	
'Read in the details from the recordset
blnTopicLocked = CBool(rsCommon("Locked"))
lngTopicID = CLng(rsCommon("Topic_ID"))
intForumID = CInt(rsCommon("Forum_ID"))
lngPostUserID = CLng(rsCommon("Author_ID"))
strTopicSubject = rsCommon("Subject")
strTopicIcon = rsCommon("Icon")
intTopicPriority = CInt(rsCommon("Priority"))
strMessage = rsCommon("Message")
dtmEventDate = rsCommon("Event_date")
dtmEventDateEnd = rsCommon("Event_date_end")
dtmMessageDateTime =  CDate(rsCommon("Message_date"))


'Clean up input to prevent XXS hack
strTopicSubject = formatInput(strTopicSubject)

'If this is the first post in the topic then allow the editor to edit the topic subject etc. as well
If rsCommon("Start_Thread_ID") = rsCommon("Thread_ID") Then strMode = "editTopic"
	
'If we are editing a poll change the mode
If rsCommon("Start_Thread_ID") = rsCommon("Thread_ID") AND lngPollEditID > 0 Then strMode = "editPoll"
	

'Clean up
rsCommon.Close	


'Split the start date of event into the various parts
If isDate(dtmEventDate) Then
	intEventYear = Year(dtmEventDate)
	intEventMonth = Month(dtmEventDate)
	intEventDay = Day(dtmEventDate)
End If

'Split the end date of event into the various parts
If isDate(dtmEventDateEnd) Then
	intEventYearEnd = Year(dtmEventDateEnd)
	intEventMonthEnd = Month(dtmEventDateEnd)
	intEventDayEnd = Day(dtmEventDateEnd)
End If




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
	blnEdit = CBool(rsCommon("Edit_posts"))
	blnPriority = CBool(rsCommon("Priority_posts"))
	blnPollCreate = CBool(rsCommon("Poll_create"))
	blnVote = CBool(rsCommon("Vote"))
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
	
	'See if editing is within the time limit
	If intEditPostTimeFrame > 0 AND (blnAdmin = False AND blnModerator = False) Then
		If DateDiff("n", dtmMessageDateTime, Now()) >= intEditPostTimeFrame Then 
			
			'Reset Server Objects
			rsCommon.Close
			Call closeDatabase()
		
			'Redirect to a page asking for the user to enter the forum password
			Response.Redirect("insufficient_permission.asp?M=eExp" & strQsSID3)
			
		End If
	End If
End If
	
'Close rs
rsCommon.close




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



'If the forum level for the user on this forum is read only set the forum to be locked
If (blnRead = False AND blnModerator = False AND blnAdmin = False) Then blnForumLocked = True

'If the forums not locked check that the topics not locked either
If blnForumLocked = False Then blnForumLocked = blnTopicLocked


'Apply forum codes
strMessage = EditPostConvertion(strMessage)	

'If the message has been edited remove who edited the post
If InStr(1, strMessage, "<edited>", 1) Then strMessage = removeEditorAuthor(strMessage)
	




'Use the application session to pass around what forum this user is within
If IntC(getSessionItem("FID")) <> intForumID Then Call saveSessionItem("FID", intForumID)
	
'get the session key
strFormID = getSessionItem("KEY")



'Set bread crumb trail
'Display the category name
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""default.asp?C=" & intCatID & strQsSID2 & """>" & strCatName & "</a>" & strNavSpacer

'Display if there is a main forum to the sub forums name
If intMasterForumID <> 0 Then strBreadCrumbTrail = strBreadCrumbTrail & "<a href=""forum_topics.asp?FID=" & intMasterForumID & strQsSID2 & """>" & strMasterForumName & "</a>" & strNavSpacer

'Display forum name
If strForumName = "" Then strBreadCrumbTrail = strBreadCrumbTrail &  strTxtNoForums Else strBreadCrumbTrail = strBreadCrumbTrail & "<a href=""forum_topics.asp?FID=" & intForumID & strQsSID2 & """>" & strForumName & "</a>"

strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtEditPost


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strTxtEditPost %></title>
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

<script  language="JavaScript">

//Function to check form is filled in correctly before submitting
function CheckForm () {
	
	var errorMsg = "";
	var formArea = document.getElementById('frmMessageForm');
<%
'If Gecko Madis API (RTE) need to strip default input from the API
If RTEenabled = "Gecko" Then Response.Write("	//For Gecko Madis API (RTE)" & vbCrLf & "	if (formArea.message.value.indexOf('<br>') > -1 && formArea.message.value.length==6) formArea.message.value = '';" & vbCrLf)


'If we are editing the first topic then  check for a subject
If strMode = "editTopic" OR strMode = "editPoll" Then
%>	
	//Check for a subject
	if (formArea.subject.value==""){
		errorMsg += "\n<% = strTxtErrorTopicSubject %>";
	}<%
End If


'If we are editing a poll then check for poll question and choices
If strMode = "editPoll" Then
%>	

	//Check for poll question
	if (formArea.pollQuestion.value==""){
		errorMsg += "\n<% = strTxtErrorPollQuestion %>";
	}
	
	//Check for poll at least two poll choices
	if ((formArea.choice1.value=="") || (formArea.choice2.value=="")){
		errorMsg += "\n<% = strTxtErrorPollChoice %>";
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
  <td align="left"><h1><% = strTxtEditPost & " - " & strTopicSubject %></h1></td>
 </tr>
</table>
<br /><%
 
'If the Post is by the logged in user or the adminstrator/moderator then display a form to edit the post
If ((lngLoggedInUserID = lngPostUserID OR blnAdmin OR blnModerator) AND (blnEdit OR blnAdmin) AND (strMode="edit" OR strMode="editTopic" OR strMode = "editPoll")) AND blnActiveMember AND (blnForumLocked = False OR blnAdmin) AND (blnTopicLocked = False Or blnAdmin) Then

	'Update active users table array
	If blnActiveUsers Then saryActiveUsers = activeUsers(strTxtEditingPost, strTopicSubject, "forum_posts.asp?TID=" & lngTopicID, intForumID)

	%>
<table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td align="left"><% = strTxtEditPost %></td>
 </tr>
 <tr class="tableRow">
  <td align="left">
   <!--#include file="includes/message_form_inc.asp" -->
  </td>
 </tr>
</table><%

'Else there is an error so show error table
Else

	'Update active users table array
	If blnActiveUsers Then saryActiveUsers = activeUsers(strTxtEditingPost & " [" & strTxtAccessDenied & "]", strTopicSubject, "forum_posts.asp?TID=" & lngTopicID, intForumID)

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
			Response.Write("<br /><br />" & strTxtForumMembershipNotAct)
			If blnMemberApprove = False Then Response.Write("<br /><br />" & strTxtToActivateYourForumMem)
		
			'If admin activation is enabled let the user know
			If blnMemberApprove Then
				Response.Write("<br /><br />" & strTxtYouAdminNeedsToActivateYourMembership)
			'If email is on then place a re-send activation email link
			ElseIf blnEmailActivation AND blnLoggedInUserEmail Then 
				Response.Write("<br /><br /><a href=""javascript:winOpener('resend_email_activation.asp" & strQsSID1 & "','actMail',1,1,475,300)"">" & strTxtResendActivationEmail & "</a>")
			End If
		End If
		

	'Else if the forum is locked display a message telling the user so
	ElseIf blnForumLocked Then
		
		Response.Write(strTxtForumLockedByAdmim)
	
	'Else the user is not the person who posted the message so display an error message
	Else 
	   	Response.Write(strTxtNoPermissionToEditPost & "<br /><br />")
		Response.Write("<a href=""javascript:history.back(1)"">" & strTxtReturnForumTopic & "</a>")
	End If 
	
%></td>
  </tr>
</table>
<%
End If

'Clean up	
Call closeDatabase()


%><br />
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