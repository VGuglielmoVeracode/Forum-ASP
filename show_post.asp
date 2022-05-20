<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_send_mail.asp" -->
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

'Dimension variables
Dim lngTopicID			'Holds the Topic ID number
Dim lngMessageID		'Holds the message ID to be deleted
Dim lngNumOfPosts		'Holds the number of posts the user has made
Dim strForumName		'Holds the name of the forum posted in
Dim strSubject			'Holds the subject of the topic
Dim strPostMessage		'Holds the post
Dim strEmailSubject		'Holds the email subject
Dim strUserName			'Holds the posters username
Dim lngEmailUserID		'Holds the posters email ID
Dim strUserEmail		'Holds the posters email
Dim strEmailMessage		'Holds the posters post
Dim strMessage			'Holds the posters post
Dim lngPostersID		'Holds the posters ID
Dim dtmMessagePostDate		'Holds the date the message was posted
Dim strPostersUsername		'Holds the username of poster
Dim lngLastPostID		'Holds the last post ID


'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)

End If

'Check the form ID to prevent XCSRF
Call checkFormID(Request.QueryString("XID"))



'******************************************
'***	   Get topic data		***
'******************************************

'Read in the message ID number to be shown
lngMessageID = LngC(Request.QueryString("PID"))

'Read in the forum and topic ID from the database for this message

'Initliase the SQL query to get the topic and forumID from the database
strSQL = "SELECT " & strDbTable & "Forum.Forum_name, " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Forum_ID, " & strDbTable & "Thread.Author_ID, " & strDbTable & "Thread.Message_Date, " & strDbTable & "Thread.Message, " & strDbTable & "Author.Username " & _
"FROM	" & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & ", " & strDbTable & "Author" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Topic.Forum_ID AND " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID AND " & strDbTable & "Thread.Author_ID = " & strDbTable & "Author.Author_ID AND " & strDbTable & "Thread.Thread_ID=" & lngMessageID & ";"
	
'Query the database
rsCommon.Open strSQL, adoCon 

'If there is a record returened read in the forum ID
If NOT rsCommon.EOF Then
	strSubject = rsCommon("Subject")
	strForumName = rsCommon("Forum_name")
	strMessage = rsCommon("Message")
	lngTopicID = CLng(rsCommon("Topic_ID"))
	intForumID = CInt(rsCommon("Forum_ID"))
	lngPostersID = CLng(rsCommon("Author_ID"))
	strPostersUsername = rsCommon("Username")
	dtmMessagePostDate = CDate(rsCommon("Message_Date"))
End If

'Clean up
rsCommon.Close
	


'******************************************
'***	    Check permissions		***
'******************************************

'Check the users permissions
Call forumPermissions(intForumID, intGroupID)


'Get the Post to be shown from the database
If blnAdmin OR blnModerator AND lngTopicID <> "" Then 
	
	
	
	'******************************************
	'***		Show post		***
	'******************************************
	
	'Initalise the SQL string with an SQL update command to	update the post to be shown
	strSQL = "UPDATE " & strDbTable & "Thread" & strRowLock & " " & _
	"SET " & strDbTable & "Thread.Hide = " & strDBFalse & " " & _
	"WHERE " & strDbTable & "Thread.Thread_ID = " & lngMessageID & ";"
			
	'Write the updated number of posts to the database
	adoCon.Execute(strSQL)
	
	
	
	'******************************************
	'***		Show topic		***
	'******************************************
	
	'Initalise the SQL string with an SQL update command to	update the topic to be shown
	strSQL = "UPDATE " & strDbTable & "Topic" & strRowLock & " " & _
	"SET " & strDbTable & "Topic.Hide = " & strDBFalse & " " & _
	"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & ";"
			
	'Write the updated number of posts to the database
	adoCon.Execute(strSQL)
	
	
	'If logging enabled log the new topic has been created
	If blnLoggingEnabled AND blnModeratorLogging Then Call logAction(strLoggedInUsername, "Approved Post in Topic '" & decodeString(strSubject) & "' - TopicID " & lngTopicID)
	
	
	
	
	
	'***********************************
	'***	   Update  Stats   	 ***
	'***********************************
	
	'Update the stats for this topic in the tblTopic table
	Call updateTopicStats(lngTopicID)


	
	'******************************************
	'***	   Send	Email Notification	 **
	'******************************************

	If blnEmail Then

		'**********************************************************
		'*** Format the	post if	it is to be sent with the email	 **
		'**********************************************************


		'Set the e-mail	subject
		strEmailSubject	= decodeString(strSubject)

		'If we are to send an e-mail notification and send the post with the e-mail then format	the post for the e-mail
		If blnSendPost = True Then

			'Format	the post to be sent with the e-mail
			strPostMessage = "<br /><b>" & strTxtForum & ":</b> " &	strForumName & _
			"<br /><b>" &	strTxtTopic & ":</b> " & formatInput(strSubject) & _
			"<br /><b>" &	strTxtPostedBy & ":</b> " & strPostersUsername & _
			"<br /><b>" &	strTxtVerifiedBy & ":</b> " & strLoggedInUsername & "<br /><br />" & _
			strMessage

			'Change	the path to the	emotion	symbols	to include the path to the images
			strPostMessage = Replace(strPostMessage, "src=""smileys/", "src=""" & strForumPath & "smileys/", 1, -1, 1)
		End If



		'*******************************************
		'***	   Send	Email Notification	 ***
		'*******************************************

		'Initalise the strSQL variable with an SQL statement to	query the database get the details for the email
		strSQL = "SELECT DISTINCT " & strDbTable & "EmailNotify.Author_ID, " & strDbTable & "Author.Username,	" & strDbTable & "Author.Author_email "  & _
		"FROM	" & strDbTable & "Author" & strDBNoLock & ", " & strDbTable & "EmailNotify" & strDBNoLock & " "  & _
		"WHERE " & strDbTable & "Author.Author_ID = " & strDbTable & "EmailNotify.Author_ID "  & _
			"AND (" & strDbTable & "EmailNotify.Forum_ID = " & intForumID & " OR " & strDbTable & "EmailNotify.Topic_ID = " & lngTopicID & ") "  & _
			"AND " & strDbTable & "Author.Author_email Is Not Null "  & _
			"AND " & strDbTable & "Author.Author_email <> '' "  & _
			"AND " & strDbTable & "Author.Banned = " & strDBFalse & " " & _
			"AND " & strDbTable & "Author.Active=" & strDBTrue & " " & _
			"AND " & strDbTable & "Author.Last_visit > " & formatDbDate(DateAdd("yyyy", -1, now())) & ";"

		'Query the database
		rsCommon.Open strSQL, adoCon

		'If a record is	returned by the	recordset then read in the details and send the	e-mail
		Do While NOT rsCommon.EOF

			'Read in the details from the recordset	for the	e-mail
			strUserName = rsCommon("Username")
			lngEmailUserID = CLng(rsCommon("Author_ID"))
			strUserEmail = rsCommon("Author_email")

			'If the	user wants to be e-mailed and the user has enetered their e-mail and they are not the original topic writter then send an e-mail
			If lngEmailUserID <> lngPostersID AND lngEmailUserID <> lngLoggedInUserID AND Trim(strUserEmail) <> "" Then

				'Initailise the	e-mail body variable with the body of the e-mail
				strEmailMessage	= strTxtHi & " " & decodeString(strUserName) & "," & _
				"<br /><br />" & strTxtEmailAMeesageHasBeenPosted &	" " & strMainForumName & " " & strTxtThatYouAskedKeepAnEyeOn  & _
				"<br /><br />" & strTxtEmailClickOnLinkBelowToView & " : -" & _
				"<br /><a href=""" & strForumPath &	"forum_posts.asp?TID="	& lngTopicID & "&PID=" & lngMessageID & "&Type=EmailNotification#" & lngMessageID & """>" & strForumPath & "forum_posts.asp?TID=" & lngTopicID & "&PID=" & lngMessageID & "&Type=EmailNotification#" & lngMessageID & "</a>" & _
				"<br /><br />" & strTxtClickTheLinkBelowToUnsubscribe & " :	-" & _
				"<br /><a href=""" & strForumPath &	"email_notify.asp?TID=" & lngTopicID &	"&FID="	& intForumID & "&M=Unsubscribe"">" & strForumPath & "email_notify.asp?TID=" & lngTopicID & "&FID=" & intForumID & "&M=Unsubscribe</a>"

				'If we are to send the post then attach	it as well
				If blnSendPost = True Then
					strEmailMessage	= strEmailMessage & "<br /><br /><hr />" & strPostMessage
				End If

				'Call the function to send the e-mail
				Call SendMail(strEmailMessage, decodeString(strUserName), decodeString(strUserEmail), strWebsiteName, decodeString(strForumEmailAddress), decodeString(strEmailSubject), strMailComponent, true)
			End If

			'Move to the next record in the	recordset
			rsCommon.MoveNext
		Loop

		'Close the recordset
		rsCommon.Close
	End If
	
End If



'Update the number of topics and posts in the database
Call updateForumStats(intForumID)




'Reset Server Objects
Call closeDatabase()



'Return to the page showing the threads
Response.Redirect("forum_posts.asp?TID=" & lngTopicID & "&PID=" & lngMessageID & strQsSID3 & "#" & lngMessageID)
%>