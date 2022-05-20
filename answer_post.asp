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






'Set the response buffer to true as we maybe redirecting
Response.Buffer = True 

'Dimension variables
Dim strMode		'Holds the mode of the page
Dim lngPostID
Dim lngTopicID
Dim strSubject
Dim lngPostAuthorID
Dim lngAuthorAnswered
Dim lngAuthorPoints




'Read in the topic ID number
lngPostID = LngC(Request("PID"))
strMode = Request.QueryString("mode")


'If the person is not an admin or a moderator then send them away
If lngPostID = "" OR bannedIP() OR  blnActiveMember = False OR blnBanned Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If


'If in demo mode redirect
If blnDemoMode Then
	Call closeDatabase()
	Response.Redirect("insufficient_permission.asp?M=DEMO" & strQsSID3)
End If

'Check the form ID to prevent XCSRF
Call checkFormID(Request.QueryString("XID"))




'Initliase the SQL query to get the topic details from the database
strSQL = "SELECT " & strDbTable & "Topic.Forum_ID," & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Thread.Author_ID, " & strDbTable & "Thread.Answer, " & strDbTable & "Author.Answered, " & strDbTable & "Author.Points, " & strDbTable & "Permissions.Post " & _
"FROM " & strDbTable & "Topic" & strRowLock & ",  " & strDbTable & "Thread" & strRowLock & ",  " & strDbTable & "Author" & strRowLock & ", " & strDbTable & "Permissions" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID " & _
	" AND " & strDbTable & "Thread.Author_ID = " & strDbTable & "Author.Author_ID " & _
	" AND " & strDbTable & "Topic.Forum_ID = " & strDbTable & "Permissions.Forum_ID " & _
	" AND (" & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intGroupID & ") " & _
	" AND " & strDbTable & "Thread.Thread_ID = " & lngPostID & ";"

Response.Write(strSQL)

'Set the cursor	type property of the record set	to Forward Only
rsCommon.CursorType = 0

'Set the Lock Type for the records so that the record set is only locked when it is updated
rsCommon.LockType = 3

'Query the database
rsCommon.Open strSQL, adoCon

'If there is a record returened read in the forum ID
If NOT rsCommon.EOF Then
	intForumID = CInt(rsCommon("Forum_ID"))
	lngTopicID = CLng(rsCommon("Topic_ID"))
	lngPostAuthorID = CLng(rsCommon("Author_ID"))
	strSubject = rsCommon("Subject")
	If isNull(rsCommon("Answered")) Then lngAuthorAnswered = 0 Else lngAuthorAnswered = CLng(rsCommon("Answered"))
	If isNull(rsCommon("Points")) Then lngAuthorPoints = 0 Else lngAuthorPoints = CLng(rsCommon("Points"))
	blnPost = CBool(rsCommon("Post"))
End If

'Call the moderator function and see if the user is a moderator
If blnAdmin = False Then blnModerator = isModerator(intForumID, intGroupID)

'Close the rs
rsCommon.Close



'Check that the user is admin
If ((strAnswerPosts = "admin" AND blnAdmin) OR (strAnswerPosts = "admin_mods" AND blnAdmin OR blnModerator) OR (strAnswerPosts = "all_members" AND blnGuest = False AND blnPost)) Then
	
	'Get the Forum from the database to be locked
	If strMode = "Set" Then
		
		'Incremenet the number of anwsers the user has made
		lngAuthorAnswered = lngAuthorAnswered + 1
		lngAuthorPoints = lngAuthorPoints + intPointsAnswered
		
		'SQL to add the answered post
		strSQL = "UPDATE " & strDbTable & "Thread" & strRowLock & " " & _
		"SET " & strDbTable & "Thread.Answer = " & strDBTrue & " " & _
		"WHERE " & strDbTable & "Thread.Thread_ID = " & lngPostID & ";"
		
		'Update log file
		If blnLoggingEnabled AND blnModeratorLogging Then Call logAction(strLoggedInUsername, "Set Anwser Post '" & decodeString(strSubject) & "' - PostID " & lngPostID)
	
	ElseIf strMode = "Remove" Then
		
		'Decrement the number of anwsers the user has made
		lngAuthorAnswered = lngAuthorAnswered - 1
		lngAuthorPoints = lngAuthorPoints - intPointsAnswered
		
		'SQL to remove the answered post
		strSQL = "UPDATE " & strDbTable & "Thread" & strRowLock & " " & _
		"SET " & strDbTable & "Thread.Answer = " & strDBFalse & " " & _
		"WHERE " & strDbTable & "Thread.Thread_ID = " & lngPostID & ";"
		
		'Update log file
		If blnLoggingEnabled AND blnModeratorLogging Then Call logAction(strLoggedInUsername, "Removed Anwser Post '" & decodeString(strSubject) & "' - PostID " & lngPostID)
	
	End If
	
	'Write to the database
	adoCon.Execute(strSQL)
	
	
	'Updated the number of anwsers the user has preposed
	strSQL = "UPDATE " & strDbTable & "Author" & strRowLock & " " & _
	"SET " & strDbTable & "Author.Answered = " & lngAuthorAnswered & ", " & strDbTable & "Author.Points = " & lngAuthorPoints & " " & _
	"WHERE " & strDbTable & "Author.Author_ID = " & lngPostAuthorID & ";"
	
	'Write to the database
	adoCon.Execute(strSQL)
	
	
End If


'Reset Server Objects
Call closeDatabase()

'If this is a lock from the admin area then return there
If Request.QueryString("FID") <> "" Then
	Response.Redirect("forum_topics.asp?FID=" & Request.QueryString("FID") & strQsSID3)
Else
	'Return to the page showing the threads
	Response.Redirect("forum_posts.asp?TID=" & lngTopicID & strQsSID3)
End If
%>