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
Dim lngTopicID
Dim strSubject


'Read in the topic ID number
lngTopicID = LngC(Request("TID"))
strMode = Request.QueryString("mode")


'If the person is not an admin or a moderator then send them away
If lngTopicID = "" OR bannedIP() OR  blnActiveMember = False OR blnBanned Then
	
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
strSQL = "SELECT " & strDbTable & "Topic.Forum_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Priority, " & strDbTable & "Topic.Locked, " & strDbTable & "Topic.Poll_ID, " & strDbTable & "Topic.Moved_ID, " & strDbTable & "Topic.Hide, " & strDbTable & "Topic.Icon, " & strDbTable & "Topic.Event_date, " & strDbTable & "Topic.Event_date_end " & _
"FROM " & strDbTable & "Topic" & strRowLock & " " & _
"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & ";"


'Set the cursor	type property of the record set	to Forward Only
rsCommon.CursorType = 0

'Set the Lock Type for the records so that the record set is only locked when it is updated
rsCommon.LockType = 3

'Query the database
rsCommon.Open strSQL, adoCon

'If there is a record returened read in the forum ID
If NOT rsCommon.EOF Then
	intForumID = CInt(rsCommon("Forum_ID"))
	strSubject = rsCommon("Subject")
End If


'Call the moderator function and see if the user is a moderator
If blnAdmin = False Then blnModerator = isModerator(intForumID, intGroupID)

'Close the rs
rsCommon.Close




'Check that the user is admin
If blnAdmin OR blnModerator Then
	
	'Get the Forum from the database to be locked
	If strMode = "Lock" Then
		strSQL = "UPDATE " & strDbTable & "Topic" & strRowLock & " " & _
		"SET " & strDbTable & "Topic.Locked = " & strDBTrue & " " & _
		"WHERE " & strDbTable & "Topic.Topic_ID ="  & lngTopicID & ";"
		
		'Update log file
		If blnLoggingEnabled AND blnModeratorLogging Then Call logAction(strLoggedInUsername, "Locked Topic '" & decodeString(strSubject) & "' - TopicID " & lngTopicID)
		
	'Unlock forum
	ElseIf strMode = "UnLock" Then
		strSQL = "UPDATE " & strDbTable & "Topic" & strRowLock & " " & _
		"SET " & strDbTable & "Topic.Locked = " & strDBFalse & " " & _
		"WHERE " & strDbTable & "Topic.Topic_ID ="  & lngTopicID & ";"
		
		'Update log file
		If blnLoggingEnabled AND blnModeratorLogging Then Call logAction(strLoggedInUsername, "Unlocked Topic '" & decodeString(strSubject) & "' - TopicID " & lngTopicID)
		
	'Hide topic
	ElseIf strMode = "Hide" Then
		strSQL = "UPDATE " & strDbTable & "Topic" & strRowLock & " " & _
		"SET " & strDbTable & "Topic.Hide = " & strDBTrue & " " & _
		"WHERE " & strDbTable & "Topic.Topic_ID ="  & lngTopicID & ";"
		
		'Update log file
		If blnLoggingEnabled AND blnModeratorLogging Then Call logAction(strLoggedInUsername, "Hide Topic '" & decodeString(strSubject) & "' - TopicID " & lngTopicID)
		
	'Show topic
	ElseIf strMode = "Show" Then
		strSQL = "UPDATE " & strDbTable & "Topic" & strRowLock & " " & _
		"SET " & strDbTable & "Topic.Hide = " & strDBFalse & " " & _
		"WHERE " & strDbTable & "Topic.Topic_ID ="  & lngTopicID & ";"
		
		'Update log file
		If blnLoggingEnabled AND blnModeratorLogging Then Call logAction(strLoggedInUsername, "Approved Topic '" & decodeString(strSubject) & "' - TopicID " & lngTopicID)
	End If
	
	'Write to the database
	adoCon.Execute(strSQL)
	
	'Update the stats for this topic in the tblTopic table
	Call updateTopicStats(lngTopicID)
	
	
	'Update the number of topics and posts in the database
	Call updateForumStats(intForumID)
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