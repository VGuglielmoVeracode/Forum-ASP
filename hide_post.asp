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
Dim lngTopicID			'Holds the Topic ID number
Dim lngMessageID		'Holds the message ID to be deleted
Dim strSubject


'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)

End If

'Check the form ID to prevent XCSRF
Call checkFormID(Request.QueryString("XID"))



'Read in the message ID number to be hidden
lngMessageID = LngC(Request.QueryString("PID"))



'Read in the forum and topic ID from the database for this message

'Initliase the SQL query to get the topic and forumID from the database
strSQL = "SELECT " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Forum_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Thread.Author_ID, " & strDbTable & "Author.No_of_posts " & _
"FROM " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & ", " & strDbTable & "Author" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID AND " & strDbTable & "Thread.Author_ID = " & strDbTable & "Author.Author_ID AND " & strDbTable & "Thread.Thread_ID=" & lngMessageID & ";"
	
'Query the database
rsCommon.Open strSQL, adoCon 

'If there is a record returened read in the forum ID
If NOT rsCommon.EOF Then
	lngTopicID = CLng(rsCommon("Topic_ID"))
	intForumID = CInt(rsCommon("Forum_ID"))
	strSubject = rsCommon("Subject")
End If

'Clean up
rsCommon.Close
	



'******************************************
'***	    Check permissions		***
'******************************************

'Check the users permissions
Call forumPermissions(intForumID, intGroupID)




'Get the Post to be hidden
If blnAdmin OR blnModerator Then 		
	
	'******************************************
	'***		Hide post		***
	'******************************************
			
	'Initalise the SQL string with an SQL update command to	update the post to be hidden
	strSQL = "UPDATE " & strDbTable & "Thread" & strRowLock & " " & _
	"SET " & strDbTable & "Thread.Hide = " & strDBTrue & " " & _
	"WHERE " & strDbTable & "Thread.Thread_ID = " & lngMessageID & ";"
			
	'Write the updated number of posts to the database
	adoCon.Execute(strSQL)
	
	
	
	'******************************************
	'***		Hide Topic????				***
	'******************************************
	
	'See if there are any posts in the topic
	strSQL = "SELECT " & strDbTable & "Thread.Thread_ID " & _
	"FROM " & strDbTable & "Thread" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Thread.Hide = " & strDBFalse & " AND " & strDbTable & "Thread.Topic_ID = " & lngTopicID & ";"
	
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'If no record is returned hide the topic as well
	If rsCommon.EOF Then
	
		'Initalise the SQL string with an SQL update command to	update the topic to be hidden
		strSQL = "UPDATE " & strDbTable & "Topic" & strRowLock & " " & _
		"SET " & strDbTable & "Topic.Hide = " & strDBTrue & " " & _
		"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & ";"
				
		'Write the updated number of posts to the database
		adoCon.Execute(strSQL)
		
	End If
	
	'Close Rs
	rsCommon.Close
	
	
	'If logging enabled log the new topic has been created
	If blnLoggingEnabled AND blnModeratorLogging Then Call logAction(strLoggedInUsername, "Hidden Post in Topic '" & decodeString(strSubject) & "' - TopicID " & lngTopicID)
	
	
	
	'***********************************
	'***	   Update  Stats   	 ***
	'***********************************
	
	'Update the stats for this topic in the tblTopic table
	Call updateTopicStats(lngTopicID)
	
	'Update the number of topics and posts in the database
	Call updateForumStats(intForumID)
End If






'Reset Server Objects
Call closeDatabase()



'Return to the page showing the threads
Response.Redirect("forum_posts.asp?TID=" & lngTopicID & "&PID=" & lngMessageID & strQsSID3 & "#" & lngMessageID)
%>