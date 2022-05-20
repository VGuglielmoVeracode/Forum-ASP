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
Dim strMode			'Holds the mode of the page
Dim lngTopicID			'Holds the Topic ID number
Dim lngMessageID		'Holds the message ID to be deleted
Dim lngDelMsgAuthorID		'Holds the deleted message Author ID
Dim lngNumOfPosts		'Holds the number of posts the user has made
Dim lngLastPostID		'Holds the last post ID
Dim lngNumOfPoints		'Holds the number of points the user has
Dim strSubject			'Holds the topic subject



'If in demo mode redirect
If blnDemoMode Then
	Call closeDatabase()
	Response.Redirect("insufficient_permission.asp?M=DEMO" & strQsSID3)
End If


'Inti
lngLastPostID = 0


'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)

End If


'Check the form ID to prevent XCSRF
Call checkFormID(Request.QueryString("XID"))


'Read in the message ID number to be deleted
lngMessageID = LngC(Request.QueryString("PID"))



'Read in the forum and topic ID from the database for this message

'Initliase the SQL query to get the topic and forumID from the database
strSQL = "SELECT " & strDbTable & "Topic.Topic_ID, " & strDbTable & "Topic.Forum_ID, " & strDbTable & "Topic.Subject " & _
"FROM " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID " & _
	"AND " & strDbTable & "Thread.Thread_ID = " & lngMessageID & ";"
	
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
	

'Get the users permissions
Call forumPermissions(intForumID, intGroupID)





'Get the Post to be deleted from the database
	
'Initalise the strSQL variable with an SQL statement to get the post from the database
strSQL = "SELECT " & strDbTable & "Thread.* " & _
"FROM " & strDbTable & "Thread" & strRowLock & " " & _
"WHERE " & strDbTable & "Thread.Thread_ID = "  & lngMessageID & ";"

'Set the cursor type property of the record set to Forward Only
rsCommon.CursorType = 0

'Set set the lock type of the recordset to optomistic while the record is deleted
rsCommon.LockType = 3

'Query the database
rsCommon.Open strSQL, adoCon  

'Read in the author ID of the message to be deleted
If NOT rsCommon.EOF Then lngDelMsgAuthorID = CLng(rsCommon("Author_ID"))




'Check to make sure the user is deleting the post enetered the post or a moderator with detlete rights or the forum adminstrator
If (lngDelMsgAuthorID = lngLoggedInUserID OR blnAdmin = True OR blnModerator = True) AND (blnDelete = True OR blnAdmin = True) Then



	'First we need to delete any entry in the GuestName table incase this was a guest poster posting the message
	strSQL = "DELETE FROM " & strDbTable & "GuestName " & strRowLock & " " & _
	"WHERE " & strDbTable & "GuestName.Thread_ID = "  & lngMessageID & ";"

	'Excute SQL
	adoCon.Execute(strSQL)
	
	
	'Delete Post SQL
	strSQL = "DELETE FROM " & strDbTable & "Thread" & strRowLock & " " & _
	"WHERE " & strDbTable & "Thread.Thread_ID = "  & lngMessageID & ";"
	
	'Excute SQL
	adoCon.Execute(strSQL)

	
	'We need to requry the database before moving on as Access can take a few moments to delete the record
	rsCommon.Requery
	
	'Close the recordset
	rsCommon.Close
	
	
		
	'Initalise the strSQL variable with an SQL statement to query the database to get the number of posts the user has made
	strSQL = "SELECT " & strDbTable & "Author.No_of_posts, " & strDbTable & "Author.Points " & _
	"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Author.Author_ID = " & lngDelMsgAuthorID & ";"
		
	'Query the database
	rsCommon.Open strSQL, adoCon
		
	'If there is a record returned by the database then read in the no of posts and decrement it by 1
	If NOT rsCommon.EOF Then
		
		'Read in the no of posts the user has made and username
		lngNumOfPosts = CLng(rsCommon("No_of_posts"))
		If isNull(rsCommon("Points")) = False Then lngNumOfPoints = CLng(rsCommon("Points")) Else lngNumOfPoints = 0
		
		'decrement the number of posts by 1
		lngNumOfPosts = lngNumOfPosts - 1
		
		'Incase the number of posts is less than 0
		If lngNumOfPosts < 0 Then lngNumOfPosts = 0
	End If
		
	'Close the recordset
	rsCommon.Close
	
	
	
	
	'Check there are other Posts for the Topic, if not delete the topic as well	
	'Initalise the strSQL variable with an SQL statement to get the Threads from the database
	strSQL = "SELECT " & strDBTop1 & " " & strDbTable & "Thread.Thread_ID " & _
	"FROM " & strDbTable & "Thread" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Thread.Topic_ID = "  & lngTopicID & " " & _
	"ORDER BY " & strDbTable & "Thread.Message_date ASC" & strDBLimit1 & ";"
	
	'Query the database
	rsCommon.Open strSQL, adoCon

	
	'If there are posts left in the database for this topic get some details for them
	If NOT rsCommon.EOF Then
		
		'Get the post ID of the last post
		lngLastPostID = CLng(rsCommon("Thread_ID"))
	End If
	
	'Close the recordset
	rsCommon.Close
	

	
	'Read in details of the last topic to either update or delete (depends if any topics are left in db)
	
	'Initalise the strSQL variable with an SQL statement to get the topic from the database
	strSQL = "SELECT " & strDbTable & "Topic.* " & _
	"FROM " & strDbTable & "Topic" & strRowLock & " " & _
	"WHERE " & strDbTable & "Topic.Topic_ID = "  & lngTopicID & ";"
			
	'Set the cursor type property of the record set to Forward Only
	rsCommon.CursorType = 0
			
	'Set set the lock type of the recordset to optomistic while the record is deleted
	rsCommon.LockType = 3
			
	'Query the database
	rsCommon.Open strSQL, adoCon 
	
	
	
	'If there are no other posts in the topic, delete the topic
	If lngLastPostID = 0 Then
		
		'Decreae number of user points
		lngNumOfPoints = lngNumOfPoints - intPointsTopic
		
		'If there is a poll and no more posts left delete the poll as well
		If CLng(rsCommon("Poll_ID")) <> 0 Then 
			
			'Delete the Poll choices 
			strSQL = "DELETE FROM " & strDbTable & "PollChoice " & strRowLock & " WHERE " & strDbTable & "PollChoice.Poll_ID = " & CLng(rsCommon("Poll_ID")) & ";" 
			
			'Write to database 
			adoCon.Execute(strSQL) 
			
			'Delete the Poll Votes 
			strSQL = "DELETE FROM " & strDbTable & "PollVote " & strRowLock & " WHERE " & strDbTable & "PollVote.Poll_ID = " & CLng(rsCommon("Poll_ID")) & ";" 
			
			'Write to database 
			adoCon.Execute(strSQL)
		
			'Delete the Poll 
			strSQL = "DELETE FROM " & strDbTable & "Poll " & strRowLock & " WHERE " & strDbTable & "Poll.Poll_ID = " & CLng(rsCommon("Poll_ID")) & ";" 
			
			'Write to database 
			adoCon.Execute(strSQL)  
		End If
		
		'delete any rating for this topic
		strSQL = "DELETE FROM " & strDbTable & "TopicRatingVote " & strRowLock & " " & _
		"WHERE " & strDbTable & "TopicRatingVote.Topic_ID = " & lngTopicID & ";"
		
		'Excute SQL
		adoCon.Execute(strSQL)
		
		
		'Delete Post Topic
		strSQL = "DELETE FROM " & strDbTable & "Topic" & strRowLock & " " & _
		"WHERE " & strDbTable & "Topic.Topic_ID = "  & lngTopicID & ";"
		
		'Excute SQL
		adoCon.Execute(strSQL)
		
		
		'SQL to update posts and points
		strSQL = "UPDATE " & strDbTable & "Author " & strRowLock & " " & _
		"SET " & strDbTable & "Author.No_of_posts = " & lngNumOfPosts & ", " & strDbTable & "Author.Points = " & lngNumOfPoints & " " & _
		"WHERE " & strDbTable & "Author.Author_ID = " & lngDelMsgAuthorID & ";"	
		
		'Write the updated number of posts to the database
		adoCon.Execute(strSQL)
		
		
		
		'Update the number of topics and posts in the database
		Call updateForumStats(intForumID)
		
		'Reset Server Objects
		rsCommon.Close
		Call closeDatabase()
		
		'Return to the page showing the the topics in the forum
		Response.Redirect("forum_topics.asp?FID=" & intForumID & "&PN=" & Request.QueryString("PN") & strQsSID3)
	
	
	
	
	
	'Else there are other posts in the topic, so let's update some details for the new last post
	Else 
		
		'Subtract number of user points
		lngNumOfPoints = lngNumOfPoints - intPointsReply
	
		'Close Rs
		rsCommon.Close
		
		'Update the Topic Stats for this topic
		Call updateTopicStats(lngTopicID)
	End If
	
	
	
	'SQL to update posts and points
	strSQL = "UPDATE " & strDbTable & "Author " & strRowLock & " " & _
	"SET " & strDbTable & "Author.No_of_posts = " & lngNumOfPosts & ", " & strDbTable & "Author.Points = " & lngNumOfPoints & " " & _
	"WHERE " & strDbTable & "Author.Author_ID = " & lngDelMsgAuthorID & ";"
		
	
	'Write the updated number of posts to the database
	adoCon.Execute(strSQL)
	
	
	'If loging is enabled write to log file
	If blnLoggingEnabled AND (blnDeletePostLogging OR (blnModeratorLogging AND (blnAdmin OR blnModerator))) Then Call logAction(strLoggedInUsername, "Deleted Post in '" & decodeString(strSubject) & "' - PostID " & lngMessageID)
	
	
	
Else
	rsCommon.Close
End If	
	

'Update the number of topics and posts in the database
Call updateForumStats(intForumID)


'Reset Server Objects
Call closeDatabase()


'Return to the page showing the threads
Response.Redirect("forum_posts.asp?FID=" & intForumID & "&TID=" & lngTopicID & "&PN=" & Request.QueryString("PN") & strQsSID3)
%>