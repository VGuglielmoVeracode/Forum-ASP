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



'Set the timeout of the page
Server.ScriptTimeout = 1000


'Set the response buffer to true as we maybe redirecting
Response.Buffer = True


'Use sll for this page (required as POP Up Topic Admin is SSL)
blnSslEnabledPage = True


'Dimension variables
Dim rsNumOfPosts		'Holds the database recordset for the number of posts the user has made
Dim rsForum			'Holds the forum for order
Dim strMode			'Holds the mode of the page
Dim lngTopicID 			'Holds the topic ID number to return to
Dim lngPollID			'Holds the poll ID number if there is one
Dim lngDelMsgAuthorID		'Holds the deleted message Author ID
Dim lngNumOfPosts		'Holds the number of posts the user has made
Dim lngNumOfPoints		'Holds the number of points the user has
Dim strSubject			'Holds the topic subject


'If in demo mode redirect
If blnDemoMode Then
	Call closeDatabase()
	Response.Redirect("insufficient_permission.asp?M=DEMO" & strQsSID3)
End If



'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)
End If





'Check the form ID to prevent XCSRF
Call checkFormID(Request("XID"))


'Read in the message ID number to be deleted
lngTopicID = LngC(Request("TID"))




'Initliase the SQL query to get the topic and forumID from the database
strSQL = "SELECT " & strDbTable & "Topic.Forum_ID, " & strDbTable & "Topic.Subject " & _
"FROM " & strDbTable & "Topic " & strDBNoLock & " " & _
"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & ";"

'Query the database
rsCommon.Open strSQL, adoCon

'If there is a record returned read in the forum ID
If NOT rsCommon.EOF Then
	intForumID = CInt(rsCommon("Forum_ID"))
	strSubject = rsCommon("Subject")
End If

'Clean up
rsCommon.Close



'Call the moderator function and see if the user is a moderator
If blnAdmin = False Then blnModerator = isModerator(intForumID, intGroupID)



'Check to make sure the user is deleting the topic is a moderator or the forum adminstrator
If blnAdmin = True OR blnModerator = True Then

	'See if there is a poll, if there is get the poll ID and delete

	'Initalise the strSQL variable 
	strSQL = "SELECT " & strDbTable & "Topic.Poll_ID " & _
	"FROM " & strDbTable & "Topic " & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Topic.Topic_ID = "  & lngTopicID  & ";"

	'Query the database
	rsCommon.Open strSQL, adoCon

	'Get the Poll ID
	If NOT rsCommon.EOF Then lngPollID = CLng(rsCommon("Poll_ID"))

	'Close the recordset
	rsCommon.Close


	'Get the Posts to be deleted from the database
	strSQL = "SELECT " & strDbTable & "Thread.* " & _
	"FROM " & strDbTable & "Thread " & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Thread.Topic_ID = "  & lngTopicID  & ";"

	'Query the database
	rsCommon.Open strSQL, adoCon

	'Get the number of posts the user has made and take one away
	Set rsNumOfPosts = Server.CreateObject("ADODB.Recordset")


	'Loop through all the posts for the topic and delete them
	Do While NOT rsCommon.EOF
	
		'First we need to delete any entry in the GuestName table incase this was a guest poster posting the message
		strSQL = "DELETE FROM " & strDbTable & "GuestName " & strRowLock & " " & _
		"WHERE " & strDbTable & "GuestName.Thread_ID = " & CLng(rsCommon("Thread_ID")) & ";"
	
		'Excute SQL
		adoCon.Execute(strSQL)
		

		'Initalise the strSQL variable with an SQL statement to query the database to get the number of posts the user has made
		strSQL = "SELECT " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.No_of_posts, " & strDbTable & "Author.Points " & _
		"FROM " & strDbTable & "Author " & strDBNoLock & " " & _
		"WHERE " & strDbTable & "Author.Author_ID = " & CLng(rsCommon("Author_ID")) & ";"

		'Query the database
		rsNumOfPosts.Open strSQL, adoCon

		'If there is a record returned by the database then read in the no of posts and decrement it by 1
		If NOT rsNumOfPosts.EOF Then

			'Read in the no of posts the user has made and username
			lngNumOfPosts = CLng(rsNumOfPosts("No_of_posts"))
			If isNull(rsNumOfPosts("Points")) = False Then lngNumOfPoints = CLng(rsNumOfPosts("Points")) Else lngNumOfPoints = 0

			'Decrement by 1 unless the number of posts is already 0
			If lngNumOfPosts > 0 Then

				'decrement the number of posts by 1
				lngNumOfPosts = lngNumOfPosts - 1
				lngNumOfPoints = lngNumOfPoints - intPointsReply

				'Initalise the SQL string with an SQL update command to update the number of posts the user has made
				strSQL = "UPDATE " & strDbTable & "Author " & strRowLock & " " & _
				"SET " & strDbTable & "Author.No_of_posts = " & lngNumOfPosts & ", " & strDbTable & "Author.Points = " & lngNumOfPoints & " " & _
				"WHERE " & strDbTable & "Author.Author_ID = " & CLng(rsCommon("Author_ID")) & ";"

				'Write the updated number of posts to the database
				adoCon.Execute(strSQL)
			End If
		End If

		'Close the recordset
		rsNumOfPosts.Close

		'Move to the next record
		rsCommon.MoveNext
	Loop
	

	'Delete the posts in this topic
	strSQL = "DELETE FROM " & strDbTable & "Thread " & strRowLock & " WHERE " & strDbTable & "Thread.Topic_ID = "  & lngTopicID & ";"

	'Write to database
	adoCon.Execute(strSQL)


	'Delete the Poll in this topic, if there is one
	If lngPollID > 0 Then

		'Delete the Poll choices
		strSQL = "DELETE FROM " & strDbTable & "PollChoice " & strRowLock & " WHERE " & strDbTable & "PollChoice.Poll_ID = "  & lngPollID & ";"

		'Write to database
		adoCon.Execute(strSQL)
		
		'Delete the Poll Votes 
		strSQL = "DELETE FROM " & strDbTable & "PollVote " & strRowLock & " WHERE " & strDbTable & "PollVote.Poll_ID = " & lngPollID & ";" 
			
		'Write to database 
		adoCon.Execute(strSQL)

		'Delete the Poll
		strSQL = "DELETE FROM " & strDbTable & "Poll " & strRowLock & " WHERE " & strDbTable & "Poll.Poll_ID = "  & lngPollID & ";"

		'Write to database
		adoCon.Execute(strSQL)
	End If

	'delete any rating for this topic
	strSQL = "DELETE FROM " & strDbTable & "TopicRatingVote " & strRowLock & " " & _
	"WHERE " & strDbTable & "TopicRatingVote.Topic_ID = " & lngTopicID & ";"
		
	'Excute SQL
	adoCon.Execute(strSQL)
	
	'delete any email notifications for this topic
	strSQL = "DELETE FROM " & strDbTable & "EmailNotify " & strRowLock & " " & _
	"WHERE " & strDbTable & "EmailNotify.Topic_ID = " & lngTopicID & ";"
		
	'Excute SQL
	adoCon.Execute(strSQL)

	'Delete the topic from the database
	'Initalise the strSQL variable with an SQL statement to get the topic from the database
	strSQL = "DELETE FROM " & strDbTable & "Topic " & strRowLock & " WHERE " & strDbTable & "Topic.Topic_ID = "  & lngTopicID & ";"

	'Write the updated date of last post to the database
	adoCon.Execute(strSQL)

	'Reset Server Objects
	rsCommon.Close
	Set rsNumOfPosts = Nothing
	
	'If loging is enabled write to log file
	If blnLoggingEnabled AND (blnDeletePostLogging OR (blnModeratorLogging AND (blnAdmin OR blnModerator))) Then Call logAction(strLoggedInUsername, "Deleted Topic '" & decodeString(strSubject) & "' - TopicID " & lngTopicID)
		
End If


'Update the number of topics and posts in the database
Call updateForumStats(intForumID)


'Reset main server variables
Call closeDatabase()

'If from topic index and not tool options redirect back
If Request.QueryString("PN") <> "" Then Response.Redirect("forum_topics.asp?FID=" & intForumID & "&PN=" & Server.URLEncode(Request.QueryString("PN")) & strQsSID3)

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<script language="JavaScript">
	window.opener.location.href = 'forum_topics.asp?FID=<% = intForumID %>&DL=1<% = strQsSID2 %>'
	window.close();
</script>
</head>
</html>