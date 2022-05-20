<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="language_files/admin_language_file_inc.asp" -->
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
Dim lngTopicID 		'Holds the topic ID number to return to
Dim lngPollID		'Holds the poll ID
Dim strMode		'Holds the page mode
Dim intPollLoopCounter	'Holds the poll loop counter
Dim strPollQuestion	'Holds the poll question
Dim blnMultipleVotes	'Set to true if multiple votes are allowed
Dim blnPollNoReply	'Set to true if this is a no reply poll
Dim strPollChoice	'Holds the poll choice
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



'Read in the message ID number to be deleted
lngTopicID = LngC(Request("TID"))


'If the person is not an admin or a moderator then send them away
If lngTopicID = "" Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
End If


'Check the form ID to prevent XCSRF
Call checkFormID(Request.QueryString("XID"))



'Initliase the SQL query to get the topic details from the database
strSQL = "SELECT " & strDbTable & "Topic.Forum_ID, " & strDbTable & "Topic.Poll_ID, " & strDbTable & "Topic.Subject " & _
"FROM " & strDbTable & "Topic" & strRowLock & " " & _
"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & ";"


'Query the database
rsCommon.Open strSQL, adoCon

'If there is a record returened read in the forum ID
If NOT rsCommon.EOF Then
	intForumID = CInt(rsCommon("Forum_ID"))
	lngPollID = CLng(rsCommon("Poll_ID"))
	strSubject = rsCommon("Subject")
End If





'Call the moderator function and see if the user is a moderator
If blnAdmin = False Then blnModerator = isModerator(intForumID, intGroupID)



'Check to make sure the user is deleting the post is a moderator or the forum adminstrator
If (blnAdmin = True OR blnModerator = True) AND lngPollID > 0 Then
	
	'Update the poll id with 0
	strSQL = "UPDATE " & strDbTable & "Topic SET Poll_ID = 0 WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & ";"
	
	'Write to database
	adoCon.Execute(strSQL)
	
	'Delete the Poll choices
	strSQL = "DELETE FROM " & strDbTable & "PollChoice " & strRowLock & " WHERE " & strDbTable & "PollChoice.Poll_ID="  & lngPollID & ";"

	'Write to database
	adoCon.Execute(strSQL)
	
	'Delete the Poll Votes 
	strSQL = "DELETE FROM " & strDbTable & "PollVote " & strRowLock & " WHERE " & strDbTable & "PollVote.Poll_ID=" & lngPollID & ";" 
			
	'Write to database 
	adoCon.Execute(strSQL)

	'Delete the Poll
	strSQL = "DELETE FROM " & strDbTable & "Poll " & strRowLock & " WHERE " & strDbTable & "Poll.Poll_ID="  & lngPollID & ";"

	'Write to database
	adoCon.Execute(strSQL)
	
	'If loging is enabled write to log file
	If blnLoggingEnabled AND (blnDeletePostLogging OR (blnModeratorLogging AND (blnAdmin OR blnModerator))) Then Call logAction(strLoggedInUsername, "Deleted Poll in '" & decodeString(strSubject) & "' PollID - " & lngPollID & " - TopicID " & lngTopicID)
End If


'Clean up
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<script language="JavaScript">
	window.opener.location.href = 'forum_posts.asp?TID=<% = lngTopicID %><% = strQsSID2 %>'
	window.close();
</script>
</head>
</html>