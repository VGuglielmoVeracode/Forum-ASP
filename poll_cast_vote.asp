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
Response.Buffer = true


'Declare variables
Dim lngTopicID			'Holds the topic number
Dim lngPollID			'Holds the poll ID
Dim lngPollVoteChoice		'Holds the users poll choice they are voting for
Dim blnForumLocked		'Make sure the forum hasn't been locked
Dim lngTotalChoiceVote		'Holds the number of votes the poll choice has received
Dim blnMultipleVotes		'set to true if multiple votes are allowed
Dim lngLastVoteUserID		'Holds the IP address of the voter
Dim blnAlreadyVoted		'Set to true if the user has already voted
Dim intResponseNum		'Holds the response number if there is one


'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)

End If



'Check the user has come to the file from a poll vote page if not send them to teh forum homepage
If Request.Form("PID") = "" OR Request.Form("TID") = "" Then Response.Redirect("default.asp" & strQsSID1)


'Initlise variables
blnForumLocked = True
blnAlreadyVoted = False


'Read in the details form the poll form
intForumID = IntC(Request.Form("FID"))
lngTopicID = LngC(Request.Form("TID"))
lngPollID = LngC(Request.Form("PID"))
lngPollVoteChoice = LngC(Request.Form("voteChoice"))



'Check the user is allowed to vote in this forum


'Read in the forum permssions from the database
Call forumPermissions(intForumID, intGroupID)

'Initalise the strSQL variable with an SQL statement to query the database (just to get if the forum is locked!!)
strSQL = "SELECT " & strDbTable & "Forum.Locked " & _
"FROM " & strDbTable & "Forum" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Forum.Forum_ID = " & intForumID & ";"

'Query the database
rsCommon.Open strSQL, adoCon


If NOT rsCommon.EOF Then

	'Read in wether the forum is locked or not
	blnForumLocked = CBool(rsCommon("Locked"))

End If

'Close the recordset
rsCommon.Close



'If the forum isn't locked and the user has the right to vote then let them vote
If blnForumLocked = False AND blnVote AND lngPollVoteChoice <> "" AND lngPollVoteChoice > 0 AND blnBanned = False Then


	'First check to see if multiple votes are allowed and if the user has voted before

	'Initlise the SQL query
	strSQL = "SELECT " & strDbTable & "Poll.Multiple_votes " & _
	"FROM " & strDbTable & "Poll" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Poll.Poll_ID = " & lngPollID & ";"

	'Query the database
	rsCommon.Open strSQL, adoCon

	'Check users is not making mutiple votes
	If NOT rsCommon.EOF Then

		'Read in if multiple votes are allowed
		blnMultipleVotes = CBool(rsCommon("Multiple_votes"))
		
		'Close the recordset
		rsCommon.Close


		'If multiple votes are not allowed check the last ID of the last voter
		If blnMultipleVotes = False Then

			'Check the user has not already voted by reading in a cookie from there system
			'Read in the Poll ID number of the last poll the user has voted in
			If IntC(getCookie("pID", "PID" & lngPollID)) = lngPollID OR getSessionItem("PID" & lngPollID) = lngPollID Then blnAlreadyVoted = True


			'If the user hasn't already voted then save their ID and move on to save their vote
			If blnAlreadyVoted = False Then
				
				'Check the database to see if the user has voted
				strSQL = "SELECT " & strDbTable & "PollVote.* " & _
				"FROM " & strDbTable & "PollVote" & strRowLock & " " & _
				"WHERE " & strDbTable & "PollVote.Poll_ID = " & lngPollID & " AND " & strDbTable & "PollVote.Author_ID = " & lngLoggedInUserID & ";"

				'Set the cursor type property of the record set to Forward Only
				rsCommon.CursorType = 0
		
				'Set the Lock Type for the records so that the record set is only locked when it is updated
				rsCommon.LockType = 3
				
				'Query the database
				rsCommon.Open strSQL, adoCon
				
				'If a record is returned then the user has voted so set blnAlreadyVoted to true
				If NOT rsCommon.EOF Then
					
					blnAlreadyVoted = True
				
				
				'Else the user has not voted so save there ID to database and a cookie and move on to save vote
				Else
					'Don't save user ID if this is a Guest, otherwise only 1 guest can vote
					If intGroupID <> 2 Then
						
						'Use ADO to update database as we already have a query running
						rsCommon.AddNew
						rsCommon.Fields("Poll_ID") = lngPollID
						rsCommon.Fields("Author_ID") = lngLoggedInUserID
						rsCommon.Update
					End If				
				End If
				
				'Save to a cookie as well
				'Write a cookie with the Poll ID number so the user cannot keep voting on this poll
				If blnEuCookieLaw = False OR blnGuest = False Then  Call setCookie("pID", "PID" & lngPollID, lngPollID, True)
				
				'Also save to app session
				Call saveSessionItem("PID" & lngPollID, lngPollID)
				
				'Close the recordset
				rsCommon.Close
			End If
		End If
	Else
	
		'Close the recordset
		rsCommon.Close
	End If

	


	'If the already voted boolean is not set then save the vote
	If blnAlreadyVoted = False Then


		'Save the voters choice

		'Initlise the SQL query
		strSQL = "SELECT " & strDbTable & "PollChoice.Choice_ID, " & strDbTable & "PollChoice.Votes " & _
		"FROM " & strDbTable & "PollChoice" & strRowLock & " " & _
		"WHERE " & strDbTable & "PollChoice.Choice_ID = " & lngPollVoteChoice & ";"

		'Set the cursor type property of the record set to Forward Only
		rsCommon.CursorType = 0

		'Set the Lock Type for the records so that the record set is only locked when it is updated
		rsCommon.LockType = 3

		'Query the database
		rsCommon.Open strSQL, adoCon

		'If a record is returned add 1 to it
		If NOT rsCommon.EOF Then

			'Read in the Poll Chioce Votes from rs
			lngTotalChoiceVote = CLng(rsCommon("Votes"))

			'Increment by 1
			lngTotalChoiceVote = lngTotalChoiceVote + 1

			'Update recordset
			rsCommon.Fields("Votes") = lngTotalChoiceVote

			'Update the database with the new poll choices
			rsCommon.Update
			
			'Set the error number to 1 for no error
			intResponseNum = 1
		End If

		'Close the recordset
		rsCommon.Close
	End If
End If

'Celan up
Call closeDatabase()

'Set up the return error number
If lngPollVoteChoice = 0 Then intResponseNum = 2
If blnAlreadyVoted = True Then  intResponseNum = 3


'Go back to the forum posts page
Response.Redirect("forum_posts.asp?TID=" & lngTopicID & "&PN=" & Request.Form("PN") & "&RN=" & intResponseNum & strQsSID3)
%>