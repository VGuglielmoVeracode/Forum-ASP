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





'Declare variables
Dim strPollQuestion		'Holds the poll question
Dim intPollChoiceNumber		'Holds the poll choice number
Dim strPollChoice		'Holds the poll choice
Dim lngPollChoiceVotes		'Holds the choice number of votes
Dim lngTotalPollVotes		'Holds the total number of votes
Dim dblPollVotePercentage	'Holds the vote percentage for the vote choice
Dim blnAlreadyVoted		'Set to true if the user has already voted
Dim blnMultipleVotes		'set to true if multiple votes are allowed
Dim sarryPoll			'Array to hold the poll recordset
Dim intPollCurrentRecord	'Hold the current postion in array


'Initlise variables
blnAlreadyVoted = False
intPollCurrentRecord = 0


'Get the poll from the database

'Initalise the strSQL variable with an SQL statement to query the database get the thread details
strSQL = "SELECT  " & strDbTable & "Poll.Poll_question, " & strDbTable & "Poll.Multiple_votes, " & strDbTable & "Poll.Reply, " & strDbTable & "PollChoice.Choice_ID, " & strDbTable & "PollChoice.Choice, " & strDbTable & "PollChoice.Votes " & _
"FROM " & strDbTable & "Poll" & strDBNoLock & ", " & strDbTable & "PollChoice" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Poll.Poll_ID=" & strDbTable & "PollChoice.Poll_ID " & _
	"AND " & strDbTable & "Poll.Poll_ID=" & lngPollID & ";"



'Query the database
rsCommon.Open strSQL, adoCon

'If no record release rs
If rsCommon.EOF Then
	
	'Close recordsets
	rsCommon.Close

'If there is a poll then display it
Else	
	'Read in the row from the db using getrows for better performance
	sarryPoll = rsCommon.GetRows()
	
	'Close recordsets
	rsCommon.Close
	
	'Initilise total votes
	lngTotalPollVotes = 0

	'Read in the poll question
	strPollQuestion = sarryPoll(0,0)
	
	'See if multiple votes are allowed
	blnMultipleVotes = CBool(sarryPoll(1,0))
	
	'See if this is a poll only
	blnPollNoReply = CBool(sarryPoll(2,0))

	
	'Loop through and get the total number of votes
	Do While intPollCurrentRecord < (UBound(sarryPoll,2) + 1)

		'Get the total number of votes
		lngTotalPollVotes = lngTotalPollVotes + CLng(sarryPoll(5,intPollCurrentRecord))

		'Move to the next array position
        	intPollCurrentRecord = intPollCurrentRecord + 1
        Loop
        
        'Reset array position
        intPollCurrentRecord = 0

	'If multiple votes are not allowed see if the user has voted before
	If blnMultipleVotes = False Then

		'Check the user has not already voted by reading in a cookie from there system
		'Read in the Poll ID number of the last poll the user has voted in
		If IntC(getCookie("pID", "PID" & lngPollID)) = lngPollID OR IntC(getSessionItem("PID" & lngPollID)) = lngPollID Then blnAlreadyVoted = True
	End If
	

%><!-- Start Poll -->
<form name="frmPoll" id="frmPoll" method="post" action="poll_cast_vote.asp<% = strQsSID1 %>">
<table class="tableBorder" align="center" cellspacing="1" cellpadding="10">
 <tr class="tableLedger">
  <td colspan="4"><% = strTxtPollQuestion %>: <% = strPollQuestion %></td>
 </tr>
 <tr class="tableTopRow"><%

	'Display the vote choice slection column if the user CAN vote in this poll
	If blnVote = True AND blnForumLocked = False AND blnTopicLocked = False AND blnActiveMember = True AND blnAlreadyVoted = False Then

	%>
  <td width="3%"><% = strTxtVote %></td><%
	End If
%>
  <td width="47%" nowrap><% = strTxtPollChoice %></td>
  <td width="6%" align="center" nowrap><% = strTxtVotes %></td>
  <td width="47%"><% = strTxtPollStatistics %></td>
 </tr><%

 	'Loop through the Poll Choices
 	'Do....While Loop to loop through the recorset to display the Poll Choices
	Do While intPollCurrentRecord < (UBound(sarryPoll,2) + 1)

 		'Read in the poll details
 		intPollChoiceNumber = Cint(sarryPoll(3,intPollCurrentRecord))
 		strPollChoice = sarryPoll(4,intPollCurrentRecord)
 		lngPollChoiceVotes = CLng(sarryPoll(5,intPollCurrentRecord))

		'If there are no votes yet then format the percent by 0 otherwise an overflow error will happen
		If lngTotalPollVotes = 0 Then
			dblPollVotePercentage = FormatPercent(0, 2)

		'Else read in the the percentage of votes cast for the vote choice
		Else
			dblPollVotePercentage = FormatPercent((lngPollChoiceVotes / lngTotalPollVotes), 2)
		End If

        %>			
 <tr <%
 		'Create row class for alternative colours etc.
 		If (intPollCurrentRecord MOD 2 = 0 ) Then Response.Write("class=""evenTableRow"">") Else Response.Write("class=""oddTableRow"">") 

		'Display the vote radio buttons if the user CAN vote in this poll
		If blnVote AND blnForumLocked = False AND blnTopicLocked = False AND blnActiveMember AND blnAlreadyVoted = False Then

	%>
  <td align="center"><input type="radio" name="voteChoice" value="<% = intPollChoiceNumber %>" id="P<% = intPollChoiceNumber %>"></td><%

        	End If

        %>
  <td><label for="P<% = intPollChoiceNumber %>"><% = strPollChoice %></label></td>
  <td align="center"><% = lngPollChoiceVotes %></td>
  <td class="smText" nowrap><img src="<% = strImagePath %>bar_graph_image.gif" width="<% = CInt(Replace(CStr(dblPollVotePercentage), "%", "", 1, -1, 1)) * 2 %>" height="11" align="middle"> [<% = dblPollVotePercentage %>]</td>
 </tr><%

        	'Move to the next record
        	intPollCurrentRecord = intPollCurrentRecord + 1
        Loop

        %>
 <tr align="center" class="tableBottomRow">
  <td colspan="4"><%

	'Display either text msg if the user can NOT vote or a button if they can
	
	'If the forum is locked display a locked forum meesage
	If blnForumLocked = True OR  blnTopicLocked = True Then
	
		Response.Write(strTxtThisTopicIsClosedNoNewVotesAccepted)
	
	'Else the user can not vote or they are not an active member of the forum
	ElseIf blnActiveMember = False OR blnVote = False OR blnBanned Then
	
		Response.Write(strsTxYouCanNotNotVoteInThisPoll)
	
	'Else the user has already voted in this poll and multiple votes are not permitted
	ElseIf blnAlreadyVoted = True Then
	
		Response.Write(strTxtYouHaveAlreadyVotedInThisPoll)
	
	'Else display vote button
	Else
%>
   <input type="hidden" name="PID" id="PID" value="<% = lngPollID %>" />
   <input type="hidden" name="TID" id="TID" value="<% = lngTopicID %>" />
   <input type="hidden" name="FID" id="FID" value="<% = intForumID %>" />
   <input type="hidden" name="PN" id="PN" value="<% = intRecordPositionPageNum %>" />
   <input type="submit" name="cateVote" id="castVote" value="<% = strTxtCastMyVote %>" /><%
	End If

%></td>
 </tr>
</table>
</form>
<br />
<!-- End Poll --><%

End If


'Display a msg letting the user know if there vote has been cast or not
Select Case Request.QueryString("RN")
	Case "1"
		Response.Write("<script  language=""JavaScript"">" & _
		"alert('" & strTxtThankYouForCastingYourVote & "');" & _
		"</script>")
	Case "2"
		Response.Write("<script  language=""JavaScript"">" & _
		"alert('" & strTxtYouDidNotSelectAChoiceForYourVote & "');" & _
		"</script>")
	Case "3"
		Response.Write("<script  language=""JavaScript"">" & _
		"alert('" & strTxtYouHaveAlreadyVotedInThisPoll & "');" & _
		"</script>")
End Select
%>