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






Dim sarryPollChoice	'Holds the poll choices
Dim blnMultiVotes	'Holds in multi votes
Dim blnPollOnly		'Holds if poll only

'If this is a Poll Edit we need to read in the details from the database
If strMode = "editPoll" Then
	
	'Get the poll details from the database
	strSQL = "SELECT " & strDbTable & "Poll.Poll_question, " & strDbTable & "Poll.Multiple_votes, " & strDbTable & "Poll.Reply, " & strDbTable & "PollChoice.Choice " & _
	"FROM " & strDbTable & "Poll" & strDBNoLock & ", " & strDbTable & "PollChoice" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Poll.Poll_ID = " & strDbTable & "PollChoice.Poll_ID " & _
		"AND " & strDbTable & "Poll.Poll_ID=" & lngPollEditID & ";"
	
	'Query the database
	rsCommon.Open strSQL, adoCon 
	
	'Place the rs into an array
	If NOT rsCommon.EOF Then 
		
		'Place rs in array
		sarryPollChoice = rsCommon.GetRows()
		
		'Read in a few of the values
		blnMultiVotes = CBool(sarryPollChoice(1,0))
		blnPollOnly = CBool(sarryPollChoice(2,0))
		
		'Change the max number of poll choices to the same number as that entered
		intMaxPollChoices = UBound(sarryPollChoice,2) + 1
	End If
	
	'Close the recordset	
	rsCommon.Close
End If


%>	<tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
        </tr>
	<tr>
         <td align="right"><% = strTxtPollQuestion %>:</td>
         <td aligh="left"><input name="pollQuestion" id="pollQuestion" type="text" size="30" maxlength="70" value="<% If isArray(sarryPollChoice) Then Response.Write(sarryPollChoice(0,0)) %>" tabindex="11" /></td>
        </tr><%

'Loop around to display text boxes for the maximum amount of allowed poll questions
For intPollLoopCounter = 1 to intMaxPollChoices


	Response.Write(vbCrLf & "        <tr>" & _
        vbCrLf & "	 <td align=""right"">")

        'Display the poll choice text
        Response.Write(strTxtPollChoice & "&nbsp;")

	'Display poll number
	Response.Write(intPollLoopCounter)

	Response.Write(":</td>" & _
	vbCrLf & "	<td><input name=""choice" & intPollLoopCounter & """ id=""choice" & intPollLoopCounter & """ type=""text"" size=""30"" maxlength=""60""")
	'If we are editing a poll display the old poll choice
	If strMode = "editPoll" AND isArray(sarryPollChoice) Then	
		'Make sure we have not run out of choices to display, if not display poll choice
		If UBound(sarryPollChoice,2) >=(intPollLoopCounter-1) Then Response.Write(" value=""" & sarryPollChoice(3,intPollLoopCounter-1) & """")
	End If
	Response.Write(" tabindex=""" & 10 + intMaxPollChoices + 1 & """ /></td>" & _
	vbCrLf & "        </tr>")
Next

%>	<tr>
         <td align="right">&nbsp;</td>
         <td align="left">&nbsp;<input type="checkbox" name="multiVote" id="multiVote" value="True" <% If blnMultiVotes = true Then Response.Write(" checked") %>/><% = strTxtAllowMultipleVotes %></td>
        </tr>
        <tr>
         <td align="right">&nbsp;</td>
         <td align="left">&nbsp;<input type="checkbox" name="pollReply" id="pollReply" value="True" <% If blnPollOnly = true Then Response.Write(" checked") %>/><% = strTxtMakePollOnlyNoReplies %></td>
        </tr>
        <tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
        </tr>