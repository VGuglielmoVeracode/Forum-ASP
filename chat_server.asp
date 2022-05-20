<% @ Language=VBScript %>
<% Option Explicit %>
<!-- #include file="database/database_connection.asp" -->
<!-- #include file="functions/functions_common.asp" -->
<!-- #include file="functions/functions_filters.asp" -->
<!--#include file="language_files/chat_room_language_file_inc.asp" -->
<!--#include file="includes/emoticons_inc.asp" -->
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

'Set the content type
Response.ContentType = "text/javascript"


Dim strSessionID
Dim strSessionData
Dim saryUserSessionData
Dim strUsername
Dim strErrorMessage

Dim lngAuthorID

Dim strWriteChat
Dim sarryChat
Dim intArrayPass
Dim intChatArraySize
Dim intChatArrayPass
Dim lngChatPointer


Dim saryChatMembers
Dim blnMemberFound
Dim blnNewChatMember
Dim intMemberArraySize
Dim intMemberArrayPass
Dim lngMemberPointer
Dim strPresentUserActivity
Dim intLastArrayPostionPointer
Dim dtmChatStart



Const strChatVarPrefix = "ChatWWF10-20110901" 'Chat application prefix. If running mutiple forums on one site will need to change this
Const intMaxChatArraySize = 15 'default is 16 (array positions start at 0)
Const intSessionTimeout = 15 'session timeout 15 secounds
Const blnEmoticons = True 'Emoticons
Const intChatHistoryTime = 2 'The amount of chat history to show to new members in minutes



'variables
blnMemberFound = False
blnNewChatMember = False
lngAuthorID = 2
intLastArrayPostionPointer = 0



'Read in the members session ID
strSessionID = Trim(Request("SID"))

'Clean up session ID
strSessionID = formatSQLInput(strSessionID)



'Read in if the member is typing
If Request.Form("t") = "true" Then
	strPresentUserActivity = "typing"
ElseIf Request.Form("a") = "false" Then
	strPresentUserActivity = "na"
Else
	strPresentUserActivity = ""
End If





'Get chat messages from aplication
Application.Lock



'Read in the message pointer position
If isNumeric(Request.QueryString("p")) Then
	lngChatPointer = CLng(Request.QueryString("p"))
Else
	lngChatPointer = 0
End If


'Get the aplication chat array
If isArray(Application(strChatVarPrefix & "sarryChat")) Then 
	sarryChat = Application(strChatVarPrefix & "sarryChat")
Else
	ReDim sarryChat(5, -1)
End If


'Get the aplication member array
If isArray(Application(strChatVarPrefix & "saryChatMembers")) Then 
	saryChatMembers = Application(strChatVarPrefix & "saryChatMembers")
Else
	ReDim saryChatMembers(5, -1)
End If


'Get the size of the members array
intMemberArraySize = CInt(UBound(saryChatMembers, 2))


'Look through the member array to see if the member is lsited, if they are get/update details
If intMemberArraySize >= 0 Then
	
	'0 = Session_ID
	'1 = Author_ID
	'2 = Username
	'3 = Last Active Time
	'4 = Typing/Active/Not Active
	'5 = Chat Start
	
	'Loop through the array and see if member is in it
	For intArrayPass = 0 TO intMemberArraySize
		
		'If the member is found in the member array read in their details
		If strSessionID = saryChatMembers(0, intArrayPass) Then
			
			'Get 
			blnMemberFound = True
			lngAuthorID = CLng(saryChatMembers(1, intArrayPass))
			strUsername = saryChatMembers(2, intArrayPass)
			dtmChatStart = CDate(saryChatMembers(5, intArrayPass))
			'Write
			saryChatMembers(3, intArrayPass) = Now()
			saryChatMembers(4, intArrayPass) = strPresentUserActivity
			
			'Exit user check loop
			Exit For
		End If
	Next
End If






'Remove inactive members from the members array

'Get the last array postion usng array size
intMemberArraySize = CInt(UBound(saryChatMembers, 2))
intLastArrayPostionPointer = intMemberArraySize

'Loop through and update the members array
For intArrayPass = 0 TO intMemberArraySize
	
	'Check the last cactive time and remove if older than 20 secounds
	If CDate(saryChatMembers(3, intArrayPass)) < CDate(DateAdd("s", -intSessionTimeout, Now())) Then
		
		'Check that the array postion pointer is not for an outdated session (AND part for error handling as don't want intLastArrayPostionPointer to be less than 0)
		If CDate(saryChatMembers(3, intLastArrayPostionPointer)) < DateAdd("n", -intSessionTimeout, Now()) AND intLastArrayPostionPointer > 0 Then intLastArrayPostionPointer = intLastArrayPostionPointer - 1
	
		'Swap this array postion with the last in the array
		saryChatMembers(0, intArrayPass) = saryChatMembers(0, intLastArrayPostionPointer)
		saryChatMembers(1, intArrayPass) = saryChatMembers(1, intLastArrayPostionPointer)
		saryChatMembers(2, intArrayPass) = saryChatMembers(2, intLastArrayPostionPointer)
		saryChatMembers(3, intArrayPass) = saryChatMembers(3, intLastArrayPostionPointer)
		saryChatMembers(4, intArrayPass) = saryChatMembers(4, intLastArrayPostionPointer)
		saryChatMembers(5, intArrayPass) = saryChatMembers(5, intLastArrayPostionPointer)
				
		'Decrement the last array pointer
		If intLastArrayPostionPointer > 0 Then intLastArrayPostionPointer = intLastArrayPostionPointer - 1
			
		'create info message that user has left
		Call WriteChatArray("", 0, "*** " & saryChatMembers(2, intArrayPass) & " " & strTxtHasLeftTheChatRoom & " ***", "info")
		
	End If
Next

'Removed old member array postions
If intMemberArraySize > intLastArrayPostionPointer Then ReDim Preserve saryChatMembers(5, intLastArrayPostionPointer)







'If the members NOT found in the array read in their details from forum database and add to member chat array
If blnMemberFound = False Then
	
	'Open a db connection
	Call openDatabase(strCon)
	
	'See if there is a session in the db for this member
	strSQL = "SELECT " & strDbTable & "Session.Session_ID, " & strDbTable & "Session.Session_data " & _
	"FROM " & strDbTable & "Session" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Session.Session_ID = '" & strSessionID & "';"
	
	'Set error trapping
	On Error Resume Next
	
	'Get recordset
	rsCommon.Open strSQL, adoCon
				
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then Response.Write("alert('An error has occurred while executing SQL query on database.\n\nError Code: get_session_data\nFile Name:chat_server.asp')")
			
	'Disable error trapping
	On Error goto 0
	
	'If there is a session returned read in the author ID from it
	If NOT rsCommon.EOF Then
	
		strSessionData = rsCommon("Session_data")
		
		'Split the session data up into an array
		saryUserSessionData = Split(strSessionData, ";")
	
		'Loop through array to get the author ID
		For intArrayPass = 0 to UBound(saryUserSessionData)
			If InStr(saryUserSessionData(intArrayPass), "AID=") Then
				'Read in the author ID
				lngAuthorID = CLng(Replace(saryUserSessionData(intArrayPass), "AID=", "", 1, -1, 1))
			End If
		Next
		
		'Close rs
		rsCommon.Close
		
		'Response.Write("alert('AID=" & lngAuthorID & "');")
	
	
	'Else no session returned so end
	Else
	
		'close db connection
		rsCommon.Close
		Call closeDatabase()
		
		'Send alert message to user that their session is not valid
		Response.Write(vbCrLf & "chat(1,0,'warn','System','*** " & strTxtSessionDropedPleaseRefreshPage & " ***');")
		
		'End server response
		Response.Flush
		Response.End
	
	End If
	
	
	
	'Get the members details from the database
	strSQL = "SELECT " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.Group_ID,  " & strDbTable & "Author.Active, " & strDbTable & "Author.Banned " & _
	"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Author.Author_ID = " & lngAuthorID & ";"
		
	'Set error trapping
	On Error Resume Next
		
	'Get recordset
	rsCommon.Open strSQL, adoCon
					
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then Response.Write("alert('An error has occurred while executing SQL query on database.\n\nError Code: get_member_data\nFile Name:chat_server.asp')")
				
	'Disable error trapping
	On Error goto 0
		
	'If the member is returned read in the data
	If NOT rsCommon.EOF Then
		
		'If banned
		If CBool(rsCommon("Banned")) AND lngAuthorID > 1 Then
			strErrorMessage = strTxtYouAreNotPermittedToUseThisChatRoom & ".\n\n" & strTxtYourAccountIsSuspended & "."
		'If not active
		ElseIf CBool(rsCommon("Active")) = False AND lngAuthorID > 1 Then
			strErrorMessage = strTxtYouAreNotPermittedToUseThisChatRoom & ".\n\n" & strTxtYourAccountIsNotActive & "."
		'If not logged in
		ElseIf (CLng(rsCommon("Group_ID")) = 2 OR lngAuthorID = 2) AND lngAuthorID > 1 Then
			strErrorMessage = strTxtYouAreNotPermittedToUseThisChatRoom & ". \n\n" & strTxtYouAreNotLoggedIn & "."
		'Else get the members name
		Else	
			strUsername = rsCommon("Username")
		End If
			
	End If
	
	'close db connection
	rsCommon.Close
	Call closeDatabase()
	
	
	'If there is an error message kick the user
	If strErrorMessage <> "" Then
		
		'Send alert message to user that their session is not valid
		Response.Write(vbCrLf & "chat(1,0,'warn','System','*** " & strErrorMessage & " ***');")
		
		'End server response
		Response.Flush
		Response.End
		
	End If
	



	'Member is OK to chat so add them to the member chat array
	
	'0 = Session_ID
	'1 = Author_ID
	'2 = Username
	'3 = Last Active Time
	'4 = Typing/Active/Not Active
	
	'Get the array size
	intMemberArraySize = CInt(UBound(saryChatMembers, 2)) + 1
	
	'Set the time of the members chat starting (set to 1 secound back to display enter message and the last couple of secound of chat)
	dtmChatStart = DateAdd("m", -intChatHistoryTime, Now())
	
	'Grow the array
	ReDim Preserve saryChatMembers(5, intMemberArraySize)
	
	'Set the rest of the array
	saryChatMembers(0, intMemberArraySize) = strSessionID
	saryChatMembers(1, intMemberArraySize) = lngAuthorID
	saryChatMembers(2, intMemberArraySize) = strUsername
	saryChatMembers(3, intMemberArraySize) = Now()
	saryChatMembers(4, intMemberArraySize) = strPresentUserActivity	
	saryChatMembers(5, intMemberArraySize) = dtmChatStart
	
	'Create info message for chat room
	Call WriteChatArray("", 0, "*** " & strUsername & " " & strTxtHasEnteredTheChatRoom & " ***", "info")
	
	'Set the new chat member to true so that we do not show chat history
	blnNewChatMember = True
	
End If







'Write  the member array
intMemberArraySize = CInt(UBound(saryChatMembers, 2))

'Loop through and display the array
For intArrayPass = 0 TO intMemberArraySize

	'0 = Session_ID
	'1 = Author_ID
	'2 = Username
	'3 = Last Active Time
	'4 = Typing/Active/Not Active
	
	'Write out member array
	Response.Write(vbCrLf & "online(" & _
		(intArrayPass + 1) & "," & _
		saryChatMembers(1, intArrayPass) & "," & _
		"'" & saryChatMembers(3, intArrayPass) & "'," & _
		"'" & saryChatMembers(2, intArrayPass) & "'," & _
		"'" & saryChatMembers(4, intArrayPass) & "')")
Next







'Read in posted message (trim to 250 chars as the message should only be this length)
strWriteChat = Trim(Mid(Request("writeMsg"), 1, 250))

'If this is a message being posted then process
If strWriteChat <> "" Then Call WriteChatArray(strUsername, lngAuthorID, strWriteChat, "m")

	
'Get the chat array size
intChatArraySize = CInt(UBound(sarryChat, 2))	


'Loop thgrough the and display the chat array
For intArrayPass = 0 TO intChatArraySize

	'0 = Message ID
	'1 = Username
	'2 = Author_ID
	'3 = Date
	'4 = Message
	'5 = Message Type
	
	'Response.Write(vbcrlf & dtmChatStart & " < " & CDate(sarryChat(3, intArrayPass)))
	
	'Output chat room messages
	
	'Check if the users chat start time is less than the message posted
	If dtmChatStart < CDate(sarryChat(3, intArrayPass)) Then
	
		'To reduce bandwidth only output messages that have not already been oiutput to the user
		If sarryChat(0, intArrayPass) > lngChatPointer Then
		
			Response.Write(vbCrLf & "chat(" & _
					sarryChat(0, intArrayPass) & "," & _
					sarryChat(2, intArrayPass) & "," & _
					"'" & sarryChat(5, intArrayPass) & "'," & _
					"'" & sarryChat(1, intArrayPass) & "'," & _
					"'" & sarryChat(4, intArrayPass) & "')")
		End If
	End If
	
Next

'Update application variables
If UBound(sarryChat, 2) => 0 Then Application(strChatVarPrefix & "sarryChat") = sarryChat
If UBound(saryChatMembers, 2) => 0 Then Application(strChatVarPrefix & "saryChatMembers") = saryChatMembers

Application.Unlock








'Function to write to the chat array	
Function WriteChatArray(strUsername, lngAuthorID, strMessage, strMsgType)

	'Escape JavaScript
	strMessage = Replace(strMessage, "'", "\'", 1, -1, 1)
	strMessage = Replace(strMessage, """", "\""", 1, -1, 1)
	
	'HTML encode message to encode non ASCII characters (adds support for non English characters)
	'This is done now after tags are stripped and before BBcodes to avoid formatting errors
	strMessage = Server.HTMLEncode(strMessage)
	
	'Format BBcodes
	strMessage = FormatBBCodes(strMessage)
	
	'Filter smut
	strMessage = badWordFilter(strMessage)
	

	'Get the array size
	intChatArraySize = CInt(UBound(sarryChat, 2))
	
	'If the array size is less than max size then just add the message to the next array element
	If intChatArraySize < intMaxChatArraySize Then

		'increase message array to add new message
		ReDim Preserve sarryChat(5, intChatArraySize + 1)
		intChatArraySize = CInt(UBound(sarryChat, 2))
	
	'Else the array size is at it's maximum so shift everything up
	Else	
		'Loop through the array and shift up one
		For intChatArrayPass = 0 To intChatArraySize - 1
		
			'Swap this array postion with the last in the array
			sarryChat(0, intChatArrayPass) = sarryChat(0, intChatArrayPass + 1)
			sarryChat(1, intChatArrayPass) = sarryChat(1, intChatArrayPass + 1)
			sarryChat(2, intChatArrayPass) = sarryChat(2, intChatArrayPass + 1)
			sarryChat(3, intChatArrayPass) = sarryChat(3, intChatArrayPass + 1)
			sarryChat(4, intChatArrayPass) = sarryChat(4, intChatArrayPass + 1)
			sarryChat(5, intChatArrayPass) = sarryChat(5, intChatArrayPass + 1)
		Next
	
	End If
	
	'0 = Message ID
	'1 = Username
	'2 = Author_ID
	'3 = Date
	'4 = Message
	'5 = Message type
	
	'If array is empty then set the message ID to be the current unix date stamp (as it is secounds from 01/01/1970, will always be a higher number
	If intChatArraySize = 0 Then
		sarryChat(0, 0) = DateDiff("s", "01/01/1970 00:00:00", Now())
	'Else get the message ID by counting from the last array position
	Else
		sarryChat(0, intChatArraySize) = sarryChat(0, intChatArraySize-1) + 1
	End If
	'Set the rest of the array
	sarryChat(1, intChatArraySize) = Server.HTMLEncode(strUsername)
	sarryChat(2, intChatArraySize) = lngAuthorID
	sarryChat(3, intChatArraySize) = Now()
	sarryChat(4, intChatArraySize) = strMessage
	sarryChat(5, intChatArraySize) = strMsgType		
End Function





'Format Forum Codes Function to covert forum codes to HTML
Private Function FormatBBCodes(ByVal strMessage)

	'Smilies
	strMessage = Replace(strMessage, ":)", "<img src=""smileys/smiley1.gif"" align=""absmiddle"" />", 1, -1, 1)
	strMessage = Replace(strMessage, ":P", "<img src=""smileys/smiley17.gif"" align=""absmiddle"" />", 1, -1, 1)
	strMessage = Replace(strMessage, ":D", "<img src=""smileys/smiley4.gif"" align=""absmiddle"" />", 1, -1, 1)
	strMessage = Replace(strMessage, ";)", "<img src=""smileys/smiley2.gif"" align=""absmiddle"" />", 1, -1, 1)
	strMessage = Replace(strMessage, "LOL", "<img src=""smileys/smiley36.gif"" align=""absmiddle"" />", 1, -1, 0)
	strMessage = Replace(strMessage, ":$", "<img src=""smileys/smiley9.gif"" align=""absmiddle"" />", 1, -1, 1)
	strMessage = Replace(strMessage, ":s", "<img src=""smileys/smiley5.gif"" align=""absmiddle"" />", 1, -1, 1)
	strMessage = Replace(strMessage, ":x", "<img src=""smileys/smiley7.gif"" align=""absmiddle"" />", 1, -1, 1)
	strMessage = Replace(strMessage, "8(", "<img src=""smileys/smiley18.gif"" align=""absmiddle"" />", 1, -1, 1)
	strMessage = Replace(strMessage, ":o", "<img src=""smileys/smiley3.gif"" align=""absmiddle"" />", 1, -1, 1)
	strMessage = Replace(strMessage, "|)", "<img src=""smileys/smiley12.gif"" align=""absmiddle"" />", 1, -1, 1)
	strMessage = Replace(strMessage, ":(", "<img src=""smileys/smiley6.gif"" align=""absmiddle"" />", 1, -1, 1)
	strMessage = Replace(strMessage, "8D", "<img src=""smileys/smiley16.gif"" align=""absmiddle"" />", 1, -1, 1)
	strMessage = Replace(strMessage, ":|", "<img src=""smileys/smiley22.gif"" align=""absmiddle"" />", 1, -1, 1)
	strMessage = Replace(strMessage, ":-)", "<img src=""smileys/smiley1.gif"" align=""absmiddle"" />", 1, -1, 1)

	'Format bbcodes (done this way so don't get left with open tags)
	If InStr(1, strMessage, "[B]", 1) AND InStr(1, strMessage, "[/B]", 1) Then
		strMessage = Replace(strMessage, "[B]", "<strong>", 1, 1, 1)
		strMessage = Replace(strMessage, "[/B]", "</strong>", 1, 1, 1)
	End If
	
	If InStr(1, strMessage, "[I]", 1) AND InStr(1, strMessage, "[/I]", 1) Then
		strMessage = Replace(strMessage, "[I]", "<em>", 1, 1, 1)
		strMessage = Replace(strMessage, "[/I]", "</em>", 1, 1, 1)
	End If
	
	If InStr(1, strMessage, "[U]", 1) AND InStr(1, strMessage, "[/U]", 1) Then
		strMessage = Replace(strMessage, "[U]", "<u>", 1, 1, 1)
		strMessage = Replace(strMessage, "[/U]", "</u>", 1, 1, 1)
	End If
		 
	If InStr(1, strMessage, "[B]", 1) Then strMessage = Replace(strMessage, "[B]", "<strong>", 1, 1, 1) & "</strong>"
	If InStr(1, strMessage, "[I]", 1) Then strMessage = Replace(strMessage, "[I]", "<em>", 1, 1, 1) & "</em>"
	If InStr(1, strMessage, "[U]", 1) Then strMessage = Replace(strMessage, "[U]", "<u>", 1, 1, 1) & "</u>"
	
	
	
	'IRC commands
	strMessage = Replace(strMessage, "/me", "*" & Server.HTMLEncode(strUsername), 1, -1, 1)
	
	'Return the function
	FormatBBCodes = strMessage
End Function




'Bad word filter
'To reduce database calls the bad word filter is read into memory from the database
Private Function badWordFilter(ByVal strMessage)

	Dim sarryBadWordFilter
	Dim intBadWordLoop
	Dim strBadWord
	Dim strBadWordReplace
	Dim objRegExp
	
	'Get the bad word filter array
	If isArray(Application(strChatVarPrefix & "sarryBadWordFilter")) Then 
		
		sarryBadWordFilter = Application(strChatVarPrefix & "sarryBadWordFilter")
	
	
	'Else not in memory, so read in from db
	Else
	
		'Open a db connection
		Call openDatabase(strCon)
		
		'get the bad words from the bad word table
		strSQL = "SELECT " & strDbTable & "Smut.Smut, " & strDbTable & "Smut.Word_replace " & _
		"FROM " & strDbTable & "Smut" & strDBNoLock & ";"
		
		'Set error trapping
		On Error Resume Next
		
		'Get recordset
		rsCommon.Open strSQL, adoCon
					
		'If an error has occurred write an error to the page
		If Err.Number <> 0 Then Response.Write("alert('An error has occurred while executing SQL query on database.\n\nError Code: get_smut_data\nFile Name:chat_server.asp')")
				
		'Disable error trapping
		On Error goto 0
		
		'If we got data read it in
		If NOT rsCommon.EOF Then
			
			sarryBadWordFilter = rsCommon.GetRows()
		
		'Else create an array to keep code happy
		Else
			ReDim sarryBadWordFilter(2, -1)

		End If
		
		'close db connection
		rsCommon.Close
		Call closeDatabase()
		
		'Place the bad words into
		Application(strChatVarPrefix & "sarryBadWordFilter") = sarryBadWordFilter
	End If
	
	
	'Create regular experssions object
	Set objRegExp = New RegExp
	
	'If we have an array go through it and replace bad words
	For intBadWordLoop = 0 TO UBound(sarryBadWordFilter, 2)
	
		'Put the bad word into a string	for imporoved perfoamnce
		strBadWord = sarryBadWordFilter(0, intBadWordLoop)
		strBadWordReplace = sarryBadWordFilter(1, intBadWordLoop)
	
		'Tell the regular experssions object what to look for
		With objRegExp
			.Pattern = strBadWord
			.IgnoreCase = True
			.Global = True
		End With
		
		'Ignore errors, incase someone entered an incorrect bad word that breakes regular expressions
		On Error Resume Next
		
		'Replace the swear words with the words	in the database	the swear words
		strMessage = objRegExp.Replace(strMessage, strBadWordReplace)
		
		'Disable error trapping
		On Error goto 0
	
	
	
	Next
	
	'Distroy regular experssions object
	Set objRegExp = nothing
	
	'Return the function
	badWordFilter = strMessage
End Function


%>