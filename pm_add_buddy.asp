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

'Declare variables
Dim strUsername		'Holds the usrename of the new buddy
Dim strDescription	'Holds a short description of the buddy
Dim blnBlocked		'Set to true if the users is blocked from messaging
Dim intCode		'Return page code
Dim intErrorNum		'Holds the error number
Dim lngAuthorID		'Holds the authors user ID


blnSslEnabledPage = True


'Set the return page code
intCode = 1

'If Priavte messages are not on then send them away
If blnPrivateMessages = False Then 
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
End If


'If the user is not allowed then send them away
If intGroupID = 2 OR blnActiveMember = False OR blnBanned Then 
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If


'Check the session ID to stop spammers using the email form
Call checkFormID(Request.Form("formID"))



'Read in the details from the form
strUsername = Trim(Mid(Request.Form("username"), 1, 25))
strDescription = Trim(Mid(Request.Form("description"), 1, 30))
blnBlocked = BoolC(Request.Form("blocked"))


'Clean up user input
strUsername = formatSQLInput(strUsername)
strDescription = formatInput(strDescription)



'Check that the new buddy exsists
	
'Initalise the SQL string to query the database to see if the uername exists
strSQL = "SELECT " & strDbTable & "Author.Author_ID " & _
"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Author.Username = '" & strUsername & "';"

'Open the recordset
rsCommon.Open strSQL, adoCon



'If the user exsist check there not in the list and then add them
If NOT rsCommon.EOF Then
	
	'Get the author ID
	lngAuthorID = CLng(rsCommon("Author_ID"))
	
	'Close rs
	rsCommon.Close
		
	'Initalise the SQL string with a query to check to see if user is already in list
	strSQL = "SELECT " & strDbTable & "BuddyList.* " & _
	"FROM " & strDbTable & "BuddyList" & strRowLock & " " & _
	"WHERE " & strDbTable & "BuddyList.Buddy_ID = " & lngAuthorID & " " & _
		"AND " & strDbTable & "BuddyList.Author_ID = " & lngLoggedInUserID & ";"
	
	'Set the cursor type property of the record set to Forward Only
	rsCommon.CursorType = 0
	
	'Set the Lock Type for the records so that the record set is only locked when it is updated
	rsCommon.LockType = 3
	
	'Open the recordset
	rsCommon.Open strSQL, adoCon
	
	'If no record is returned the buddy is not already in the buddy list so eneter them
	If rsCommon.EOF Then
		
		'Add the new buddy
		rsCommon.AddNew
		rsCommon.Fields("Author_ID") = lngLoggedInUserID
		rsCommon.Fields("Buddy_ID") = lngAuthorID
		rsCommon.Fields("Description") = strDescription
		rsCommon.Fields("Block") = blnBlocked
		rsCommon.Update
		
		'Set the msg varaible to let the user know the buddy has been added
		intCode = 2
		
	'Else the buddy is alreay entered so set the msg varaiable to tell the user
	Else
		intErrorNum = 1
	End If

	'Close rs
	rsCommon.Close

Else
	'Close rs
	rsCommon.Close
	
	'Tell the next page to display an error msg as user is not found
	intErrorNum = 2
End If

'Clear up
Call closeDatabase()

'Remove anti SQL injection code
strUsername = Replace(strUsername, "''", "'", 1, -1, 1)

'Return to the page showing the threads
Response.Redirect("pm_buddy_list.asp?name=" & Server.URLEncode(strUsername) & "&desc=" & Server.URLEncode(strDescription) & "&code=" & intCode & "&ER=" & intErrorNum & strQsSID3)
%>