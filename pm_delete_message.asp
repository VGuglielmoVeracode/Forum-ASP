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


Dim laryMesageID 	'Holds the message id number
Dim blnOutbox




'If Priavte messages are not on then send them away
If blnPrivateMessages = False Then 
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
End If



'If the user is not allowed then send them away
If intGroupID = 2 OR blnActiveMember = False OR blnBanned OR lngLoggedInUserID = 0 Then 
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If



'Read in if this is the outbox or not
blnOutbox = BoolC(Request.Form("outbox"))



'Remove all PM's from the users inbox/outbox
If Request.Form("delAll") <> "" Then
	
	'Check the session ID to stop spammers using the email form
	Call checkFormID(Request.Form("formID"))
	
	'If this is the users Outbox set all there PM's to not be displayed in their outbox
	If blnOutbox Then
		strSQL = "UPDATE " & strDbTable & "PMMessage " & _
		"SET " & strDbTable & "PMMessage.Outbox = " & strDBFalse & " " & _
		"WHERE " & strDbTable & "PMMessage.From_ID = " & lngLoggedInUserID & ";"
	
	'If this is the users Inbox set all the PM's to not be displayed in their Inbox
	Else	
		strSQL = "UPDATE " & strDbTable & "PMMessage " & _
		"SET " & strDbTable & "PMMessage.Inbox = " & strDBFalse & " " & _
		"WHERE " & strDbTable & "PMMessage.Author_ID = "  & lngLoggedInUserID & ";"
	End If	
	
	'Update PM inbox/outbox status
	adoCon.Execute(strSQL)
End If




'Removed only the selected PM's from the users inbox/outbox
If Request.Form("delSel") <> "" Then
	
	'Check the session ID to stop spammers using the email form
	Call checkFormID(Request.Form("formID"))
	
	'Run through till all checked messages are deleted
	For each laryMesageID in Request.Form("chkDelete")
	
		'If this is the users Outbox set the PM to not be dispalyed in their outbox
		If blnOutbox Then
			strSQL = "UPDATE " & strDbTable & "PMMessage " & _
			"SET " & strDbTable & "PMMessage.Outbox = " & strDBFalse & " " & _
			"WHERE " & strDbTable & "PMMessage.From_ID = " & lngLoggedInUserID & " " & _
				"AND " & strDbTable & "PMMessage.PM_ID = " & LngC(laryMesageID) & ";"
		
		'If this is the users Inbox set the PM to not be dispalayed in their inbox
		Else	
			strSQL = "UPDATE " & strDbTable & "PMMessage " & _
			"SET " & strDbTable & "PMMessage.Inbox = " & strDBFalse & " " & _
			"WHERE " & strDbTable & "PMMessage.Author_ID = "  & lngLoggedInUserID & " " & _
				"AND " & strDbTable & "PMMessage.PM_ID = " & LngC(laryMesageID) & ";"
		End If	
		
		'Delete the message from the database
		adoCon.Execute(strSQL)
	Next
End If




'Delete PM from button on the page displaying the Private Message 
'(only avialble for the receiptent, not the sender)
If Request.QueryString("pm_id") <> "" AND Request.QueryString("XID") = getSessionItem("KEY") Then
	
	strSQL = "UPDATE " & strDbTable & "PMMessage " & _
	"SET " & strDbTable & "PMMessage.Inbox = " & strDBFalse & " " & _
	"WHERE " & strDbTable & "PMMessage.Author_ID = "  & lngLoggedInUserID & " " & _
		"AND " & strDbTable & "PMMessage.PM_ID = " & LngC(Request.QueryString("pm_id")) & ";"
		
	'Update PM inbox/outbox status
	adoCon.Execute(strSQL)
	
End If






'*** Delete all PM's that are not available in both the inbox and the outbox  ****

'Delete all PM's that are not in the inbox or outbox for entire PM table to keep the PM table up-to-date
'(May take a little longer, but removes all PM's that are no-longer being displayed in the system)	

'SQL to delete all PM's that are not in the inbox or outbox
strSQL = "DELETE FROM " & strDbTable & "PMMessage " & _
"WHERE " & strDbTable & "PMMessage.Inbox = " & strDBFalse & " " & _
	"AND " & strDbTable & "PMMessage.Outbox = " & strDBFalse & ";"
		
'Delete the message from the database
adoCon.Execute(strSQL)





'Update the number of unread PM's 
Call updateUnreadPM(lngLoggedInUserID)
	

	
'Update the notified PM session variable
If intNoOfPms = 0 Then
	Call saveSessionItem("PMN", "")
Else
	Call saveSessionItem("PMN", intNoOfPms)
End If



'Reset Server Objects
Call closeDatabase()



'Return to PM box
If blnOutbox Then
	Response.Redirect("pm_outbox.asp?MSG=DEL" & strQsSID3)
Else
	Response.Redirect("pm_inbox.asp?MSG=DEL" & strQsSID3)
End If
%>