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

'Clean up
Set rsCommon = Nothing

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


'Before deleting check the XID is OK
If Request.QueryString("XID") = getSessionItem("KEY")  Then
	'Delete the topic from the database	
	'Initalise the strSQL variable with an SQL statement to get the topic from the database
	strSQL = "DELETE FROM " & strDbTable & "BuddyList " & strRowLock & " " & _
	"WHERE (((" & strDbTable & "BuddyList.Author_ID)="  & lngLoggedInUserID & " ) " & _
		" AND ((" & strDbTable & "BuddyList.Address_ID)= " & LngC(Request.QueryString("pm_id")) & "));"
			
	'Delete the message from the database
	adoCon.Execute(strSQL)
End If

'Reset Server Objects
Call closeDatabase()

'Return to the page showing the threads
Response.Redirect("pm_buddy_list.asp" & strQsSID1)
%>