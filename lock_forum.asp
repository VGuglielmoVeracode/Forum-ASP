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



'If in demo mode redirect
If blnDemoMode Then
	Call closeDatabase()
	Response.Redirect("admin_web_wiz_forums_premium.asp" & strQsSID1)
End If


'Dimension variables
Dim strMode		'Holds the mode of the page


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
intForumID = LngC(Request.QueryString("FID"))
strMode = Request.QueryString("mode")

'Check that the user is admin
If blnAdmin Then
	
	'Get the Forum from the database to be locked
	If strMode = "Lock" Then
		strSQL = "UPDATE " & strDbTable & "Forum" & strRowLock & " " & _
		"SET " & strDbTable & "Forum.Locked = " & strDBTrue & " " & _
		"WHERE " & strDbTable & "Forum.Forum_ID ="  & intForumID & ";"
	'Unlock forum
	ElseIf strMode = "UnLock" Then
		strSQL = "UPDATE " & strDbTable & "Forum" & strRowLock & " " & _
		"SET " & strDbTable & "Forum.Locked = " & strDBFalse & " " & _
		"WHERE " & strDbTable & "Forum.Forum_ID ="  & intForumID & ";"
	End If
	
	'Write to the database
	adoCon.Execute(strSQL)
End If


'Reset Server Objects
Call closeDatabase()

'If this is a lock from the admin area then return there
If IntC(Request.QueryString("code")) = 2 Then
	Response.Redirect("admin_view_forums.asp" & strQsSID1)
Else
	'Return to the page showing the threads
	Response.Redirect("forum_topics.asp?FID=" & intForumID & strQsSID3)
End If
%>