<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
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



'If in demo mode redirect
If blnDemoMode Then
	Call closeDatabase()
	Response.Redirect("admin_web_wiz_forums_premium.asp" & strQsSID1)
End If


'Set the response buffer to true as we maybe redirecting
Response.Buffer = True




'Dimension variables
Dim intDeleteLadderGroupID	'Holds the group ID to be deleted


'Check the session key ID to prevent XCSRF
Call checkFormID(Request.QueryString("XID"))


'Get the forum ID to delete
intDeleteLadderGroupID = IntC(Request.QueryString("LID"))


'If there is a group ID to delete then do teh job
If intDeleteLadderGroupID <> "" Then
	
	'Initalise the SQL string with an SQL update command to update any groups that use this ladder group
	strSQL = "UPDATE " & strDbTable & "Group" & strRowLock & " " & _
	 	 "SET " & strDbTable & "Group.Ladder_ID = 0 " & _
	         "WHERE " & strDbTable & "Group.Ladder_ID = " & intDeleteLadderGroupID & ";"

	'Write the updated number of posts to the database
	adoCon.Execute(strSQL)


	'Delete the group form the database
	strSQL = "DELETE FROM " & strDbTable & "LadderGroup " & _
	"WHERE " & strDbTable & "LadderGroup.Ladder_ID = "  & intDeleteLadderGroupID & ";"

	'Write to database
	adoCon.Execute(strSQL)

End If



'Reset Server Objects
Call closeDatabase()


'Return to the forum categories page
Response.Redirect("admin_view_ladder_groups.asp" & strQsSID1)
%>