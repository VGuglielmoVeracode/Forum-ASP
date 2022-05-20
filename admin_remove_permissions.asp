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





'Set the response buffer to true as we maybe redirecting
Response.Buffer = True


'Dimension variables
Dim lngMemberID		'Holds the member ID number
Dim strMode		'Hols the page mode


'Check the session key ID to prevent XCSRF
Call checkFormID(Request.QueryString("XID"))

'Read in the details
intForumID = IntC(Request("FID"))
lngMemberID = LngC(Request("UID"))


'User Permission only delete the user permimissons
'Delete the user permissions for forums form the database
strSQL = "DELETE FROM " & strDbTable & "Permissions WHERE " & strDbTable & "Permissions.Forum_ID=" & intForumID & " AND " & strDbTable & "Permissions.Author_ID = " & lngMemberID & ";"


'Write to database
adoCon.Execute(strSQL)
	



'Reset Server Objects
Call closeDatabase()


'Return to the forum user permisisons
Response.Redirect("admin_user_permissions.asp?UID=" & lngMemberID & strQsSID3)

%>