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


'Update the category and forum order
If Request.Form("Submit") = "Update Order" Then

	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
		
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Category.* From " & strDbTable & "Category ORDER BY " & strDbTable & "Category.Cat_order ASC;"
	
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'Loop through the rs to change the cat order
	Do While NOT rsCommon.EOF
	
		'Initilse sql statement
		 strSQL = "UPDATE " & strDbTable & "Category " & _
		"SET " & strDbTable & "Category.Cat_order = " & IntC(Request.Form("catOrder" & rsCommon("Cat_ID"))) & " " & _
		"WHERE " & strDbTable & "Category.Cat_ID = " & CInt(rsCommon("Cat_ID")) & ";"
		
		'Write to database
		adoCon.Execute(strSQL)
		
		'Move to the next record in the recordset
		rsCommon.MoveNext
	Loop
	
	'Close the recordset
	rsCommon.Close
	
	
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Forum.* From " & strDbTable & "Forum ORDER BY " & strDbTable & "Forum.Forum_Order ASC;"
	
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	
	'Loop through rs to change the forums order
	Do While NOT rsCommon.EOF
		
		'Initilse sql statement
		 strSQL = "UPDATE " & strDbTable & "Forum " & _
		"SET " & strDbTable & "Forum.Forum_Order = " & IntC(Request.Form("forumOrder" & rsCommon("Forum_ID"))) & " " & _
		"WHERE " & strDbTable & "Forum.Forum_ID=" & CInt(rsCommon("Forum_ID")) & ";"
		
		'Write to database
		adoCon.Execute(strSQL)
		
		'Move to the next record in the recordset
		rsCommon.MoveNext
	Loop
	
	'Close the recordset
	rsCommon.Close	
End If
	
'Reset main server variables
Call closeDatabase()


'Return to the forum categories page
Response.Redirect("admin_view_forums.asp" & strQsSID1)
%>