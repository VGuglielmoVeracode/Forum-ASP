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





Response.Buffer = True 


'Check for XID
If Request.QueryString("XID") = getSessionItem("KEY") Then

	'Update the last visit date in the database for the user
	If intGroupID <> 2 Then
		'Initilse sql statement
		strSQL = "UPDATE " & strDbTable & "Author" & strRowLock & " " & _
		"SET " & strDbTable & "Author.Last_visit = " & strDatabaseDateFunction & " " & _
		"WHERE " & strDbTable & "Author.Author_ID = " & lngLoggedInUserID & ";"
		
		'Set error trapping
		On Error Resume Next
		
		'Write to database
		adoCon.Execute(strSQL)
		
		'If an error has occurred write an error to the page
		If Err.Number <> 0 Then	Call errorMsg("An error has occurred while writing to the database.", "getUserData()_update_last_visit", "functions_login.asp")
		
		'Disable error trapping
		On Error goto 0
	End If
	
	
	
	'Set a new cookie with the last date of a forum visit to now
	If blnEuCookieLaw = False OR blnGuest = False Then Call setCookie("lVisit", "LV", internationalDateTime(Now()), True)
	
	'Reset the session variable holding the users last visit to the forum to now
	Call saveSessionItem("LV", internationalDateTime(Now()))
	
	
	'Re-run the unread posts array function
	Session("sarryUnReadPosts") = ""
	Session("sarryUnReadComments") = ""
End If


'Reset Server Objects
Call closeDatabase()

'Return to the forum
Response.Redirect("default.asp" & strQsSID1)

	
%>