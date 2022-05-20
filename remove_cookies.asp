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
	
	'Clear all forum's cookies on the users system so the user is no longer logged in
	clearCookie()
	
	'Clear sessions
	dtmLastVisitDate = ""
	Call saveSessionItem("UID", "")
	Call saveSessionItem("AID", "")
End If

'Clean up
Call closeDatabase()

'Return to the forum
Response.Redirect("default.asp" & strQsSID1)
%>