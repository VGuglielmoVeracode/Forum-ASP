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

Response.ContentType = "text/html"



'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"
	

'If this is just a ping to kepe things alive don't return anything
If Request.QueryString("Ping") = "alive" Then

	Response.Write("&nbsp;")
	'Response.Write("Ping:" & FormatDateTime(Now(), 3))
	
'If this is for a form then need to create some additional fields for the form (check the sesison key for security before creating)
ElseIf Request.QueryString("XID") = getSessionItem("KEY") Then
	
	Response.Write("<input type=""hidden"" name=""session"" id=""session"" value=""true"" />")
	'Response.Write("Ping:" & FormatDateTime(Now(), 3))

'Else if session key passed does not match then formID is empty
ElseIf Request.QueryString("XID") <> getSessionItem("KEY") Then
	
	Response.Write("<input type=""hidden"" name=""session"" id=""session"" value=""false"" />")

End If

'Clean up
Call closeDatabase()

%>