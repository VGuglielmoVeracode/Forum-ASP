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


'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"


'Dimension variables
Dim strUsername
Dim strPassword


'If a login is displayed then switch to HTTPS
If intGroupID = 2 AND blnBanned = false AND Request.QueryString("M") = "" Then blnSslEnabledPage = True


'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers(strTxtAccessDenied, "", "", 0)
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtAccessDenied

%>  
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtAccessDenied %></title>
<meta name="robots" content="noindex, follow">

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtAccessDenied %></h1></td>
 </tr>
</table>
<br />
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
 <tr>
  <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
 </tr>
 <tr>
  <td><%

'Display insuficient meesage
Response.Write("<strong>" & strTxtInsufficientPermison & "</strong>")


If Request.QueryString("M") = "DEMO" Then
	
	Response.Write("<br /><br />Sorry this feature is not available in the Demo Version.")
	
'If this is a banned IP then display an error message
ElseIf Request.QueryString("M") = "IP" Then
	
	Response.Write("<br /><br />" & strTxtSorryFunctionNotPermiitedIPBanned)

'If the session ID's don't match then make sure the user has cookies enabled on there system
ElseIf Request.QueryString("M") = "sID" Then

	Response.Write("<br /><br />" & strTxtSessionIDErrorCheckCookiesAreEnabled)
	

'editing time expired
ElseIf Request.QueryString("M") = "eExp" Then

	Response.Write("<br /><br />" & strTxtYouAreOnlyPermittedToEditPostWithin & " " & intEditPostTimeFrame & " " & strTxtMinutes & ".")

'If the users account is suspended then let them know
ElseIf Request.QueryString("M") = "ACT" AND (blnActiveMember = False OR blnBanned)Then

	'If mem suspended display message
	If blnBanned Then
		Response.Write("<br /><br />" & strTxtForumMemberSuspended)
	'Else account not yet active
	Else
		Response.Write("<br /><br />" & strTxtForumMembershipNotAct)
		If blnMemberApprove = False Then Response.Write("<br /><br />" & strTxtToActivateYourForumMem)
	
		'If admin activation is enabled let the user know
		If blnMemberApprove Then
			Response.Write("<br /><br />" & strTxtYouAdminNeedsToActivateYourMembership)
		'If email is on then place a re-send activation email link
		ElseIf blnEmailActivation AND blnLoggedInUserEmail Then 
			Response.Write("<br /><br /><a href=""javascript:winOpener('resend_email_activation.asp" & strQsSID1 & "','actMail',1,1,475,300)"">" & strTxtResendActivationEmail & "</a>")
		End If
	End If
End If

%></td> 
 </tr>
</table><%	
	
'Display a login
If intGroupID = 2 AND blnBanned = false AND Request.QueryString("M") = "" Then
	%><!--#include file="includes/login_form_inc.asp" --><%
End If


'Reset Server Objects
Call closeDatabase()
%>
<br />
<div align="center"><% 

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	
	If blnTextLinks = True Then 
		Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
	Else
  		Response.Write("<a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank"" ><img src=""webwizforums_image.asp"" border=""0"" title=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ alt=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ /></a>")
	End If
	
	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

'Display the process time
If blnShowProcessTime Then Response.Write "<span class=""smText""><br /><br />" & strTxtThisPageWasGeneratedIn & " " & FormatNumber(Timer() - dblStartTime, 3) & " " & strTxtSeconds & "</span>"
%>
</div> 
<!-- #include file="includes/footer.asp" -->