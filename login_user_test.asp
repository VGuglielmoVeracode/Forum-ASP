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
Dim strReturnURL
Dim blnLoggedInOK
Dim strReferer			'Holds the page to return to


'Initilise variables
blnSslEnabledPage = True
blnLoggedInOK = True



'Logged in cookie test
If strLoggedInUserCode = "" Then blnLoggedInOK = False





'Get the forum page to return to
If blnLoggedInOK = False Then
	
	strReturnURL = "login_user.asp?" & Request.QueryString

'Redirect the user back to the forum they have just come from
ElseIf Instr(Request.QueryString, "login_user.asp") Then
	
	strReturnURL = "forum_topics.asp" & Replace(Request.QueryString, "login_user.asp", "", 1, -1, 1)

'If comming from insufficient_permission.asp page redirect back to forum index
ElseIf Instr(Request.QueryString, "insufficient_permission.asp") Then
	strReturnURL = "default.asp" & strQsSID1

'Return to forum homepage
Else
	strReturnURL =  Request.QueryString
End If



'For extra security make sure that someone is not trying to send the user to another web site
strReturnURL = Replace(strReturnURL, "http", "",  1, -1, 1)
strReturnURL = Replace(strReturnURL, ":", "",  1, -1, 1)
strReturnURL = Replace(strReturnURL, "script", "",  1, -1, 1)
strReturnURL = Replace(strReturnURL, "/", "",  1, -1, 1)

'Clean up input
strReturnURL = formatLink(strReturnURL)
strReturnURL = removeAllTags(strReturnURL)
strReturnURL = formatInput(strReturnURL)

If strReturnURL = "" OR Instr(strReturnURL, "?") = 0 Then strReturnURL = "default.asp" & strQsSID1



'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtLoginUser



%>  
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtLoginUser %></title>

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

If blnLoggedInOK Then Response.Write(vbCrLf & "<meta http-equiv=""Refresh"" content=""1;URL=" & strReturnURL & """ />" & vbCrLf)	
%>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtLoginUser %></h1></td>
 </tr>
</table>
<br /><% 

'Reset Server Objects
Call closeDatabase()

'Display heading text
If blnLoggedInOK Then 
	
%>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td colspan="2" align="left"><% = strTxtSuccessfulLogin %></td>
 </tr>
 <tr class="tableRow">
  <td><%
       	
'Else error table      	
Else
	
%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
 <tr>
  <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %>: <% = strTxtUn & strTxtSuccessfulLogin %></strong></td>
 </tr>
 <tr>
  <td><%
End If 

'If the account has not been activated have a link to resend activation email	
If blnBanned = False AND blnEmailActivation AND blnLoggedInUserEmail AND blnActiveMember = False AND blnMemberApprove = False Then

	Response.Write(strTxtForumMembershipNotAct & "<br /><br />" & strTxtToActivateYourForumMem & "<br /><br /><a href=""javascript:winOpener('resend_email_activation.asp','actMail',1,1,475,300)"">" & strTxtResendActivationEmail & "</a><br /><br /><br /><a href=""" & strReturnURL & """>" & strTxtReturnToDiscussionForum & "</a>")


'If the member is suspened then tell them so
ElseIf blnBanned Then

	Response.Write(strTxtForumMemberSuspended & "<br /><br /><a href=""" & strReturnURL & """>" & strTxtReturnToDiscussionForum & "</a>")


'Else if the account neds to be activated then say so
ElseIf blnActiveMember = False Then
	
	Response.Write("<br />" & strTxtForumMembershipNotAct)
	If blnMemberApprove = False Then Response.Write("<br /><br />" & strTxtToActivateYourForumMem)
		
	'If admin activation is enabled let the user know
	If blnMemberApprove Then
		Response.Write("<br /><br />" & strTxtYouAdminNeedsToActivateYourMembership)
	'If email is on then place a re-send activation email link
	ElseIf blnEmailActivation AND blnLoggedInUserEmail Then 
		Response.Write("<br /><br /><a href=""javascript:winOpener('resend_email_activation.asp" & strQsSID1 & "','actMail',1,1,475,300)"">" & strTxtResendActivationEmail & "</a>")
	End If


'If this is a successful login then display some text
ElseIf blnLoggedInOK Then
	
	Response.Write(strTxtSuccessfulLoginReturnToForum & "<br /><br /><a href=""" & strReturnURL & """>" & strTxtReturnToDiscussionForum & "</a>")


'Display that the login was not successful
Else
	Response.Write(strTxtUnSuccessfulLoginText & "<br /><br /><a href=""login_user.asp?" &  removeAllTags(Request.QueryString) & strQsSID2 & """>" & strTxtUnSuccessfulLoginReTry & "</a>")
End If

%></td>
 </tr>
</table>
<br />
<div align="center">
<% 
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