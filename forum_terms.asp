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



'Dimension variables
Dim rsForum			'Holds the recorset of the forums
Dim strReturnPage		'Holds the page to return to 
Dim strForumName 		'Holds the forum name
Dim strNewSessionKeyID
Dim strFormKey
Dim strMode


'Initialise variables
blnSslEnabledPage = True



'read in the forum ID number
If isNumeric(Request.QueryString("FID")) Then
	intForumID = IntC(Request.QueryString("FID"))
Else
	intForumID = 0
End If


'Read in the page mode
strMode = Trim(Request.QueryString("M"))


'If this feature is disabled by the member API then redirect the user
If blnMemberAPI AND blnMemberAPIDisableAccountControl Then
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)

End If


'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)

End If



'If new reg is suspended and a redirect is active, then redirect
If blnRegistrationSuspeneded AND NOT (strRegistrationRedirect = "" OR strRegistrationRedirect = "http://") Then
		
	'Clean up
	Call closeDatabase()
	
	'do redirect
	Response.Redirect(strRegistrationRedirect)
End If


'If new reg create session keys
If strMode = "reg" Then
	'Create a new session key
	strNewSessionKeyID = LCase(hexValue(12))
	Call saveSessionItem("KEY", strNewSessionKeyID)
	
	'Create registration key
	strFormKey = LCase(hexValue(14))
	Call saveSessionItem("IDX", strFormKey)
End If



'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	If strMode = "reg" Then 
		saryActiveUsers = activeUsers("", strTxtRegisterNewUser, "forum_terms.asp?M=reg", 0)
	Else
		saryActiveUsers = activeUsers("", strTxtForumRulesAndPolicies, "forum_terms.asp", 0)
	End If
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtRegisterNewUser


'Clean up
Call closeDatabase()

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtForumRulesAndPolicies %></title>

<%

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & vbCrLf & "<!--//" & _
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
  <td align="left"><h1><% If strMode = "reg" Then Response.Write(strTxtRegisterNewUser) Else Response.Write(strTxtForumRulesAndPolicies) %></h1></td>
 </tr>
</table>
<br /><%

'If the registration is suspended then display a message saying so
If blnRegistrationSuspeneded AND strMode = "reg" Then
	
%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><br /><% = strTxtNewRegSuspendedCheckBackLater %></td>
  </tr>
</table>
<br /><%


'Else display registration rules
Else

%>
<form method="post" name="frmregister" action="register.asp?FID=<% = Server.HTMLEncode(intForumID) %><% = strQsSID2 %>">
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
  <tr class="tableLedger">
   <td colspan="2" align="left"><% = strTxtForumRulesAndPolicies %></td>
  </tr>
  <tr class="tableRow"> 
   <td align="justify">
    <% = strRegistrationRules %>	
   </td>
  </tr><%

'If new reg display registration button
If strMode = "reg" Then
	
%>
  <tr class="tableBottomRow"> 
   <td align="center"> 
    <input type="button" name="Button" id="Button" value="<% = strTxtCancel %>" onclick="window.open('default.asp', '_self')" />
    <input type="hidden" name="<% = strNewSessionKeyID %>Reg" id="<% = strNewSessionKeyID %>" value="<% = strFormKey %>" />
    <input type="hidden" name="mode" id="mode" value="reg" />
    <input type="submit" name="Registration" id="Registration" value="<% = srtTxtAccept %>" tabindex="1" />
   </td>
  </tr><%

End If

%>
</table>
</form><%

End If

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