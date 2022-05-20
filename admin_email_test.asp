<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
<!--#include file="functions/functions_send_mail.asp" -->
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







'Dimension variables
Dim strAdminEmail		'Holds the admins email address
Dim strEmailBody		'Holds the body of the e-mail message	
Dim strSubject			'Holds the subject of the e-mail
Dim blnEmailSent


blnEmailSent = False


'Reset Server Objects
Call closeDatabase()


'Read in the details from the form
strMailComponent = Request.Form("component")
strMailServer = Request.Form("mailServer")
strMailServerUser = Request.Form("mailServerUser")
strMailServerPass = Request.Form("mailServerPass")
strWebSiteName = Request.Form("siteName")
strForumPath = Request.Form("forumPath")
strAdminEmail = Request.Form("email")
lngMailServerPort = LngC(Request.Form("mailServerPort"))


'Create the body of the email test message
strEmailBody = "This is a test email sent by Web Wiz Forums to test your email settings are correct" & _
vbCrLf & vbCrLf & "Please also click the link below to check that the 'Web address path to forum' has been entered correctly:-" & _
vbCrlf & strForumPath


'Send the e-mail using the Send Mail function created on the send_mail_function.inc file
blnEmailSent = SendMail(strEmailBody, "Forum Admin", decodeString(strAdminEmail), strWebsiteName, decodeString(strAdminEmail), "Web Wiz Forums Email Settings Test", strMailComponent, false)




%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Admin Test Email</title>
<meta name="generator" content="Web Wiz Forums" />
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
</head>
<body OnLoad="self.focus();">
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
  <tr>
    <td align="center"><h1>Test Email Settings</h1></td>
  </tr>
</table>
<br />
<table cellspacing="1" cellpadding="3" class="tableBorder" align="center" width="350">
    <tr class="tableLedger">
      <td colspan="2">Test Email Settings</td>
    <tr class="tableRow"><%

'Message for demo mode
If blnDemoMode Then

%>
      <td align="left" colspan="2"><br /><strong>Emails not sent</strong>
      <br /><br />Emails can not be sent from this software while in demo mode.
      <br /><br /></td><%

'Email sucessfully sent
ElseIf blnEmailSent Then

%>    
      <td align="left" colspan="2"><br /><strong>The email has been successfully sent.</strong>
      <br /><br />Please check your mailbox to make sure you have received the test email.
      <br /><br />Check within the test email that the 'Web address path to your forum' has been entered correctly.
      <br /><br /></td><%


'Email error message     
Else

%>    
      <td align="left" colspan="2"><br /><strong>The email has NOT been sucessfully sent!!</strong>
      <br /><br />Please check your Email Settings are correct.
      <br /><br /><strong>Server Error Message:-</strong>
      <br /><% = strEmailErrorMessage %>
      <br /><br /></td><%

End If

%>
  </tr>
</table>
<br />
<br />
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
  <tr>
    <td align="center"><input type="button" name="ok" onclick="javascript:window.close();" value="Close Window"><br />
      <br /><% 
    
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	
	If blnTextLinks = True Then 
		Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
	Else
  		Response.Write("<a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank"" ><img src=""webwizforums_image.asp"" border=""0"" title=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ alt=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ /></a>")
	End If
	
	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")

'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
End If
%>
    </td>
  </tr>
</table>
</body>
</html>
