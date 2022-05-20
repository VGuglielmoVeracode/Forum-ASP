<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_hash1way.asp" -->
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
Dim strUsername			'Holds the users username
Dim strPassword			'Holds the usres password
Dim blnAutoLogin		'Holds whether the user wnats to be automactically logged in
Dim lngUserID			'Holds the users Id number
Dim strUserCode			'Holds the users ID code
Dim lngLoopCounter		'Holds the loop counter
Dim blnIncorrectLogin		'Set to true if login is incorrect
Dim blnSecurityCodeOK		'Set to false if the security is not OK
Dim strReferer			'Holds the page to return to
Dim blnActive			'Set to true if user is active
Dim blnCAPTCHArequired		'Set to true if CAPTCHA is required
Dim intLoginResponse		'Holds the login response from the login function
Dim strForumName
Dim blnTermsAgreed




'Intialise variables
blnSslEnabledPage = True
blnAutoLogin = false
blnIncorrectLogin = false
blnCAPTCHArequired = false
blnSecurityCodeOK = true


'Get the session ID
strSessionKey = getSessionItem("KEY")
strFormKey = getSessionItem("IDX")



'If this feature is disabled by the member API then redirect the user
If blnMemberAPI AND blnMemberAPIDisableAccountControl Then
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)

End If



'******************************************
'***	      Check IP address		***
'******************************************

'If the	user is	user is	using a	banned IP redirect to an error page
If bannedIP() Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)

End If



'read in the forum ID number
If isNumeric(Request.QueryString("FID")) Then
	intForumID = IntC(Request.QueryString("FID"))
Else
	intForumID = 0
End If


'******************************************
'***	  Check the form key		***
'******************************************
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then	
	If Request.Form(strSessionKey) <> strFormKey Then
			
		'clean up before redirecting
	        Call closeDatabase()
	
		'redirect to insufficient permissions page
	        Response.Redirect("insufficient_permission.asp" & strQsSID1)
	End If
	        
	'Distroy session variable
	Call saveSessionItem("IDX", "")

	'Get the encrypted form name
	strUserNameFormName = "MemberName" & strFormKey
	strPasswordFormName = "P" & HashEncode("Password" & strFormKey)
	
	'Read in the users details from the form
	strUsername = Trim(Mid(Request.Form(strUserNameFormName), 1, 20))
	strPassword = Trim(Mid(Request.Form(strPasswordFormName), 1, 20))
	blnAutoLogin = BoolC(Request.Form("AutoLogin"))
	blnTermsAgreed = BoolC(Request.Form("terms"))
End If


'If a username has been entered check that the password is correct
If strUsername <> "" AND Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	
	
	'**********************************
	'***	 Log user in		***
	'**********************************

	'Call the function to login the user
	intLoginResponse = CInt(loginUser(strUsername, strPassword, blnCAPTCHArequired, "user"))
	
	'Key to loginUser function
	'0 = Login Failed
	'1 = Login OK
	'2 = CAPTCHA Code OK
	'3 = CAPTCHA Code Incorrect
	'4 = CAPTHCA required
	
	
	'If login reponse is 0 then login has failed
	If intLoginResponse = 0 Then blnIncorrectLogin = True
	
	'If login reponse is 3 Then CAPTCHA security code was incorrect
	If intLoginResponse = 3 Then 
		blnSecurityCodeOK = False
		blnCAPTCHArequired = True
	End If
	
	
	'If the login response is 1 the user is logged in
	If intLoginResponse = 1 Then
		
		'Reset Server Objects
		Call closeDatabase()
		
		
		'Get the URL to return to
		If Request("returnURL") <> "" Then
			strReturnURL = Request("returnURL")
		Else
			strReturnURL = Replace(Request.ServerVariables("script_name"), Left(Request.ServerVariables("script_name"), InstrRev(Request.ServerVariables("URL"), "/")), "") & "?" & Request.Querystring
		End If
		
		
		'Clean up input
		strReturnURL = formatLink(strReturnURL)
		strReturnURL = removeAllTags(strReturnURL)
		
		'Replace &amp; with &
		strReturnURL = Replace(strReturnURL, "&amp;", "&",  1, -1, 1)
		
		'For extra security make sure that someone is not trying to send the user to another web site or sneaking through stuff they shouldn't
		strReturnURL = Replace(strReturnURL, "https", "",  1, -1, 1)
		strReturnURL = Replace(strReturnURL, "http", "",  1, -1, 1)
		strReturnURL = Replace(strReturnURL, ":", "",  1, -1, 1)
		strReturnURL = Replace(strReturnURL, "script", "",  1, -1, 1)
		strReturnURL = Replace(strReturnURL, "%", "",  1, -1, 1)
		strReturnURL = Replace(strReturnURL, "#", "",  1, -1, 1)
		strReturnURL = Replace(strReturnURL, "/", "",  1, -1, 1)
		strReturnURL = Replace(strReturnURL, "\", "",  1, -1, 1)
		
		
		If InStr(strReturnURL, "SID") = 0 Then strReturnURL = strReturnURL & strQsSID3
		
		'Go to login user test
		Response.Redirect("login_user_test.asp?" & strReturnURL)

	End If
	
End If


'Setup username field
strUsername = Server.HTMLEncode(strUsername)

'Setup password feild
If blnIncorrectLogin Then 
	strPassword = ""
Else
	strPassword = Server.HTMLEncode(strPassword)
End If


'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers("", strTxtLoginUser, "login_user.asp", 0)
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtLoginUser




%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtLoginUser %></title>
<meta name="robots" content="noindex, follow">

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtLoginUser %></h1></td>
 </tr>
</table><%


'If the user has unsuccesfully tried logging in before then display a password incorrect error
If blnIncorrectLogin OR blnCAPTCHArequired OR (blnSecurityCodeOK = False AND Request.Form("securityCode") <> "") Then
%>
<br />
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><%
	
	'If the login has failed (for extra security only say the password is incorect if the security code matches)
	If blnIncorrectLogin AND blnSecurityCodeOK Then Response.Write("<br />" & strTxtSorryUsernamePasswordIncorrect & "<br />" & strTxtPleaseTryAgain & "<br />")
	
	'If the security code is incorrect
        If blnSecurityCodeOK = False AND Request.Form("securityCode") <> "" Then Response.Write("<br />" & Replace(strTxtSecurityCodeDidNotMatch, "\n\n", "<br />") & "<br />")
	
	'If CAPTCHA s require let the user know
	If blnCAPTCHArequired Then Response.Write("<br />" & strTxtMxLFailedLoginAttemptsMade)
	%></td>
  </tr>
</table><%

End If
%>
<!--#include file="includes/login_form_inc.asp" -->
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

%></div><%


'Reset Server Objects
Call closeDatabase()

'If the user has unsuccesfully tried logging in before then display a password incorrect error
'(for extra security only say the password is incorect if the security code matches)
If blnIncorrectLogin AND blnSecurityCodeOK Then
        Response.Write(vbCrLf & "<script language=""JavaScript"">")
        Response.Write(vbCrLf & "alert('" & strTxtSorryUsernamePasswordIncorrect & "\n\n" &  strTxtPleaseTryAgain & "');")
        Response.Write(vbCrLf & "</script>")

End If

'If the security code did not match
If blnSecurityCodeOK = False AND Request.Form("securityCode") <> "" Then
        Response.Write(vbCrLf & "<script language=""JavaScript"">")
        Response.Write(vbCrLf & "alert('" & strTxtSecurityCodeDidNotMatch & ".');")
        Response.Write(vbCrLf & "</script>")
End If
%>
<script>document.getElementById('frmLogin').<% If Request.Form("QUIK") AND blnCAPTCHArequired Then Response.Write("securityCode") Else Response.Write("name") %>.focus()</script>
<!-- #include file="includes/footer.asp" -->