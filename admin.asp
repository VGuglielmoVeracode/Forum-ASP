<% @ Language=VBScript %>
<% Option Explicit %>
<%

Response.Buffer = True 

'First we need to tell the common.asp page to stop redirecting or we'll get in a bit of a loop
blnDisplayForumClosed = True

%>
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
Dim lngLoopCounter		'Holds the loop counter
Dim strUsername			'Holds the users username
Dim strPassword			'Holds the usres password
Dim blnSecurityCodeOK		'Set to true is CATCHA entered correctly
Dim blnIncorrectLogin		'Holds in the user login is correct
Dim intLoginResponse		'Holds the login response from the login function
Dim strFormID
Dim blnAutoLogin
Dim strAdminReferer
Dim strSessionKey
Dim strFormKey
Dim strUsernameFormName
Dim strPasswordFormName


'Intilise varaibles
blnSslEnabledPage = True
blnSecurityCodeOK = True
blnIncorrectLogin = False
blnAutoLogin = False


'read in the session key
strSessionKey = getSessionItem("KEY")
strFormKey = getSessionItem("IDX")



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
	strUserNameFormName = "A" & HashEncode("AdminUser" & strFormKey)
	strPasswordFormName = "P" & HashEncode("Password" & strFormKey)
	
	'Read in the users details from the form
	strUsername = Trim(Mid(Request.Form(strUserNameFormName), 1, 20))
	strPassword = Trim(Mid(Request.Form(strPasswordFormName), 1, 20))
End If



'If this is a new login checkout the login details are correct
If strPassword <> "" Then
	
	
	'Check to see if the user has slected auto-login
	If NOT getCookie("sLID", "UID") = "" Then blnAutoLogin = True
	
	
	'Set the Login incorrect variable to True incase login now fails
	blnIncorrectLogin = True
	
	
	'Call the function to login the user
	intLoginResponse = CInt(loginUser(strUsername, strPassword, False, "admin"))
	
	'Key to loginUser function
	'0 = Login Failed
	'1 = Login OK
	'2 = CAPTCHA Code OK
	'3 = CAPTCHA Code Incorrect
	'4 = CAPTHCA required
	
	
	'If login reponse is 0 then login has failed
	If intLoginResponse = 0 Then blnIncorrectLogin = True
	
	'If login reponse is 3 Then CAPTCHA security code was incorrect
	If intLoginResponse = 3 Then blnSecurityCodeOK = False
		
	'If login is correct setup session and redierct
	If intLoginResponse = 1 Then
		
		'Extra protection make the admin session only valid for the domain the user has logged in through
		
		'Get the refer
		strAdminReferer = LCase(Request.ServerVariables("HTTP_REFERER"))
		
		'Trim the referer down to size
		strAdminReferer = Replace(strAdminReferer, "http://", "")
		strAdminReferer = Replace(strAdminReferer, "https://", "")
		If NOT strAdminReferer = "" Then strAdminReferer = Mid(strAdminReferer, 1, InStr(strAdminReferer, "/")-1)
		If Len(strAdminReferer) > 25 Then strAdminReferer = Mid(strAdminReferer, 1 ,25)
		
		'Save the refer into teh session
		Call saveSessionItem("REF", strAdminReferer)

		
		'For extra security create a new session key for the user
		Call saveSessionItem("KEY", LCase(hexValue(8)))
		
		'Clean up
		Call closeDatabase()
				
		'Redirect to admin section
		Response.Redirect("admin_menu.asp" & strQsSID1)
	End If
End If


	
'Create a form key ID (done for extra security)
strFormKey = LCase(hexValue(14))
Call saveSessionItem("IDX", strFormKey)
	
'Create encrypted form fields
strUsernameFormName = "A" & HashEncode("AdminUser" & strFormKey)
strPasswordFormName = "P" & HashEncode("Password" & strFormKey)


	
'If in demo mode prefill the user name
If blnDemoMode Then strLoggedInUsername = "Administrator"

'Clean up
Call closeDatabase()
	        
%>  
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Web Wiz Forums Control Panel</title>
<meta name="robots" content="noindex, nofollow">

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

<!-- Check the from is filled in correctly before submitting -->
<script language="JavaScript">

//Function to check form is filled in correctly before submitting
function CheckForm () {

	var errorMsg = "";

	//Check for a Username
	if (document.frmLogin.<% = strUsernameFormName %>.value==""){
		errorMsg += "\nUsername \t- Enter the Administrator Forum Username"; 	
	}
	
	//Check for a Password
	if (document.frmLogin.<% = strPasswordFormName %>.value==""){
		errorMsg += "\nPassword \t- Enter the Administrator Forum Password";
	}
	
	//If there is aproblem with the form then display an error
	if (errorMsg != ""){
		msg = "_____________________________________________________________________\n\n";
		msg += "Your Login to the Forum Admin has failed because there are problem(s) with the form.\n";
		msg += "Please correct the problem(s) and re-submit the form.\n";
		msg += "_____________________________________________________________________\n\n";
		msg += "The following field(s) need to be corrected: -\n";
		
		errorMsg += alert(msg + errorMsg + "\n\n");
		return false;
	}
	return true;
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<table width="518" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
  <td align="center"><h1>Web Wiz Forums Control Panel Login</h1></td>
  </tr>
</table>
<div align="center"><a href="default.asp" target="_top">Return to the Main Forum</a></div><%


'If the user has unsuccesfully tried logging in before then display a password incorrect error
If blnIncorrectLogin OR blnSecurityCodeOK = False Then
%>
<br />
<table class="errorTable" cellspacing="1" cellpadding="3" align="center" style="width:500px;">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong><%
	
	'If the login has failed (for extra security only say the password is incorect if the security code matches)
	If blnIncorrectLogin AND blnSecurityCodeOK Then Response.Write("<br /><br />" & strTxtSorryUsernamePasswordIncorrect & "<br />" & strTxtPleaseTryAgain & "<br />")
	
	%></td>
  </tr>
</table><%

End If
%>
<br />
<% If blnDemoMode Then 
	Response.Write("<center>" & _
		"<h2>***** DEMO MODE ****</h2>" & _
		"<strong>Web Wiz Forums is in Demo Mode</strong><br />While public access is permitted to the Admin Control Panel to test features not all features and tools are available to use in Demo Mode<br />" & _
		"<h2>***** DEMO MODE ****</h2>")
End If
%>
<form method="post" name="frmLogin" action="admin.asp<% = strQsSID1 %>"  onSubmit="return CheckForm();">
 <table  cellpadding="4" cellspacing="1" align="center" class="tableBorder" style="width:500px;">
  <tr class="tableLedger"> 
   <td colspan="3">Forum Control Panel Login</td>
  </tr>
  <tr class="tableRow"> 
   <td width="50%" align="right">Admin Username:</td>
   <td width="50%"><%

'If this is an admin they don't need to retype their username
If (intGroupID = 1 OR blnDemoMode) AND blnWindowsAuthentication = False Then
	
	Response.Write(strLoggedInUsername & "<input type=""hidden"" name=""" & strUsernameFormName & """ id=""" & strUsernameFormName & """ value=""" & strLoggedInUsername & """ />")   

'Else text box to write in username as well
Else
   	Response.Write("<input type=""text"" name=""" & strUsernameFormName & """ id=""" & strUsernameFormName & """ size=""15"" maxlength=""20"" />")

End If 
   %></td>
   <td width="71" rowspan="3" class="tableRow"><img src="<% = strImagePath %>admin_login.png" alt="Control Panel Login" /></td>
  </tr>
  <tr class="tableRow"> 
   <td width="50%" align="right">Admin Password:</td>
   <td width="50%" valign="top"> <input type="password" name="<% = strPasswordFormName %>" id="<% = strPasswordFormName %>" size="15" maxlength="20"<% If blnDemoMode Then Response.Write(" value=""letmein""") %> />
   </td>
  </tr>
  <tr class="tableBottomRow"> 
   <td valign="top" height="2" colspan="3" align="center"><input type="hidden" name="<% = strSessionKey %>" id="<% = strSessionKey %>" value="<% = strFormKey %>" /><input type="submit" name="Submit" value="Login &gt;&gt;" />
  </td>
  </tr>
 </table>
</form>
<center>
 <p class="text">Use the same Administration username and password as you use to login to the main forum<br />
  <br />
  If you have forgotten your password then use the forgotten password form in the main forum to <br />
  email yourself a new password, if enabled<br />
  <br />
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
</center>
</body>
</html>
