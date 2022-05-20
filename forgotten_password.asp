<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_send_mail.asp" -->
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


'Dimension variables
Dim objCDOMail			'Holds the CDO mail object
Dim objJMail			'Holds the Jmail object
Dim strUsername			'Holds the users username
Dim strPassword			'Holds the usres password
Dim strEmailAddress		'Holds the users e-mail address
Dim strReturnPage		'Holds the page to return to 
Dim blnInvalidUsername 		'Set to true if the username entered does not exsit
Dim blnInvalidEmail 		'Set to true if the user has not given there e-mail address	
Dim blnEmailSent		'Set to true if the e-mail has been sent
Dim strEmailBody		'Holds the body of the e-mail message	
Dim strSubject			'Holds the subject of the e-mail
Dim strSalt			'Holds the salt value for the password
Dim strEncyptedPassword		'Holds the encrypted password
Dim strUserCode			'Holds the user code for the user
Dim strUserInput		'Holds teh user input
Dim strSessionKey
Dim strFormKey
Dim strUsernameFormName


'Intialise variables
blnSslEnabledPage = True
blnInvalidUsername = False
blnInvalidEmail = False
blnEmailSent = False

'If e-mail notify is not turned on then close the window
If blnEmail = False Then
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect "default.asp" & strQsSID1
End If




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
	strUsernameFormName = "IUsr" & HashEncode("AdminUser" & strFormKey)
	
	'Read in the users details from the form
	strUserInput = Trim(Mid(Request.Form(strUsernameFormName), 1, 60))
	
End If


'Replace harmful SQL quotation marks with doubles
strUserInput = formatSQLInput(strUserInput)




 
   
'If a username has been entered check that the password is correct
If strUserInput <> "" Then
	
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Author.Username, " & strDbTable & "Author.Password, " & strDbTable & "Author.User_code, " & strDbTable & "Author.Salt, " & strDbTable & "Author.Author_email " & _
	"FROM " & strDbTable & "Author" & strRowLock & " " & _
	"WHERE " & strDbTable & "Author.Username = '" & strUserInput & "' OR " & strDbTable & "Author.Author_email = '" & strUserInput & "';"
	
	'Set the cursor type property of the record set to Forward Only
	rsCommon.CursorType = 0
	
	'Set the Lock Type for the records so that the record set is only locked when it is updated
	rsCommon.LockType = 3
	
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	
	
	'If the query has returned a value to the recordset then generate new password and send it to the user in an email
	If NOT rsCommon.EOF Then
	
		'Read in the users username and email address from the recordset
		strUsername = rsCommon("Username")
		strEmailAddress = rsCommon("Author_email")
		
		'If there is a password in the db to send to change the password and email the user
		If NOT strEmailAddress = "" Then
			
			
			'Read in user code to see if the member is suspended
			strUserCode = rsCommon("User_code")
			
			'For extra security create a new user code for the user
			strUserCode = userCode(strUsername)
			
			
			'Generate a new password using an 8 character long hex values
			strPassword = LCase(hexValue(8))
			
			'If pass is to be encrypted then do so
			If blnEncryptedPasswords Then
				
				'Create a salt value for the new password
				strSalt = getSalt(8)
				
				'Concatenate salt value to the password
				strEncyptedPassword = strPassword & strSalt
				
				'Encrypt the password
				strEncyptedPassword = HashEncode(strEncyptedPassword) 
			
			'Else the password is not to be encrypted
			Else
				strEncyptedPassword = strPassword
			End If
			
			
			'Save new password back to the database with the salt
			rsCommon.Fields("Password") = strEncyptedPassword
			rsCommon.Fields("Salt") = strSalt	
			rsCommon.Fields("User_code") = strUserCode		
			
			'Update the database with the new password
			rsCommon.Update
			
		
		
			'Initailise the e-mail body variable with the body of the e-mail
			strEmailBody = strTxtHi & _
			vbCrLf & vbCrLf & strTxtEmailPasswordRequest & " " & strMainForumName & "." & _
			vbCrLf & vbCrLf & strTxtEmailPasswordRequest2 & _
			vbCrLf & vbCrLf & "----------------------------" & _
			vbCrLf & strTxtUsername & ": - " & strUsername & _
			vbCrLf & strTxtPassword & ": - " & strPassword & _
			vbCrLf & "----------------------------" & _
			vbCrLf & vbCrLf & strTxtEmailPasswordRequest3 & _
			vbCrLf & vbCrLf & "   " & strForumPath
			
			'Initalise the subject of the e-mail
			strSubject = strTxtForumLostPasswordRequest
			
			'Send the e-mail using the Send Mail function created on the send_mail_function.inc file
			blnEmailSent = SendMail(strEmailBody, decodeString(strUsername), decodeString(strEmailAddress), strWebsiteName, decodeString(strForumEmailAddress), strSubject, strMailComponent, false)
			
		Else
			'Set the Invalid e-mail variable to True
			blnInvalidEmail = True	
		End If
	
	
	Else
		'Set the Invalid username variable to True
		blnInvalidUsername = True		
		
	End If
	
	'Clean up
	rsCommon.Close
End If


'Create a form key ID (done for extra security)
strFormKey = LCase(hexValue(14))
Call saveSessionItem("IDX", strFormKey)
	
'Create encrypted form fields
strUsernameFormName = "IUsr" & HashEncode("AdminUser" & strFormKey)
	

'Setup username field
strUserInput = Server.HTMLEncode(strUserInput)


'Reset Server Objects
Call closeDatabase()


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtForgottenPassword

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strTxtForgottenPassword %></title>
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
<!-- Check the from is filled in correctly before submitting -->
<script  language="JavaScript">

//Function to check form is filled in correctly before submitting
function CheckForm () {

	var errorMsg = "";
	var formArea = document.getElementById('frmMailPass');
	
	//Check for a Username
	if (formArea.<% = strUsernameFormName %>.value==""){
	
		msg = "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine1 %>\n";
		msg += "<% = strTxtErrorDisplayLine2 %>\n";
		msg += "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine3 %>\n";
	
		alert(msg + "\n<% = strTxtErrorUsername %>");
		formArea.name.focus();
		return false;
	}
	
	return true
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtForgottenPassword %></h1></td>
</tr>
</table>
<br /><%

'If the user has entered a username that does not exsit then display an error message or security code incorrect
If blnInvalidUsername OR blnInvalidEmail Then
%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><%
    	
    	'If no match in db then
	If blnEmailSent = False AND strEmailErrorMessage <> "" Then Response.Write(strTxtTheEmailFailedToSendPleaseContactAdmin & "<br /><br /><strong>Server Error Message:-</strong><br />" & strEmailErrorMessage & "<br />")
	
	'If no match in db then
	If blnInvalidUsername Then Response.Write(strTxtNoRecordOfUsername & "<br />" & strTxtPleaseTryAgain & "<br />")
		
	'If no match in db then
	If blnInvalidEmail Then Response.Write(strTxtNoEmailAddressInProfile & "<br />" & strTxtReregisterForForum & "<br />")
	
	
	
    	%></td>
  </tr>
</table>
<br /><%

'If the password has been e-mailed to the user then let them know
ElseIf blnEmailSent Then
%>
<table cellspacing="1" cellpadding="3" class="tableBorder" align="center" width="350">
    <tr class="tableLedger">
      <td colspan="2"><% = strTxtForgottenPassword %></td>
    <tr class="tableRow">
      <td align="center" colspan="2"><br /><% = strTxtPasswordEmailToYou %><br /><br /></td>
  </tr>
</table>
<br /><%
  
End If


'show the email form
If blnInvalidEmail = False AND blnEmailSent = False Then
%>
<form method="post" name="frmMailPass" id="frmMailPass" action="forgotten_password.asp<% = strQsSID1 %>" onSubmit="return CheckForm();">
  <table cellspacing="1" cellpadding="3" class="tableBorder" align="center" width="350">
    <tr class="tableLedger">
      <td colspan="2"><% = strTxtForgottenPassword %></td>
    </tr>
    <tr class="tableSubLedger">
      <td align="center" colspan="2"><% = strTxtPleaseEnterYourUsername %></td>
    </tr>
    <tr class="tableRow">
      <td width="50%"><% = strTxtUserNameOrEmailAddress %></td>
      <td width="50%"><input type="text" name="<% = strUsernameFormName %>" id="<% = strUsernameFormName %>" size="30" maxlength="60" value="<% = strUserInput %>">
      </td>
    </tr>
    <tr class="tableBottomRow">
      <td align="center" colspan="2"><input type="hidden" name="<% = strSessionKey %>" id="<% = strSessionKey %>" value="<% = strFormKey %>" /><input type="submit" name="Submit" value="<% = strTxtEmailPassword %>">
      </td>
  </table>
</form><%
  
End If

%>
<br />
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


'If the user details are not recognised display error
If blnInvalidUsername Then
        Response.Write(vbCrLf & "<script language=""JavaScript"">")
        Response.Write(vbCrLf & "alert('" & strTxtNoRecordOfUsername & "\n\n" &  strTxtPleaseTryAgain & "');")
        Response.Write(vbCrLf & "</script>")

End If

'If no email address for user
If blnInvalidEmail Then
        Response.Write(vbCrLf & "<script language=""JavaScript"">")
        Response.Write(vbCrLf & "alert('" & strTxtNoEmailAddressInProfile & "\n\n" &  strTxtReregisterForForum & "');")
        Response.Write(vbCrLf & "</script>")

End If


%>
<!-- #include file="includes/footer.asp" -->
