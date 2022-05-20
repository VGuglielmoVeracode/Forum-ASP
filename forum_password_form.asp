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

'Dimension variables
Dim strPassword			'Holds the forum password
Dim blnAutoLogin		'Holds whether the user wnats to be automactically logged in
Dim strForumCode		'Holds the users ID code
Dim strReturnPage		'Holds the page to return to
Dim strReturnPageProperties	'Holds the properties of the return page
Dim strFormID			'Holds the ID property for the form
Dim strForumName

'Intilaise varibales
blnSslEnabledPage = True


'Get the forum page to return to
Select Case Request.QueryString("RP")
	'Read in the thread and forum to return to
	Case "PT"
		strReturnPage = "forum_posts.asp"
		strReturnPageProperties = "?RP=PT&TID=" & LngC(Request.QueryString("TID")) & strQsSID3

	'Else return to the forum main page
	Case Else
		'Read in the forum and topic to return to
		strReturnPage = "forum_topics.asp"
		strReturnPageProperties = "?RP=TC&FID=" & IntC(Request.QueryString("FID")) & strQsSID3
End Select


'Read in the forum id number
intForumID = IntC(Request("FID"))

'Read in the users details from the form
strPassword = LCase(Trim(Mid(Request.Form("password"), 1, 15)))
blnAutoLogin = BoolC(Request.Form("AutoLogin"))

'If user has eneterd a password make sure it is correct
If NOT strPassword = "" Then
	
	'Check the form ID
	Call checkFormID(Request.Form("formID"))
	

	'Read in the forum name from the database
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Forum_code " & _
	"FROM " & strDbTable & "Forum" & strDBNoLock & " " & _
	"WHERE Forum_ID = " & intForumID

	'Query the database
	rsCommon.Open strSQL, adoCon


	'If the query has returned a value to the recordset then check the password is correct
	If NOT rsCommon.EOF Then

		'Encrypt the entered password
		strPassword = HashEncode(strPassword)
		

		'Check the password is correct, if it is get the user ID and set a cookie
		If strPassword = rsCommon("Password") Then

			'Read in the users ID number and whether they want to be automactically logged in when they return to the forum
			strForumCode = rsCommon("Forum_code")

			'Save in the session
			Call saveSessionItem("FP" & intForumID, "1")

			'Write a cookie with the Forum ID number so the user logged in throughout the forum
			If blnAutoLogin = True Then

				Call setCookie("fID", "Forum" & intForumID, strForumCode, True)
			
			'Else only temp cookie
			Else
				Call setCookie("fID", "Forum" & intForumID, strForumCode, False)
			
			End If

			'Reset Server Objects
			rsCommon.Close
			Call closeDatabase()

			'Redirect the user back to the forum page they have just come from
			Response.Redirect(strReturnPage & strReturnPageProperties)
		End If
	End If

	'Clean up
	rsCommon.Close
End If

'Reset Server Objects
Call closeDatabase()



'Create a form ID
strFormID = getSessionItem("KEY")




'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers("", strTxtLoginUser, "forum_topics.asp?FID=" & intForumID, 0)
End If


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
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

<script  language="JavaScript">

//Function to check form is filled in correctly before submitting
function CheckForm () {

	var errorMsg = "";
	var formArea = document.getElementById('frmLogin');

	//Check for a Password
	if (formArea.password.value==""){
		errorMsg += "\n<% = strTxtErrorEnterPassword %>";
	}

	//If there is aproblem with the form then display an error
	if (errorMsg != ""){
		msg = "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine1 %>\n";
		msg += "<% = strTxtErrorDisplayLine2 %>\n";
		msg += "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine3 %>\n";

		errorMsg += alert(msg + errorMsg + "\n\n");
		return false;
	}
	
	document.getElementById('formID').value='<% = strFormID %>';

	return true;
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtLoginUser %></h1></td>
 </tr>
</table>
<br /><%


'If the user has unsuccesfully tried logging in before then display a password incorrect error
If strPassword <> "" Then
%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><% = strTxtForumPasswordIncorrect %><br /><br /><% = strTxtPleaseTryAgain %></td>
  </tr>
</table>
<br /><%

End If
%>
<form method="post" name="frmLogin" id="frmLogin" action="forum_password_form.asp<% = strReturnPageProperties %>" onSubmit="return CheckForm();" onReset="return confirm('<% = strResetFormConfirm %>');">
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center" width="350">
 <tr class="tableLedger">
  <td colspan="2"><% = strTxtLoginUser %></td>
 </tr>
 <tr class="tableSubLedger">
  <td colspan="2"><strong><% = strTxtPasswordRequiredForForum %></strong></td>
 </tr>
 <tr class="tableRow">
  <td width="50%"><% = strTxtPassword %></td>
  <td width="50%"><input type="password" name="password" id="password" size="15" maxlength="15" value="" />
  </td>
 </tr>   
 <tr class="tableRow">
  <td width="50%"><% = strTxtAutoLogin %></td>
  <td width="50%"><% = strTxtYes %><input type="radio" name="AutoLogin" id="AutoLogin" value="true" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="AutoLogin" value="false" checked="checked" /></td>
 </tr>
 <tr class="tableBottomRow">
  <td colspan="2" align="center">
   <input type="hidden" name="formID" id="formID" value="" />
   <input type="hidden" name="FID" id="FID" value="<% = intForumID %>" />
   <input type="submit" name="Submit" id="Submit" value="<% = strTxtLoginToForum %>" />
   <input type="reset" name="Reset" id="Reset" value="<% = strTxtResetForm %>" />
 </td>
 </tr>
</table>
</form>
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