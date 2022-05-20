<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
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





'Set the response buffer to true
Response.Buffer = True


'Dimension variables
Dim rsAdminDetails	'recordset holding the admin details
Dim strEncyptedPassword	'Holds the new password
Dim blnUserNameOK	'Set to ture if the Name is not already in the database
Dim strCheckUserName	'Holds the Name from the database that we are checking against
Dim blnUpdated		'Set to true if the username and password are updated
Dim strUsername		'Holds the users username
Dim strPassword		'Holds the usres password
Dim strUserCode		'Holds the users ID code
Dim strSalt		'Holds the salt value
Dim blnPasswordComplexityOK	'Set if password is complex enough


'Initialise variables
blnUserNameOK = True
blnUpdated = False
blnPasswordComplexityOK = True

'Redirect if this is not the main forum account
If lngLoggedInUserID <> 1 Then

	'Reset Server Objects
	Call closeDatabase()

	Response.Redirect("admin_menu.asp" & strQsSID1)
End If


'If in demo mode redirect
If blnDemoMode Then
	
	Call closeDatabase()
	Response.Redirect("admin_web_wiz_forums_premium.asp" & strQsSID1)
End If



'If the user is changing there username and password then update the database
If Request.Form("postBack") Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))

	'Read in the userName and password from the form
	strUserName = Trim(Mid(Request.Form("userName"), 1, 20))
	strPassword = Trim(Mid(Request.Form("password"), 1, 20))

	'If there is no userName entered then don't save
	If strUserName = "" Then blnUserNameOK = False
		
	'Check for passowrd complexity
	blnPasswordComplexityOK = passwordComplexity(strPassword, intMinPasswordLength)

	'Clean up user input
        strUserName = formatSQLInput(strUserName)

	'Intialise the ADO recordset object
	Set rsCommon = Server.CreateObject("ADODB.Recordset")

	'Read in the userNames from the database to check the userName does not alreday exsist
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Author.UserName FROM " & strDbTable & "Author WHERE " & strDbTable & "Author.UserName = '" & strUserName & "' AND NOT " & strDbTable & "Author.Author_ID = 1;"

	'Query the database
	rsCommon.Open strSQL, adoCon

	'If there is a record returned then the userName is already in use
	If NOT rsCommon.EOF Then blnUserNameOK = False

	'Remove SQL safe single quote double up set in the format SQL function
        strUserName = Replace(strUserName, "''", "'", 1, -1, 1)


	'Clean up
	rsCommon.Close

	'If the UserName dose not already exsists then save the users details to the database
	If blnUserNameOK AND blnPasswordComplexityOK Then


		'Only encrypt password if this is enabled
		If blnEncryptedPasswords Then
			
			'Generate new salt
	                strSalt = getSalt(Len(strPassword))
	
	                'Concatenate salt value to the password
	                strEncyptedPassword = strPassword & strSalt
	
	                'Re-Genreate encypted password with new salt value
	                strEncyptedPassword = HashEncode(strEncyptedPassword)
		
		'Else the password is not set to be encrypted so place the un-encrypted password into the strEncyptedPassword variable
		Else
			strEncyptedPassword = strPassword
		End If


		'Intialise the strSQL variable with an SQL string to open a record set for the Author table
		strSQL = "SELECT " & strDbTable & "Author.UserName,  " & strDbTable & "Author.Password, " & strDbTable & "Author.Salt, " & strDbTable & "Author.User_code "
		strSQL = strSQL & "From " & strDbTable & "Author "
		strSQL = strSQL & "WHERE " & strDbTable & "Author.Author_ID=1;"


		'Set the Lock Type for the records so that the record set is only locked when it is updated
		rsCommon.LockType = 3

		'Set the Cursor Type to dynamic
		rsCommon.CursorType = 0

		'Open the author table
		rsCommon.Open strSQL, adoCon

		'Randomise the system timer
		Randomize Timer

		'Calculate a code for the user
                strUserCode = userCode(strUserName)

		With rsCommon
			'Update the recordset
			If blnDemoMode = False Then
				.Fields("Username") = strUserName
				.Fields("Password") = strEncyptedPassword
				.Fields("Salt") = strSalt
				.Fields("User_code") = strUserCode
	
				'Update the database with the new user's details
				.Update
			End If

			'Re-run the NewUser query to read in the updated recordset from the database
			.Requery

			're-Save the login code for the user to the users app session
			call saveSessionItem("AID", strUserCode)

			'Read back in the new userName
			strUserName = rsCommon("UserName")

			'Clean up
			.Close
		End With
		
		'Set the update field to true
		blnUpdated = True

	End If
End If

'Reset Server Objects
Call closeDatabase()



%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Change Admin Username &amp; Password</title>
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
<script  language="JavaScript" type="text/javascript">

//Function to check form is filled in correctly before submitting
function CheckForm () {

	//Initialise variables
        var errorMsg = "";
        var errorMsgLong = "";

	//Check for a username
        if (document.frmChangePassword.userName.value.length < <% = intMinUsernameLength %>){
                errorMsg += "\nUserName \t- Your Username must be at least <% = intMinUsernameLength %> characters";
        }

        //Check for a password
        if (document.frmChangePassword.password.value.length < <% = intMinPasswordLength %>){
                errorMsg += "\nPassword \t- Your Password must be at least <% = intMinPasswordLength %> characters";
        }

        //Check both passwords are the same
        if ((document.frmChangePassword.password.value) != (document.frmChangePassword.password2.value)){
                errorMsg += "\nPassword Error\t- The passwords entered do not match";
                document.frmChangePassword.password.value = "";
                document.frmChangePassword.password2.value = "";
        }

        //If there is a problem with the form then display an error
	if ((errorMsg != "") || (errorMsgLong != "")){
		msg = "_________________________________________________________________\n\n";
		msg += "The form has not been submitted because there are problem(s) with the form.\n";
		msg += "Please correct the problem(s) and re-submit the form.\n";
		msg += "_________________________________________________________________\n\n";
		msg += "The following field(s) need to be corrected: -\n";

		errorMsg += alert(msg + errorMsg + "\n" + errorMsgLong);
		return false;
	}
	

	return true;
}
</script>
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center">
  <h1>Change Admin Username &amp; Password</h1><br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    Make sure you <strong>remember</strong> the new<strong> username</strong> and <strong>password</strong> <br />
    as you <strong>will not</strong> be able to Login or <strong>Administer the Forum without them</strong>!!!<br />
    <br />
    Passwords are one way 160bit encrypted and so can NOT be retrieved.<br />
  </p>
</div>
<br />
<%

If blnUserNameOK = False Then
%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center" class="lgText">Sorry the Username you requested is already taken.<br />
      Please choose another Username.</td>
  </tr>
</table><br />
<%

End If



If blnPasswordComplexityOK = False Then
%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center" class="lgText">Sorry the password entered does not meet the Admin Password Complexity.<br /><br />
      The password must be at least <% = intMinPasswordLength %> Characters Long, Contain at least 1 Uppercase Character, 1 Lowercase Character, 1 Number and 1  Non-Alphanumeric Symbol.</td>
  </tr>
</table><br />
<%

End If


If blnUpdated Then
%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td align="center" class="lgText">Your Username and/or Password have been updated.</td>
  </tr>
</table>
<%

End If
%>
<form action="admin_change_admin_username.asp<% = strQsSID1 %>" method="post" name="frmChangePassword" id="frmChangePassword" onsubmit="return CheckForm();">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td colspan="2" class="tableLedger">Change Admin Username and Password </td>
    </tr>
    <tr>
      <td width="40%" align="right" class="tableRow">Username:&nbsp;</td>
      <td width="60%" class="tableRow"><input type="text" name="userName" size="15" maxlength="20" value="<% = strUserName %>"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
    <tr>
      <td width="40%" align="right" class="tableRow">Password:&nbsp; </td>
      <td width="60%" class="tableRow"><input type="password" name="password" size="15" maxlength="20"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td width="40%" align="right" class="tableRow">Confirm Password:&nbsp; </td>
      <td width="60%" class="tableRow"><input type="password" name="password2" size="15" maxlength="20"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td colspan="2" align="center" class="tableBottomRow">
      	<input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
      	<input type="hidden" name="postBack" value="True" />
        <input type="submit" name="Submit" value="Update Details" />
        <input type="reset" name="Reset" value="Clear" />      </td>
    </tr>
  </table>
</form>
<!-- #include file="includes/admin_footer_inc.asp" -->
