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




Dim strReturnURL	'Holds the URL to return to
Dim strSessionKey
Dim strFormKey
Dim strUsernameFormName
Dim strPasswordFormName



'Don't display the login if the forums account control is disabled through the member API
If blnMemberAPIDisableAccountControl = False OR blnMemberAPI = False Then
	
	'read in the session key
	strSessionKey = getSessionItem("KEY")
	
	'Create a form key ID (done for extra security)
	strFormKey = LCase(hexValue(14))
	Call saveSessionItem("IDX", strFormKey)
	
	'Create encrypted form fields
	strUsernameFormName = "MemberName" & strFormKey
	strPasswordFormName = "P" & HashEncode("Password" & strFormKey)
	
	

	'Get the URL to return to
	If Request("returnURL") <> "" Then
		strReturnURL = Request("returnURL")
	Else
		strReturnURL = Replace(Request.ServerVariables("script_name"), Left(Request.ServerVariables("script_name"), InstrRev(Request.ServerVariables("URL"), "/")), "") & "?" & Request.Querystring
	End If

	'For extra security make sure that someone is not trying to send the user to another web site
	strReturnURL = Replace(strReturnURL, "http", "",  1, -1, 1)
	strReturnURL = Replace(strReturnURL, ":", "",  1, -1, 1)
	strReturnURL = Replace(strReturnURL, "script", "",  1, -1, 1)
	
	'Clean up input
	strReturnURL = formatLink(strReturnURL)
	strReturnURL = removeAllTags(strReturnURL)
	
	'Replace &amp; with &
	strReturnURL = Replace(strReturnURL, "&amp;", "&",  1, -1, 1)



%>
<script  language="JavaScript">
function CheckForm () {

	var errorMsg = "";
	var formArea = document.getElementById('frmLogin');

	//Check for a Username
	if (formArea.<% = strUsernameFormName %>.value==""){
		errorMsg += "\n<% = strTxtErrorUsername %>";
	}

	//Check for a Password
	if (formArea.<% = strPasswordFormName %>.value==""){
		errorMsg += "\n<% = strTxtErrorPassword %>";
	}
	
	//Check for trems
	if (formArea.terms_yes.checked==false){
		errorMsg += "\n<% = strTxtErrorTerms %>";
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

	//Disable submit button
	document.getElementById('Submit').disabled=true;

	return true;
}
</script>
<br />
<div id="progressFormArea">
<form method="post" name="frmLogin" id="frmLogin" action="login_user.asp<% = strQsSID1 %>" onSubmit="return CheckForm();" onReset="return confirm('<% = strResetFormConfirm %>');">
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td colspan="2" align="left"><% = strTxtLoginUser %></td>
 </tr>
 <tr class="tableRow">
  <td width="50%"><% = strTxtUsername %></td>
  <td width="50%"><input type="text" name="<% = strUsernameFormName %>" id="<% = strUsernameFormName %>" size="15" maxlength="20" value="<% = strUsername %>" tabindex="1" /> <a href="registration_rules.asp?FID=<% = intForumID & strQsSID2 %>"><% = strNotYetRegistered %></a></td>
 </tr>
 <tr class="tableRow">
  <td><% = strTxtPassword %></td>
  <td><input type="password" name="<% = strPasswordFormName %>" id="<% = strPasswordFormName %>" size="15" maxlength="20" value="<% = strPassword %>" tabindex="2" /><%
    	
	'If email notification is enabled then also show the forgotten password link
	If blnEmail = True Then
		
		%> <a href="forgotten_password.asp<% = strQsSID1 %>"><% = strTxtClickHereForgottenPass %></a><%
	      
	End If
	  
	  %></td>
 </tr>   
 <tr class="tableRow">
  <td><% = strTxtAutoLogin %></td>
  <td><% = strTxtYes %><input type="radio" name="AutoLogin" value="true" checked="checked" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="AutoLogin" value="false" /></td>
 </tr>
 <tr class="tableRow">
  <td><% = strTxtAddMToActiveUsersList %></td>
  <td><% = strTxtYes %><input type="radio" name="NS" value="true" checked="checked" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="NS" value="false" /></td>
 </tr>
 <tr class="tableRow">
  <td><% = srtTxtIAgreeToThe %> <a href="forum_terms.asp" target="_blank"><% = strTxtForumRulesAndPolicies %></a></td>
  <td><% = strTxtYes %><input type="radio" name="terms" id="terms_yes" value="true" checked="checked" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="terms" id="terms_no" value="false" /></td>
 </tr>
 <tr class="tableBottomRow">
  <td colspan="2" align="center">
   <input type="hidden" name="returnURL" id="returnURL" value="<% = strReturnURL %>" tabindex="3" />
   <input type="hidden" name="<% = strSessionKey %>" id="<% = strSessionKey %>" value="<% = strFormKey %>" />
   <input type="submit" name="Submit" id="Submit" value="<% = strTxtLoginUser %>" tabindex="4" />
   <input type="reset" name="Reset" id="Reset" value="<% = strTxtResetForm %>" tabindex="5" />
  </td>
 </tr>
</table>
</form>
</div><%

End If

%>