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
Dim strMode			'Holds the mode of the page, set to true if changes are to be made to the database
Dim strAdminEmail 		'Holds the forum adminsters email
Dim blnEmailNotify		'Set to true to turn email notify on
Dim blnMailActivate		'Set to true if the user wants membership to be activated by email
Dim blnEmailClient		'set to true if the email client is enalbed

'Initialise variables
blnEmailNotify = False

'Read in the details from the form
strMailComponent = Request.Form("component")
strMailServer = Request.Form("mailServer")
strMailServerUser = Request.Form("mailServerUser")
strMailServerPass = Request.Form("mailServerPass")
strWebSiteName = Request.Form("siteName")
strForumPath = Request.Form("forumPath")
strAdminEmail = Request.Form("email")
blnEmailNotify = BoolC(Request.Form("userNotify"))
blnSendPost = BoolC(Request.Form("sendPost"))
blnMailActivate = BoolC(Request.Form("mailActvate"))
blnEmailClient = BoolC(Request.Form("client"))
blnMemberApprove = BoolC(Request.Form("adminApp"))
lngMailServerPort = LngC(Request.Form("mailServerPort"))
blnEmailNotificationSendAll = BoolC(Request.Form("AllNotitfications"))



'If the user is changing the email setup then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	
	Call addConfigurationItem("mail_component", strMailComponent)
	Call addConfigurationItem("mail_server", strMailServer)
	Call addConfigurationItem("website_name", strWebSiteName)
	Call addConfigurationItem("forum_path", strForumPath)
	Call addConfigurationItem("forum_email_address", strAdminEmail)
	Call addConfigurationItem("email_notify", blnEmailNotify)
	Call addConfigurationItem("Email_post", blnSendPost)
	Call addConfigurationItem("Email_activate", blnMailActivate)
	Call addConfigurationItem("Email_sys", blnEmailClient)
	Call addConfigurationItem("Mail_username", strMailServerUser)
	Call addConfigurationItem("Mail_password", strMailServerPass)
	Call addConfigurationItem("Member_approve", blnMemberApprove)
	Call addConfigurationItem("Mail_server_port", lngMailServerPort)
	Call addConfigurationItem("Email_all_notifications", blnEmailNotificationSendAll)

	
	
	'Update variables
	Application.Lock
	
	Application(strAppPrefix & "strMailComponent") = strMailComponent
	Application(strAppPrefix & "strMailServer") = strMailServer
	Application(strAppPrefix & "strWebsiteName") = strWebSiteName
	Application(strAppPrefix & "strForumPath") = strForumPath
	Application(strAppPrefix & "strForumEmailAddress") = strAdminEmail
	Application(strAppPrefix & "blnEmail") = blnEmailNotify
	Application(strAppPrefix & "blnSendPost") = blnSendPost
	Application(strAppPrefix & "blnEmailActivation") = blnMailActivate
	Application(strAppPrefix & "blnEmailMessenger") = blnEmailClient
	Application(strAppPrefix & "strMailServerUser") = strMailServerUser
	Application(strAppPrefix & "strMailServerPass") = strMailServerPass
	Application(strAppPrefix & "blnMemberApprove") = blnMemberApprove
	Application(strAppPrefix & "lngMailServerPort") = lngMailServerPort
	Application(strAppPrefix & "blnEmailNotificationSendAll") = blnEmailNotificationSendAll
	
	'Empty the application level variable so that the changes made are seen in the main forum
	Application(strAppPrefix & "blnConfigurationSet") = false
	
	Application.UnLock
End If



'Initialise the SQL variable with an SQL statement to get the configuration details from the database
strSQL = "SELECT " & strDbTable & "SetupOptions.Option_Item, " & strDbTable & "SetupOptions.Option_Value " & _
"FROM " & strDbTable & "SetupOptions" &  strDBNoLock & " " & _
"ORDER BY " & strDbTable & "SetupOptions.Option_Item ASC;"

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the deatils from the database
If NOT rsCommon.EOF Then
	
	'Place into an array for performance
	saryConfiguration = rsCommon.GetRows()

	'Read in the e-mail setup from the database
	strMailComponent = getConfigurationItem("mail_component", "string")
	strMailServer = getConfigurationItem("mail_server", "string")
	strWebSiteName = getConfigurationItem("website_name", "string")
	strForumPath = getConfigurationItem("forum_path", "string")
	strAdminEmail = getConfigurationItem("forum_email_address", "string")
	blnEmailNotify = CBool(getConfigurationItem("email_notify", "bool"))
	blnSendPost = CBool(getConfigurationItem("Email_post", "bool"))
	blnMailActivate = CBool(getConfigurationItem("Email_activate", "bool"))
	blnEmailClient = CBool(getConfigurationItem("Email_sys", "bool"))
	strMailServerUser = getConfigurationItem("Mail_username", "string")
	strMailServerPass = getConfigurationItem("Mail_password", "string")
	blnMemberApprove = CBool(getConfigurationItem("Member_approve", "bool"))
	lngMailServerPort = CLng(getConfigurationItem("Mail_server_port", "numeric"))	
	blnEmailNotificationSendAll = CBool(getConfigurationItem("Email_all_notifications", "bool"))
End If


'Release Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Email Settings</title>
<meta name="generator" content="Web Wiz Forums" />
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

	//Check for a mail server
	if (((document.frmEmailsetup.component.value=="AspEmail") || (document.frmEmailsetup.component.value=="Jmail")) && (document.frmEmailsetup.mailServer.value=="")){
		alert("Please enter an working incoming mail server \nWithout one the Jmail/AspEmail component will fail");
		document.frmEmailsetup.mailServer.focus();
		return false;
	}

	//Check for a website name
	if (document.frmEmailsetup.siteName.value==""){
		alert("Please enter your Website Name");
		document.frmEmailsetup.siteName.focus();
		return false;
	}

	//Check for a path to the forum
	if (document.frmEmailsetup.forumPath.value==""){
		alert("Please enter the Web Address path to the Forum");
		document.frmEmailsetup.forumPath.focus();
		return false;
	}

	//Check for an email address
	if (document.frmEmailsetup.email.value==""){
		alert("Please enter your E-mail Address");
		document.frmEmailsetup.email.focus();
		return false;
	}

	//Check that the email address is valid
	if (document.frmEmailsetup.email.value.length>0&&(document.frmEmailsetup.email.value.indexOf("@",0)==-1||document.frmEmailsetup.email.value.indexOf(".",0)==-1)) {
		alert("Please enter your valid E-mail address\nWithout a valid email address the email notification will not work");
		document.frmEmailsetup.email.focus();
		return false;
	}

	return true
}	
	
//Function to test email window
function OpenTestEmailWin(formName){

	now = new Date; 
	submitAction = formName.action;
	submitTarget = formName.target;
	
	//Open the window first 	
   	winOpener('','testEmail',1,1,500,275)
   		
   	//Now submit form to the new window
   	formName.action = 'admin_email_test.asp?ID=' + now.getTime()<% = strQsSID2 %>;	
	formName.target = 'testEmail';
	formName.submit();
	
	//Reset submission
	formName.action = submitAction;
	formName.target = submitTarget;
}
</script>
<script language="javascript" src="includes/default_javascript_v9.js" type="text/javascript"></script><!-- #include file="includes/admin_header_inc.asp" -->
<div align="center">
 <h1>Email Settings </h1>
 <br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br />
  <table border="0" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td align="center" class="tableLedger">Important - Please Read First!</td>
    </tr>
    <tr>
      <td align="left" class="tableRow">To be able to use the email notification you need to have either <span>Persists AspEmail, CDONTS</span>, <span>CDOSYS</span>, <span>W3 JMail</span>, or <span>SeverObject AspMail</span> component installed on the web server. <br />
          <br />
          Check with your web hosts or use the  <a href="admin_server_test.asp">Server Compatibility Test</a> page to see which email components are installed on the server.<br />
          <br />
          Persits AspEmail is the recommend component since it supports both authentication and TLS Encryption when sending the email to the SMTP server.<br />
          <br />
          Use the 'Test Email Settings' button to check that the settings entered are correct. If the email fails to arrive check with your SMTP Server Admin for the correct settings to relay email through their servers.<br />
  <br />
  <strong>Outgoing SMTP Server Authentication</strong><br />
Most web hosts now require that you use authentication to login to outgoing SMTP Servers to prevent their servers being used to relay SPAM. The following is a list of the  components that support SMTP authentication:- <br />
     <ul>
       <li><span class="text"><strong>Persits AspEmail</strong> - SMTP Server Authentication and TLS Encryption supported</span></li>
       <li><strong>CDOSYS</strong> - SMTP Server Authentication supported</li>
       <li><span class="text"><strong>Dimac JMail ver.4+</strong> - SMTP Server Authentication supported</span></li>
      </ul></td>
    </tr>
  </table>
</div>
<br />
<form action="admin_email_notify_configure.asp<% = strQsSID1 %>" method="post" name="frmEmailsetup" id="frmEmailsetup" onsubmit="return CheckForm();">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr align="left">
      <td colspan="2" class="tableLedger">Email Component Setup </td>
    </tr>
    <tr>
      <td width="59%" align="left" class="tableRow">Email Component to use:<br />
      <span class="smText">You must have the component you select installed on the web server.<br />You can use the <a href="admin_server_test.asp<% = strQsSID1 %>" class="smLink">Server Compatibility Test Tool</a> to see which components you have installed on the server.</span></td>
      <td width="41%" height="12" valign="top" class="tableRow"><select name="component"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
          <option value="CDOSYS"<% If strMailComponent = "CDOSYS" Then Response.Write(" selected") %>>CDOSYS</option>
          <option value="CDOSYSSSL"<% If strMailComponent = "CDOSYSSSL" Then Response.Write(" selected") %>>CDOSYS with SSL/TLS</option>
          <option value="CDOSYSp"<% If strMailComponent = "CDOSYSp" Then Response.Write(" selected") %>>CDOSYS (Pick Up - Used for localhost)</option>
          <option value="CDONTS"<% If strMailComponent = "CDONTS" Then Response.Write(" selected") %>>CDONTS</option>
          <option value="Jmail"<% If strMailComponent = "Jmail" Then Response.Write(" selected") %>>Dimac JMail</option>
          <option value="Jmail4"<% If strMailComponent = "Jmail4" Then Response.Write(" selected") %>>Dimac Jmail ver.4+</option>
          <option value="AspEmail"<% If strMailComponent = "AspEmail" Then Response.Write(" selected") %>>Persits AspEmail</option>
          <option value="AspEmailTLS"<% If strMailComponent = "AspEmailTLS" Then Response.Write(" selected") %>>Persits AspEmail with TLS</option>
          <option value="AspMail"<% If strMailComponent = "AspMail" Then Response.Write(" selected") %>>ServerObjects AspMail</option>
        </select>      </td>
    </tr>
    <tr>
      <td width="59%" align="left" class="tableRow">Outgoing SMTP Mail Server (<strong>NOT needed for CDONTS</strong>):<br />
      <span class="smText">You only need this if you are using an email component other than CDONTS. It must be a working mail server or the forum will crash.</span></td>
      <td width="41%" height="12" valign="top" class="tableRow"><input type="text" id="mailServer" name="mailServer" maxlength="50" value="<% If strMailServer <> "" Then Response.Write(strMailServer) Else Response.Write("localhost") %>" size="30"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        <br />
      <span class="text">(eg. mail.myweb.com)</span></td>
    </tr>
    <tr class="tableRow">
      <td height="12" align="left">Outgoing SMTP Mail Server Username <br />
          <span class="smText">If the outgoing SMTP Server you are using requires  username authentication then specify it here.<br />
            Please see the list above for email components that support authentication. </span></td>
      <td height="12" valign="top"><input type="text" name="mailServerUser" id="mailServerUser" maxlength="50" value="<% = strMailServerUser %>" size="30"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr class="tableRow">
      <td height="12" align="left">Outgoing SMTP Mail Server Password <br />
          <span class="smText">If the outgoing SMTP Server you are using requires password authentication then specify it here.<br />
            Please see the list above for email components that support authentication. </span></td>
      <td height="12" valign="top"><input type="password" name="mailServerPass" id="mailServerPass" maxlength="50" value="<% = strMailServerPass %>" size="30"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr class="tableRow">
      <td height="12" align="left">Outgoing SMTP Port Number <br />
          <span class="smText">The standard SMTP Port is Port 25, but in some rare cases this is changed, if you are unsure leave it as Port 25.</span></td>
      <td height="12" valign="top"><input type="text" name="mailServerPort" id="mailServerPort" maxlength="8" value="<% = lngMailServerPort %>" size="8"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
    <tr>
      <td height="12" colspan="2" align="left" class="tableLedger">Forum Email Details </td>
    </tr>
    <tr>
      <td width="59%" align="left" class="tableRow">Website name*<br />
        <span class="smText">The name of your website eg. My Website</span></td>
      <td width="41%" height="12" valign="top" class="tableRow"><input type="text" name="siteName" id="siteName" maxlength="50" value="<% = strWebsiteName %>" size="30"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
    <tr>
      <td width="59%"  height="2" align="left" class="tableRow">Web address path to forum*<br />
        <span class="smText">The web address that you would type into your web browsers address bar in order to get to the forum. <br />
      eg. http://www.mywebsite.com/forum </span></td>
      <td width="41%" height="2" valign="top" class="tableRow"><input type="text" name="forumPath" id="forumPath" maxlength="70" value="<% = strForumPath %>" size="30"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
    <tr>
      <td width="59%"  height="23" align="left" class="tableRow">Forum Email Address* <br />
        <span class="smText">Without a valid email address you wont be able to send emails from the forum, activate forum memberships (if enabled), etc. </span><br />      </td>
      <td width="41%" height="23" valign="top" class="tableRow"><input type="text" name="email" id="email" maxlength="50" value="<% = strAdminEmail %>" size="30"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;</td>
    </tr>
    
    <tr>
      <td height="12" colspan="2" align="left" class="tableLedger">Test Email Settings</td>
    </tr>
    <tr>
      <td width="59%" align="left" class="tableRow">Test Email Settings<br />
        <span class="smText">Use the button on the right to send a test email to the email address above to test if the settings are correct</span></td>
      <td width="41%" height="12" valign="top" class="tableRow"><input type="button" name="testEamil" id="testEamil" value="Test Email Settings" onclick="OpenTestEmailWin(document.frmEmailsetup)" />      </td>
    </tr>
    <tr>
    
    <tr>
      <td  height="7" colspan="2" align="left" class="tableLedger">Enable Email Features </td>
    </tr>
    <tr>
      <td width="59%"  height="7" align="left" class="tableRow">Forum Wide Emails<br />
      <span class="smText">If enabled allows emails to be sent from the forum for Post Notification, Forgotten Passwords, Private Messages, etc.</span></td>
      <td width="41%" height="7" valign="top" class="tableRow">Yes
        <input type="radio" name="userNotify" value="True" <% If blnEmailNotify = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;No
        <input type="radio" name="userNotify" value="False" <% If blnEmailNotify = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
    <tr>
      <td width="59%"  height="13" align="left" class="tableRow">Send Post with Email Notification<br />
        <span class="smText">Allow the full message that has been posted in the forum to be sent with the email 
      notification.</span></td>
      <td width="41%" height="13" valign="top" class="tableRow">Yes
      <input type="radio" name="sendPost" value="True" <% If blnSendPost = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;No
      <input type="radio" name="sendPost" value="False" <% If blnSendPost = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
     <tr>
      <td width="59%"  height="13" align="left" class="tableRow">Send All Post Notifications<br />
        <span class="smText">Enable this option if you want all email notifications sent regardless if the members has been sent one already. If enabled will slow your forum and could lead to spam complaints and domain/hosting blacklisting.</span></td>
      <td width="41%" height="13" valign="top" class="tableRow">Yes
      <input type="radio" name="AllNotitfications" value="True" <% If blnEmailNotificationSendAll = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;No
      <input type="radio" name="AllNotitfications" value="False" <% If blnEmailNotificationSendAll = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
    <tr>
      <td  height="13" align="left" class="tableRow">Email Activation of Membership<br />
        <span class="smText">With this option available new members will be required to activate their membership via a validation email (This option is disabled if Administrator Member Activation is enabled).</span> </td>
      <td height="13" valign="top" class="tableRow">Yes
      <input type="radio" name="mailActvate" value="True" <% If blnMailActivate = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;No
      <input type="radio" name="mailActvate" value="False" <% If blnMailActivate = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td  height="13" align="left" class="tableRow">Administrator Member Activation<br />
      <span class="smText">New members will not be able to use their account till the admin activates their membership. An activation email will be sent to the forum email address entered above.</span></td>
     <td height="13" valign="top" class="tableRow">Yes
      <input type="radio" name="adminApp" value="True" <% If blnMemberApprove = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
&nbsp;&nbsp;No
<input type="radio" name="adminApp" value="False" <% If blnMemberApprove = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td width="59%"  height="13" align="left" class="tableRow">Built in Email Client<br />
      <span class="smText">The built in email client allows members to send emails to other forum members directly from the forum, as long as both parties have a valid email address in their profile.</span></td>
      <td width="41%" height="13" valign="top" class="tableRow">Yes
      <input type="radio" name="client" value="True" <% If blnEmailClient = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;No
      <input type="radio" name="client" value="False" <% If blnEmailClient = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
    <tr align="center">
      <td height="2" colspan="2" valign="top" class="tableBottomRow" >
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Update Details" />
          <input type="reset" name="Reset" value="Reset Form" />      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
