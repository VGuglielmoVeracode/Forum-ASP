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




      

'Read in the users details for the forum
blnRegistrationSuspeneded = BoolC(Request.Form("suspendReg"))
strRegistrationRedirect = Trim(Request.Form("registrationRedirect"))
blnAvatar = BoolC(Request.Form("avatar"))
blnLongRegForm = BoolC(Request.Form("reg"))
blnModeratorProfileEdit = BoolC(Request.Form("modEdit"))
intMinPasswordLength = IntC(Request.Form("minPass"))
intMinUsernameLength = IntC(Request.Form("minUser"))
blnEnforceComplexPasswords = BoolC(Request.Form("passComplxity"))
blnRealNameReq = BoolC(Request.Form("realName"))
blnLocationReq = BoolC(Request.Form("location"))
blnSignatures = BoolC(Request.Form("signatures"))
blnHomePage = BoolC(Request.Form("homepageURL"))
blnRegistrationCAPTCHA = BoolC(Request.Form("CAPTCHA"))
strRegPrivacyNotice = Request.Form("regPrivacyNotice")

strCustRegItemName1 = Trim(Request.Form("custRegItemName1"))
blnReqCustRegItemName1 = BoolC(Request.Form("reqCustRegItemName1"))
blnViewCustRegItemName1 = BoolC(Request.Form("viewCustRegItemName1"))
strCustRegItemName2 = Trim(Request.Form("custRegItemName2"))
blnReqCustRegItemName2 = BoolC(Request.Form("reqCustRegItemName2"))
blnViewCustRegItemName2 = BoolC(Request.Form("viewCustRegItemName2"))
strCustRegItemName3 = Trim(Request.Form("custRegItemName3"))
blnReqCustRegItemName3 = BoolC(Request.Form("reqCustRegItemName3"))
blnViewCustRegItemName3 = BoolC(Request.Form("viewCustRegItemName3"))





'Make sure that blank cuistom fields are not set to be required, or will throw an error
If strCustRegItemName1 = "" Then blnReqCustRegItemName1 = False
If strCustRegItemName2 = "" Then blnReqCustRegItemName2 = False
If strCustRegItemName3 = "" Then blnReqCustRegItemName3 = False



strRegistrationRules = Request.Form("registrationRules")

'Clear registration redirect if nothing enetered
If strRegistrationRedirect = "http://" Then strRegistrationRedirect = ""


'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then	
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	Call addConfigurationItem("Reg_closed", blnRegistrationSuspeneded)
	Call addConfigurationItem("Reg_Redirect", strRegistrationRedirect)
	Call addConfigurationItem("Avatar", blnAvatar)
	Call addConfigurationItem("Long_reg", blnLongRegForm)
	Call addConfigurationItem("Mod_profile_edit", blnModeratorProfileEdit)
	Call addConfigurationItem("Min_password_length", intMinPasswordLength)
	Call addConfigurationItem("Min_usename_length", intMinUsernameLength)
	Call addConfigurationItem("Password_complexity", blnEnforceComplexPasswords)
	Call addConfigurationItem("Real_name", blnRealNameReq)
	Call addConfigurationItem("Location", blnLocationReq)
	Call addConfigurationItem("Signatures", blnSignatures)
	Call addConfigurationItem("Homepage", blnHomePage)
	Call addConfigurationItem("Registration_Rules", strRegistrationRules)
	Call addConfigurationItem("Registration_CAPTCHA", blnRegistrationCAPTCHA)
	Call addConfigurationItem("Reg_Privacy_Notice", strRegPrivacyNotice)
	
	
	Call addConfigurationItem("Cust_item_name_1", strCustRegItemName1)
	Call addConfigurationItem("Cust_item_name_req_1", blnReqCustRegItemName1)
	Call addConfigurationItem("Cust_item_name_view_1", blnViewCustRegItemName1)
	Call addConfigurationItem("Cust_item_name_2", strCustRegItemName2)
	Call addConfigurationItem("Cust_item_name_req_2", blnReqCustRegItemName2)
	Call addConfigurationItem("Cust_item_name_view_2", blnViewCustRegItemName2)
	Call addConfigurationItem("Cust_item_name_3", strCustRegItemName3)
	Call addConfigurationItem("Cust_item_name_req_3", blnReqCustRegItemName3)
	Call addConfigurationItem("Cust_item_name_view_3", blnViewCustRegItemName3)
	

		
					
	'Update variables
	Application.Lock
	
	Application(strAppPrefix & "blnRegistrationSuspeneded") = CBool(blnRegistrationSuspeneded)
	Application(strAppPrefix & "strRegistrationRedirect") = strRegistrationRedirect
	Application(strAppPrefix & "blnAvatar") = CBool(blnAvatar)
	Application(strAppPrefix & "blnLongRegForm") = CBool(blnLongRegForm)
	Application(strAppPrefix & "blnModeratorProfileEdit") = CBool(blnModeratorProfileEdit)
	Application(strAppPrefix & "intMinPasswordLength") = Cint(intMinPasswordLength)
	Application(strAppPrefix & "intMinUsernameLength") = Cint(intMinUsernameLength)
	Application(strAppPrefix & "blnEnforceComplexPasswords") = CBool(blnEnforceComplexPasswords)
	Application(strAppPrefix & "blnRealNameReq") = CBool(blnRealNameReq)
	Application(strAppPrefix & "blnLocationReq") = CBool(blnLocationReq)
	Application(strAppPrefix & "blnSignatures") = CBool(blnSignatures)
	Application(strAppPrefix & "blnHomePage") = CBool(blnHomePage)
	Application(strAppPrefix & "strRegistrationRules") = strRegistrationRules
	Application(strAppPrefix & "blnRegistrationCAPTCHA") = CBool(blnRegistrationCAPTCHA)
	Application(strAppPrefix & "strRegPrivacyNotice") = strRegPrivacyNotice
	
	Application(strAppPrefix & "strCustRegItemName1") = strCustRegItemName1
	Application(strAppPrefix & "blnReqCustRegItemName1") = CBool(blnReqCustRegItemName1)
	Application(strAppPrefix & "blnViewCustRegItemName1") = CBool(blnViewCustRegItemName1)
	Application(strAppPrefix & "strCustRegItemName2") = strCustRegItemName2
	Application(strAppPrefix & "blnReqCustRegItemName2") = CBool(blnReqCustRegItemName2)
	Application(strAppPrefix & "blnViewCustRegItemName2") = CBool(blnViewCustRegItemName2)
	Application(strAppPrefix & "strCustRegItemName3") = strCustRegItemName3
	Application(strAppPrefix & "blnReqCustRegItemName3") = CBool(blnReqCustRegItemName3)
	Application(strAppPrefix & "blnViewCustRegItemName3") = CBool(blnViewCustRegItemName3)
	
	Application(strAppPrefix & "blnConfigurationSet") = false
	Application.UnLock
End If





'Initialise the SQL variable with an SQL statement to get the configuration details from the database
strSQL = "SELECT " & strDbTable & "SetupOptions.Option_Item, " & strDbTable & "SetupOptions.Option_Value " & _
"FROM " & strDbTable & "SetupOptions" &  strDBNoLock & " " & _
"ORDER BY " & strDbTable & "SetupOptions.Option_Item ASC;"

	
'Query the database
rsCommon.Open strSQL, adoCon

'Read in the forum from the database
If NOT rsCommon.EOF Then
	
	'Place into an array for performance
	saryConfiguration = rsCommon.GetRows()
	
	'Read in the colour info from the database
	blnRegistrationSuspeneded = CBool(getConfigurationItem("Reg_closed", "bool"))
	strRegistrationRedirect = getConfigurationItem("Reg_Redirect", "string")
	blnAvatar = CBool(getConfigurationItem("Avatar", "bool"))
	blnLongRegForm = CBool(getConfigurationItem("Long_reg", "bool"))
	blnModeratorProfileEdit = CBool(getConfigurationItem("Mod_profile_edit", "bool"))
	intMinPasswordLength = CInt(getConfigurationItem("Min_password_length", "numeric"))
	intMinUsernameLength = CInt(getConfigurationItem("Min_usename_length", "numeric"))
	blnEnforceComplexPasswords = CBool(getConfigurationItem("Password_complexity", "bool"))
	blnRealNameReq = CBool(getConfigurationItem("Real_name", "bool"))
	blnLocationReq = CBool(getConfigurationItem("Location", "bool"))
	blnSignatures = CBool(getConfigurationItem("Signatures", "bool"))
	blnHomePage = CBool(getConfigurationItem("Homepage", "bool"))
	blnRegistrationCAPTCHA = CBool(getConfigurationItem("Registration_CAPTCHA", "bool"))
	strRegPrivacyNotice = getConfigurationItem("Reg_Privacy_Notice", "string")
	
	strCustRegItemName1 = getConfigurationItem("Cust_item_name_1", "string")
	blnReqCustRegItemName1 = CBool(getConfigurationItem("Cust_item_name_req_1", "bool"))
	blnViewCustRegItemName1 = CBool(getConfigurationItem("Cust_item_name_view_1", "bool"))
	strCustRegItemName2 = getConfigurationItem("Cust_item_name_2", "string")
	blnReqCustRegItemName2 = CBool(getConfigurationItem("Cust_item_name_req_2", "bool"))
	blnViewCustRegItemName2 = CBool(getConfigurationItem("Cust_item_name_view_2", "bool"))
	strCustRegItemName3 = getConfigurationItem("Cust_item_name_3", "string")
	blnReqCustRegItemName3 = CBool(getConfigurationItem("Cust_item_name_req_3", "bool"))
	blnViewCustRegItemName3 = CBool(getConfigurationItem("Cust_item_name_view_3", "bool"))
	
	strRegistrationRules = getConfigurationItem("Registration_Rules", "string")
	
End If

'If registration redirect is blank initialise with http://
If strRegistrationRedirect = "" Then strRegistrationRedirect = "http://"


'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Registration and Profile Settings</title>
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
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center">
  <h1>Registration and Profile Settings</h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    From here you can configure what information your members are required to give when registering and in their forum profile.<br />
    <br />
</div>
<form action="admin_registration_settings.asp<% = strQsSID1 %>" method="post" name="frmConfiguration" id="frmConfiguration" onsubmit="return CheckForm();">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
     <td colspan="2" class="tableLedger">Suspend Registration</td>
    </tr>
    <tr>
      <td width="57%" align="left" class="tableRow">Suspend New Registrations:<br />
      <span class="smText">This will prevent new Members Registering to use your forum.</span></td>
      <td width="43%" valign="top" class="tableRow">Yes
        <input type="radio" name="suspendReg" value="True" <% If blnRegistrationSuspeneded = True Then Response.Write "checked" %> />
        &nbsp; No
        <input type="radio" name="suspendReg" value="False" <% If blnRegistrationSuspeneded = False Then Response.Write "checked" %> />
      </td>
    </tr>
    <tr>
     <td class="tableRow" width="57%">Suspended Registration Redirect:<br />
      <span class="smText">If New Registrations are Suspended you can use this option to send visitors clicking the Registration button in your forum to another web page. For example you could redirect to your own websites Registration page.</span></td>
      <td class="tableRow"><input type="text" name="registrationRedirect" id="registrationRedirect" maxlength="100" value="<% = strRegistrationRedirect %>" size="40"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
     <tr>
     <td colspan="2" class="tableLedger">Registration and Profile Settings</td>
    </tr>
    <tr>
     <td class="tableRow" width="57%">Full Registration Form:<br />
      <span class="smText">If disabled then new members registering will see a shortened version of the registration form.</span></td>
     <td valign="top" class="tableRow" width="43%">Yes
      <input type="radio" name="reg" value="True" <% If blnLongRegForm = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" value="False" <% If blnLongRegForm = False Then Response.Write "checked" %> name="reg"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">Minimum Username Length:<br />
      <span class="smText">This is minimum length allowed for Members Usernames (max. 20).</span></td>
     <td valign="top" class="tableRow"><select name="minUser" id="minUser"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
      <option<% If intMinUsernameLength = 1 Then Response.Write(" selected") %>>1</option>
      <option<% If intMinUsernameLength = 2 Then Response.Write(" selected") %>>2</option>
      <option<% If intMinUsernameLength = 3 Then Response.Write(" selected") %>>3</option>
      <option<% If intMinUsernameLength = 4 Then Response.Write(" selected") %>>4</option>
      <option<% If intMinUsernameLength = 5 Then Response.Write(" selected") %>>5</option>
      <option<% If intMinUsernameLength = 6 Then Response.Write(" selected") %>>6</option>
      <option<% If intMinUsernameLength = 7 Then Response.Write(" selected") %>>7</option>
      <option<% If intMinUsernameLength = 8 Then Response.Write(" selected") %>>8</option>
      <option<% If intMinUsernameLength = 9 Then Response.Write(" selected") %>>9</option>
      <option<% If intMinUsernameLength = 10 Then Response.Write(" selected") %>>10</option>
     </select></td>
    </tr>
    <tr>
     <td class="tableRow">Minimum Password Length:<br />
      <span class="smText">This is minimum length allowed for Members Passwords (max. 20).</span></td>
     <td valign="top" class="tableRow"><select name="minPass" id="minPass"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
      <option<% If intMinPasswordLength = 1 Then Response.Write(" selected") %>>1</option>
      <option<% If intMinPasswordLength = 2 Then Response.Write(" selected") %>>2</option>
      <option<% If intMinPasswordLength = 3 Then Response.Write(" selected") %>>3</option>
      <option<% If intMinPasswordLength = 4 Then Response.Write(" selected") %>>4</option>
      <option<% If intMinPasswordLength = 5 Then Response.Write(" selected") %>>5</option>
      <option<% If intMinPasswordLength = 6 Then Response.Write(" selected") %>>6</option>
      <option<% If intMinPasswordLength = 7 Then Response.Write(" selected") %>>7</option>
      <option<% If intMinPasswordLength = 8 Then Response.Write(" selected") %>>8</option>
      <option<% If intMinPasswordLength = 9 Then Response.Write(" selected") %>>9</option>
      <option<% If intMinPasswordLength = 10 Then Response.Write(" selected") %>>10</option>
     </select></td>
    </tr>
    <tr>
     <td class="tableRow">Enforce Password Complexity:<br />
       <span class="smText">This will enable Complex Passwords, allowing members to only create passwords that contain at least 1 Uppercase Character, 1 Lowercase Character, 1 Number and 1 Non-Alphanumeric Symbol.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="passComplxity" value="True" <% If blnEnforceComplexPasswords = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="passComplxity" value="False" <% If blnEnforceComplexPasswords = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
    <tr>
    <tr>
     <td class="tableRow">Real Name Required:<br />
       <span class="smText">When enabled members are required to give their Real Name when registering and editing their forum profile.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="realName" value="True" <% If blnRealNameReq = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="realName" value="False" <% If blnRealNameReq = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
     <tr>
     <td class="tableRow">Location Required:<br />
       <span class="smText">When enabled members are required to give their Location when registering and editing their forum profile.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="location" value="True" <% If blnLocationReq = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="location" value="False" <% If blnLocationReq = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
    <tr>
     <td class="tableRow">Signatures:<br />
       <span class="smText">Allow members to create and attach signatures to their Forum Profiles and Posts.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="signatures" value="True" <% If blnSignatures = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="signatures" value="False" <% If blnSignatures = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
    <tr>
     <td class="tableRow">Avatar Images:<br />
       <span class="smText">These are the small images shown next to Members details within the forum system.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="avatar" value="True" <% If blnAvatar = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="avatar" value="False" <% If blnAvatar = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
    <tr>
     <td class="tableRow">Homepage:<br />
       <span class="smText">Allow members to add a Homepage URL to their websites to be shown within the forum system.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="homepageURL" value="True" <% If blnHomePage = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="homepageURL" value="False" <% If blnHomePage = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     </td>
    </tr>
     <td class="tableRow">Moderator Member Profile Editing and Spam Cleaner:<br />
       <span class="smText">When enabled Moderators are able to edit the Forum Profiles of all but the Admin Account and use the Spam Cleaner to disable Spammers Accounts and removed their Posts, Private Messages and Block their access.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="modEdit" value="True" <% If blnModeratorProfileEdit = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="modEdit" value="False" <% If blnModeratorProfileEdit = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />     </td>
    </tr>
    
    <tr>
     <td class="tableRow"><a href="http://www.webwizcaptcha.com" target="_blank">Web Wiz CAPTCHA</a> for New Registrations:<br />
       <span class="smText">When enabled new registrations will have a CAPTCHA security image to block automated registrations from spammers.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="CAPTCHA" value="True" <% If blnRegistrationCAPTCHA = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" value="False" <% If blnRegistrationCAPTCHA = False Then Response.Write "checked" %> name="CAPTCHA"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
    
    <tr>
     <td colspan="2" class="tableLedger">Custom Registration/Profile Item 1</td>
    </tr>
    <tr>
     <td class="tableRow" width="57%">Item Name:<br />
      <span class="smText">This is the name that you wish displayed for the Custom Registration Item.</span></td>
      <td class="tableRow"><input type="text" name="custRegItemName1" id="custRegItemName1" maxlength="25" value="<% = strCustRegItemName1 %>" size="25"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
    <td class="tableRow">Required:<br />
       <span class="smText">When enabled members are required to fill in this item when registering and editing their forum profile.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="reqCustRegItemName1" value="True" <% If blnReqCustRegItemName1 = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="reqCustRegItemName1" value="False" <% If blnReqCustRegItemName1 = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
    </td>
   </tr>
   <tr>
    <td class="tableRow">Viewed in Member Profile:<br />
       <span class="smText">When enabled all members are able to view this item in member profiles. If disabled only Admins and Moderators can view this item in member profiles.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="viewCustRegItemName1" value="True" <% If blnViewCustRegItemName1 = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="viewCustRegItemName1" value="False" <% If blnViewCustRegItemName1 = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
    </td>
   </tr>
   
   <tr>
     <td colspan="2" class="tableLedger">Custom Registration/Profile Item 2</td>
    </tr>
    <tr>
     <td class="tableRow" width="57%">Item Name:<br />
      <span class="smText">This is the name that you wish displayed for the Custom Registration Item.</span></td>
      <td class="tableRow"><input type="text" name="custRegItemName2" id="custRegItemName2" maxlength="25" value="<% = strCustRegItemName2 %>" size="25"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
    <td class="tableRow">Required:<br />
       <span class="smText">When enabled members are required to fill in this item when registering and editing their forum profile.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="reqCustRegItemName2" value="True" <% If blnReqCustRegItemName2 = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="reqCustRegItemName2" value="False" <% If blnReqCustRegItemName2 = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
    </td>
   </tr>
   <tr>
    <td class="tableRow">Viewed in Member Profile:<br />
       <span class="smText">When enabled all members are able to view this item in member profiles. If disabled only Admins and Moderators can view this item in member profiles.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="viewCustRegItemName2" value="True" <% If blnViewCustRegItemName2 = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="viewCustRegItemName2" value="False" <% If blnViewCustRegItemName2 = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
    </td>
   </tr>
   
   <tr>
     <td colspan="2" class="tableLedger">Custom Registration/Profile Item 3</td>
    </tr>
    <tr>
     <td class="tableRow" width="57%">Item Name:<br />
      <span class="smText">This is the name that you wish displayed for the Custom Registration Item.</span></td>
      <td class="tableRow"><input type="text" name="custRegItemName3" id="custRegItemName3" maxlength="25" value="<% = strCustRegItemName3 %>" size="25"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
    <td class="tableRow">Required:<br />
       <span class="smText">When enabled members are required to fill in this item when registering and editing their forum profile.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="reqCustRegItemName3" value="True" <% If blnReqCustRegItemName3 = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="reqCustRegItemName3" value="False" <% If blnReqCustRegItemName3 = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
    </td>
   </tr>
   <tr>
    <td class="tableRow">Viewed in Member Profile:<br />
       <span class="smText">When enabled all members are able to view this item in member profiles. If disabled only Admins and Moderators can view this item in member profiles.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="viewCustRegItemName3" value="True" <% If blnViewCustRegItemName3 = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="viewCustRegItemName3" value="False" <% If blnViewCustRegItemName3 = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
    </td>
   </tr>
   
   <tr>
     <td colspan="2" class="tableLedger">Registration Form Privacy Notice</td>
    </tr>
    <tr>
     <td colspan="2" class="tableRow">
     This is shown on the bottom of the Forum Registration form and where Forum Members update their Forum Profile. This is to provide your Members with information
     on why you require Personal Data from Members and what you do with their Personal Data. This '<a href="https://ico.org.uk/for-organisations/guide-to-data-protection/privacy-notices-transparency-and-control/where-should-you-deliver-privacy-information-to-individuals/#justintime" target="_blank">Just in Time Notice</a>' is to help with Guide to the General Data Protection Regulation (GDPR) Compliance. 
     <br />HTML can be used for formatting to allow you to add a link to your Organisation's Full Privacy Notice, if available. 
     <br />
     <textarea name="regPrivacyNotice" id="regPrivacyNotice" rows="8" cols="100"><% = strRegPrivacyNotice %></textarea>
     <br />
     </td>
    </tr>

    <tr>
     <td colspan="2" class="tableLedger">Forum Rules</td>
    </tr>
    <tr>
     <td colspan="2" class="tableRow">
     Enter the Rules that have to be agreed to when Registering and Logging in to the Forum (HTML can be used for formatting)
     <br />
     <textarea name="registrationRules" id="registrationRules" rows="20" cols="100"><% = strRegistrationRules %></textarea>
     <br />
     </td>
    </tr>
    
    <tr align="center">
      <td colspan="2" class="tableBottomRow">
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Update Registration Settings" />
          <input type="reset" name="Reset" value="Reset Form" />      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
