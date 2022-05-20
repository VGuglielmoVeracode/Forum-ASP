<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
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


Dim lngTotalBlocked
Dim dtmLastBlocked
Dim strLastBlockedEmail
Dim strLastBlockedIP
Dim strStopForumSpamAgreement
      

'Read in the users details for the forum
blnStopForumSpam = BoolC(Request.Form("SfsEnable"))
strStopForumSpamAgreement = Request.Form("SfsTerms")
strStopForumSpamApiKey = Request.Form("SfsAPI")
blnStopForumSpamUsername = BoolC(Request.Form("SfsUsername"))


'Only enable if they have also checked that they agree to the rtems and conditions
If blnStopForumSpam AND NOT strStopForumSpamAgreement = "I Agree" Then 
	blnStopForumSpam = False
	strStopForumSpamAgreement = ""
End If



'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	Call addConfigurationItem("SFSpam", blnStopForumSpam)
	Call addConfigurationItem("SFSpam_API_Key", strStopForumSpamApiKey)
	Call addConfigurationItem("SFSpam_Agreement", strStopForumSpamAgreement)
	Call addConfigurationItem("SFSpam_Username", blnStopForumSpamUsername)
					
	'Update variables
	Application.Lock
	Application(strAppPrefix & "SFSpam") = CBool(blnStopForumSpam)
	Application(strAppPrefix & "strStopForumSpamApiKey") = strStopForumSpamApiKey
	Application(strAppPrefix & "blnStopForumSpamUsername") = blnStopForumSpamUsername
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
	blnStopForumSpam = CBool(getConfigurationItem("SFSpam", "bool"))
	strStopForumSpamApiKey = getConfigurationItem("SFSpam_API_Key", "string")
	
	strStopForumSpamAgreement = getConfigurationItem("SFSpam_Agreement", "string")
	
	
	'Read in stats
	lngTotalBlocked = CLng(getConfigurationItem("SFSpam_no_blocked", "numeric"))
	dtmLastBlocked = CDate(getConfigurationItem("SFSpam_last_block_date", "date"))
	strLastBlockedEmail = getConfigurationItem("SFSpam_last_block_email", "string")
	strLastBlockedIP = getConfigurationItem("SFSpam_last_block_IP", "string") 
	blnStopForumSpamUsername = CBool(getConfigurationItem("SFSpam_Username", "bool"))
	
End If


'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>StopForumSpam Settings</title>
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
  <h1>StopForumSpam Settings</h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    From here you can setup, configure and view statistics  for StopForumSpam.
    </p>
    <br />
    <br />
    <a href="http://www.stopforumspam.com" target="_blank">StopForumSpam</a> is a free Database List of Spammers that persist in abusing forums and blogs. <br />
    By enabling StopForumSpam in your fourm when new users registered their IP, Username, and Email Address is checked against the StopForumSpam Database and if found the registration is rejected.<br />
    <br />
</div>
<form action="admin_stopforumspam.asp<% = strQsSID1 %>" method="post" name="frmStopForumSpam" id="frmStopForumSpam">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td colspan="2" class="tableLedger">StopForumSpam Settings</td>
    </tr>
    <tr>
      <td width="50%" class="tableRow">Enable StopFourmSpam Database Checking:<br />
      <span class="smText">When enabled new Registrations will be checked against the StopFourmSpam Database.</span></td>
      <td width="50%" valign="top" class="tableRow">Yes
        <input type="radio" name="SfsEnable" value="True" <% If blnStopForumSpam = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;No
      <input type="radio" name="SfsEnable" value="False" <% If blnStopForumSpam = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
    <tr>
     <td class="tableRow">StopForumSpam API Key:<br />
       <span class="smText">Enter your <a href="http://www.stopforumspam.com/keys" target="_blank" class="smLink">StopForumSpam API Key</a> to allow information on Spammers on your forum to be submitted to StopForumSpam.<br />
       </span></td>
     <td valign="top" class="tableRow">
      <input type="text" name="SfsAPI" id="SfsAPI" maxlength="70" value="<% = strStopForumSpamApiKey %>" size="30"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td valign="top" class="tableRow">What to Check with StopFourmSpam Database:</td>
      <td valign="top" class="tableRow">
      	 <input type="radio" name="SfsUsername" value="False" <% If blnStopForumSpamUsername = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      Check for either IP Address or Email Address
      <br /><span class="smText">If either IP Address or Email Address is found in StopForumSpam's database the registrant will be blocked</span>
       <br />
       <br />
        <input type="radio" name="SfsUsername" value="True" <% If blnStopForumSpamUsername = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        Check the Username + IP Address + Email Address
        <br /><span class="smText">If a record is found in the StopForumsSpam database that match's ALL three of the above the registrant will be blocked</span>
       
         </td>
    </tr>
     <tr>
     <td colspan="2" class="tableRow"><strong>Terms and Conditions</strong><br />
       StopForumSpam is a free service that is made available within Web Wiz Forums at no extra charge. StopForumSpam and its services are PROVIDED "AS IS" WITHOUT WARRANTY OR GUARANTEES. This includes but not limited to, no support, no guarantees of accuracy and no guarantee of service availability. The 
StopForumSpam service maybe unavailable or discontinued at any time without notice.<br />
<br />
       <label for="SfsTerms"><input type="checkbox" name="SfsTerms" id="SfsTerms" value="I Agree" <% If strStopForumSpamAgreement = "I Agree" Then Response.Write(" checked=""checked""") %> />
       
       I agree to these Terms and Conditions</label></td>
     </tr>
    <tr align="center">
      <td colspan="2" class="tableBottomRow">
        <input type="hidden" name="postBack" value="true" />
        <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
        <input type="submit" name="Submit" value="Update StopForumSpam Settings" />
        <input type="reset" name="Reset" value="Reset Form" />      </td>
    </tr>
  </table>
</form>
<br />
<table align="center" cellpadding="4" cellspacing="1" class="tableBorder">
  <tr align="left">
    <td colspan="2" class="tableLedger">StopForumSpam Statistics</td>
  </tr>
  <tr>
    <td  height="2" align="left" class="tableRow">Total Blocked Registrations </td>
    <td height="2" valign="top" class="tableRow"><% = lngTotalBlocked %></td>
    </tr>
  <tr>
    <td  height="2" align="left" class="tableRow" valign="top">Last Blocked Registration</td>
    <td valign="top" class="tableRow"><% 
    	
    	'If there is a date display it
    	If lngTotalBlocked > 0 Then
    	
    	 	Response.Write(DateFormat(dtmLastBlocked) & "&nbsp;" &  strTxtAt & "&nbsp;" & TimeFormat(dtmLastBlocked))
    	 End If
    	 
    	  %></td>
  </tr>
  <tr>
    <td width="31%" align="left" class="tableRow">Username<span class="smText"></span></td>
    <td height="12" valign="top" class="tableRow"><% = getConfigurationItem("SFSpam_last_block_username", "string") %></td>
    </tr>
  <tr>
    <td align="left" class="tableRow">Email</td>
    <td height="12" valign="top" class="tableRow"><% = strLastBlockedEmail %></td>
    </tr>
  <tr>
    <td width="31%" align="left" class="tableRow">IP Address</td>
    <td height="12" valign="top" class="tableRow"><% = strLastBlockedIP %></td>
    </tr>

</table>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
