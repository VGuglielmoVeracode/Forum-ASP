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
blnWebWizNewsPad = BoolC(Request.Form("NewsPad"))
strWebWizNewsPadURL = Request.Form("NewsPadURL")





'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
		
	Call addConfigurationItem("NewsPad", blnWebWizNewsPad)
	Call addConfigurationItem("NewsPad_URL", strWebWizNewsPadURL)

		
					
	'Update variables
	Application.Lock
	Application(strAppPrefix & "blnWebWizNewsPad") = CBool(blnWebWizNewsPad)
	Application(strAppPrefix & "strWebWizNewsPadURL") = strWebWizNewsPadURL
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
	blnWebWizNewsPad = CBool(getConfigurationItem("NewsPad", "bool"))
	strWebWizNewsPadURL = getConfigurationItem("NewsPad_URL", "string")
End If


'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Web Wiz NewsPad Settings</title>
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
  <h1> Web Wiz NewsPad Settings</h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    <br />
    <table border="0" cellpadding="4" cellspacing="1" class="tableBorder">
     <tr>
      <td align="center" class="tableLedger">Web Wiz NewsPad</td>
     </tr>
     <tr>
      <td align="left" class="tableRow"><a href="http://www.webwiznewspad.com" target="_blank">Web Wiz NewsPad</a> is a eNewsletters and Blog Tool allowing you to created eNewsletters and Blog Posts for your website. Web Wiz NewsPad can be used as a standalone tool or you can integrate it with Web Wiz Forums.<br />
<br />
<strong>Mass Email Web Wiz Forums Members</strong><br />
Web Wiz NewsPad tightly integrates with Web Wiz Forums allowing you to send mass emails and eNewsletters to your Web Wiz Forums Members. With timed batch sending NewsPad enables you to send emails to thousands of your members without overloading your mail server or degrading your website's performance.<br />
<br />
For a full list of features of Web Wiz NewsPad please see; <a href="http://www.webwiznewspad.com" target="_blank">Web Wiz NewsPad</a>.</td>
     </tr>
    </table>
    <br />
</div>
<form action="admin_newspad_configure.asp<% = strQsSID1 %>" method="post" name="frmConfiguration" id="frmConfiguration">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
     <td colspan="2" class="tableLedger">Web Wiz NewsPad Integration </td>
     </tr>
    <tr>
     <td width="57%" class="tableRow"><a href="http://www.webwiznewspad.com" target="_blank">Web Wiz NewsPad</a> Integration<br />
     <span class="smText"> By enabling Web Wiz NewsPad integration, your members can select when registering or in their forum profile to signup to eNewsletters and Blog Posts sent using Web Wiz NewsPad </span></td>
     <td width="43%" valign="top" class="tableRow">Yes
      <input type="radio" name="NewsPad" id="NewsPad" value="True" <% If blnWebWizNewsPad = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
&nbsp;&nbsp;No
<input type="radio" value="False" <% If blnWebWizNewsPad = False Then Response.Write "checked" %> name="NewsPad" id="NewsPad"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td class="tableRow">Web Wiz NewsPad URL <br />
      <span class="smText">If you type in the URL to NewsPad it will use the NewsPad RSS Feed to display links to your NewsPad Blog Posts on your forums index page. </span><br />
      <span class="smText">eg. http://www.mywebsite.com/NewsPad</span></td>
     <td valign="top" class="tableRow"><input name="NewsPadURL" type="text" id="NewsPadURL" value="<% = strWebWizNewsPadURL %>" size="35" maxlength="75"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr align="center">
      <td colspan="2" class="tableBottomRow">
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Update NewsPad Settings" />
          <input type="reset" name="Reset" value="Reset Form" />      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
