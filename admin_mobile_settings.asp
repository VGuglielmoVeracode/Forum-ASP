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
blnMobileView = BoolC(Request.Form("MobileView"))
strHeaderMobile = Request.Form("HeaderMobile")
strFooterMobile = Request.Form("FooterMobile")
blnShowMobileHeaderFooter = BoolC(Request.Form("EnableHeaderFooter"))



'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	Call addConfigurationItem("Mobile_View", blnMobileView)
	Call addConfigurationItem("Header_mobile", strHeaderMobile)
	Call addConfigurationItem("Footer_mobile", strFooterMobile)
	Call addConfigurationItem("Show_mobile_header_footer", blnShowMobileHeaderFooter)
	
	
					
	'Update variables
	Application.Lock
	Application(strAppPrefix & "blnMobileView") = CBool(blnMobileView)
	Application(strAppPrefix & "strHeaderMobile") = strHeaderMobile
	Application(strAppPrefix & "strFooterMobile") = strFooterMobile
	Application(strAppPrefix & "blnShowMobileHeaderFooter") = CBool(blnShowMobileHeaderFooter)
	
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

'Read in the forum from the database
If NOT rsCommon.EOF Then
	
	'Place into an array for performance
	saryConfiguration = rsCommon.GetRows()

	'Read in the colour info from the database
	blnMobileView = CBool(getConfigurationItem("Mobile_View", "bool"))
	strHeaderMobile	= getConfigurationItem("Header_mobile", "string")
	strFooterMobile = getConfigurationItem("Footer_mobile", "string")
	blnShowMobileHeaderFooter = CBool(getConfigurationItem("Show_mobile_header_footer", "bool"))
End If



'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Mobile Device Settings</title>
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
  <h1>Mobile Device Settings</h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    From here you can configure your forum for Mobile Devices such as SmartPhones and Tablets.<br />
    <br />
</div>
<form action="admin_mobile_settings.asp<% = strQsSID1 %>" method="post" name="frmConfiguration" id="frmConfiguration" onsubmit="return CheckForm();">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td colspan="2" class="tableLedger">Mobile Device Settings</td>
    </tr>
    <tr>
     <td class="tableRow" width="50%">Enable Mobile Optimised View:<br />
      <span class="smText">When enabled a Mobile Optimised View will be displayed to Mobile Devices such as SmartPhones and Tablets.</span></td>
     <td valign="top" class="tableRow" width="50%">Yes
      <input type="radio" name="MobileView" value="True" <% If blnMobileView = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="MobileView" value="False" <% If blnMobileView = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
    <tr>
     <td colspan="2" class="tableLedger">Mobile View Header</td>
    </tr>
    
    <tr>
     <td class="tableRow" width="50%">Enable Mobile Custom Header and Footer:<br />
      <span class="smText">This will enable the Custom Header and Footer set below.</span></td>
     <td valign="top" class="tableRow" width="50%">Yes
      <input type="radio" name="EnableHeaderFooter" value="True" <% If blnShowMobileHeaderFooter = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="EnableHeaderFooter" value="False" <% If blnShowMobileHeaderFooter = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    
    <tr>
     <td colspan="2" class="tableRow">
     If you have a custom header you wish to use when viewing using the Mobile Optimised View, such as a mobile website template enter the HTML below that you want displayed at the top of your forum*
     <br />
     <textarea name="HeaderMobile" id="HeaderMobile" rows="7" cols="100"><% = strHeaderMobile %></textarea>
     </td>
    </tr>
    
    <tr>
     <td colspan="2" class="tableLedger">Mobile View Footer</td>
    </tr>
    <tr>
     <td colspan="2" class="tableRow">
     If you have a custom footer you wish to use when viewing using the Mobile Optimised View, such as a mobile website template enter the HTML below that you want displayed at the bottom of your forum*
     <br />
     <textarea name="FooterMobile" id="FooterMobile" rows="7" cols="100"><% = strFooterMobile %></textarea>
     <br />
     <span class="smText">*Custom Headers or Footers may not be compatible and may require modification to correctly display or prevent conflicts with page elements and features of Web Wiz Forums pages.</span>
     </td>
    </tr>
    
    <tr align="center">
      <td colspan="2" class="tableBottomRow">
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Update Mobile Settings" />
          <input type="reset" name="Reset" value="Reset Form" />      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
