<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
<!--#include file="includes/ISO_country_list_inc.asp" -->
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




'Set the response buffer to true as we maybe redirecting
Response.Buffer = True


'Dimension variables
Dim intIsoLoop
Dim strBlockLIst

'Read in the details from the form
strBlockList = Request.Form("chkCountryCode")



'If the user is changing the email setup then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	
	Call addConfigurationItem("country_reg_block_list", strBlockList)

	
	
	'Update variables
	Application.Lock
	
	Application(strAppPrefix & "strCountryBlockRegList") = strCountryBlockRegList
	
	
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
	strCountryBlockRegList = getConfigurationItem("country_reg_block_list", "string")
End If


'Reset Server Objects
rsCommon.Close
Call closeDatabase()




%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Registration Country Block List</title>
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
<div align="center"><h1>Registration Country Block List</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br />
  <span class="text">From here you can block countries that you do not permit for new registrations.<br />
  <br />
  The service checks in real time the IP address (IPv4 and IPv6 supported) of the new member when they register and if their IP address belongs to a country that you have blocked they will be not be able to register on your forum.</span><br />
</div>
<br />
<form action="admin_registration_country_blocking.asp<% = strQsSID1 %>" method="post" name="frmCountryList" id="frmCountryList">
	
	
<table align="center" cellpadding="4" cellspacing="1" class="tableBorder" style="width:450px">
    <tr align="left">
      <td colspan="2" class="tableLedger">Registration Country Block List</td>
    </tr>
  <%
    
	
'Loop through ISO country array to display all the coutries in a drop down
For intIsoLoop = 1 to UBound(saryISOCountryCode,2)
     
     %>
    <tr class="tableRow">
      <td width="3%"><input type="checkbox" name="chkCountryCode" id="CountryCode" value="<% = saryISOCountryCode(0,intIsoLoop) %>" <% If InStr(strCountryBlockRegList, saryISOCountryCode(0,intIsoLoop)) Then Response.Write("checked=""checked"" ") %>/></td>
      <td nowrap="nowrap"><% = saryISOCountryCode(1,intIsoLoop) %></td>
    </tr><%
     
Next
		



%>
    <tr align="center">
      <td colspan="3" valign="top" class="tableRow">
      	 <input type="hidden" name="postBack" value="true" />
         <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
      	<input type="submit" name="Submit" value="Update Registration Country Block List" />
      </td>
    </tr>
  </table>	
	
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
