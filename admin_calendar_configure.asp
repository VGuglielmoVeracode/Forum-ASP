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
blnCalendar = BoolC(Request.Form("calendar"))
blnDisplayBirthdays = BoolC(Request.Form("showBirthdays"))


'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	Call addConfigurationItem("Calendar", blnCalendar)
	Call addConfigurationItem("Show_birthdays", blnDisplayBirthdays)
					
	
	'Update variables
	Application.Lock
	Application(strAppPrefix & "blnCalendar") = CBool(blnCalendar)	
	Application(strAppPrefix & "blnDisplayBirthdays") = CBool(blnDisplayBirthdays)
	Application(strAppPrefix & "blnConfigurationSet") = false
	Application.UnLock
End If


'Initialise the SQL variable with an SQL statement to get the configuration details from the database
strSQL = "SELECT " & strDbTable & "SetupOptions.Option_Item, " & strDbTable & "SetupOptions.Option_Value " & _
"FROM " & strDbTable & "SetupOptions" &  strDBNoLock & " " & _
"ORDER BY " & strDbTable & "SetupOptions.Option_Item ASC;"

	
'Query the database
rsCommon.Open strSQL, adoCon

'Read in the forum colours from the database
If NOT rsCommon.EOF Then
	
	'Place into an array for performance
	saryConfiguration = rsCommon.GetRows()
	
	'Read in the colour info from the database
	blnCalendar = CBool(getConfigurationItem("Calendar", "bool"))
	blnDisplayBirthdays = CBool(getConfigurationItem("Show_birthdays", "bool"))
End If


'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Events Calendar Settings </title>
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
  <h1>Events Calendar Settings </h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    From here you can configure the Events Calendar. <br />
    <br />
    <table border="0" cellpadding="4" cellspacing="1" class="tableBorder">
     <tr>
      <td align="center" class="tableLedger">Events Calendar System</td>
     </tr>
     <tr>
      <td align="left" class="tableRow">The Events Calendar is very tightly integrated into the forum allowing members to create new Topics which can be given an Event Date. Those Forum Topics that are given Event Dates are displayed within the Calendar system.<br />
      <br />
       This has a number of advantages, including allowing members to post replies to Calendar Events, and only displaying Calendar Events to those who have access to the forums in which the Calendar Event was created, allowing for  Private Events to be entered into the Calendar System.<br />
      <br />
      To allow members to create Calendar Events you need to grant members permission to 'Create Events' within forums, using the Forum Permission system.</td>
     </tr>
    </table>
    <br />
</div>
<form action="admin_calendar_configure.asp<% = strQsSID1 %>" method="post" name="frmPMCal" id="frmPMCal">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td colspan="2" class="tableLedger">Events Calendar </td>
    </tr>
    <tr>
     <td width="57%" class="tableRow">Events Calendar:<br />
       <span class="smText">This allows enables the 'Events Calendar' within your forum, to display community events.</span></td>
     <td width="43%" valign="top" class="tableRow">Yes
      <input type="radio" name="calendar" value="True" <% If blnCalendar = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" value="False" <% If blnCalendar = False Then Response.Write "checked" %> name="calendar"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
      <td width="57%" class="tableRow">Show Members Birthdays:<br />
        <span class="smText">If enabled this will show Members Birthdays within the Calendar System.</span></td>
      <td width="43%" valign="top" class="tableRow">Yes
       <input type="radio" name="showBirthdays" value="True" <% If blnDisplayBirthdays = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;No
        <input type="radio" name="showBirthdays" value="False" <% If blnDisplayBirthdays = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
    <tr align="center">
      <td colspan="2" class="tableBottomRow">
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Update Settings" />
          <input type="reset" name="Reset" value="Reset Form" />      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
