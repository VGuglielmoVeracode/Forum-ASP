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
blnPrivateMessages = BoolC(Request.Form("privateMsg"))
intPmInbox = IntC(Request.Form("pmInboxSize"))
intPmOutbox = IntC(Request.Form("pmOutboxSize"))
intPmFlood = IntC(Request.Form("PmFlood"))
strPMoverAction = Request.Form("PmOverAction")
blnPmFlashFiles = BoolC(Request.Form("flash"))
blnPmYouTube = BoolC(Request.Form("YouTube"))
blnPmIgnoreSpamFilter = BoolC(Request.Form("SpamFilters"))
	

'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	Call addConfigurationItem("Private_msg", blnPrivateMessages)
	Call addConfigurationItem("PM_inbox", intPmInbox)
	Call addConfigurationItem("PM_outbox", intPmOutbox)
	Call addConfigurationItem("PM_Flood", intPmFlood)
	Call addConfigurationItem("PM_overusage_action", strPMoverAction)				
	Call addConfigurationItem("PM_Flash", blnPmFlashFiles)
	Call addConfigurationItem("PM_YouTube", blnPmYouTube)
	Call addConfigurationItem("PM_spam_ignore", blnPmIgnoreSpamFilter)
	
	'Update variables
	Application.Lock
	
	Application(strAppPrefix & "blnPrivateMessages") = CBool(blnPrivateMessages)
	Application(strAppPrefix & "intPmInbox") = intPmInbox
	Application(strAppPrefix & "intPmOutbox") = intPmOutbox
	Application(strAppPrefix & "intPmFlood") = CInt(intPmFlood)
	Application(strAppPrefix & "strPMoverAction") = strPMoverAction
	Application(strAppPrefix & "blnPmFlashFiles") = CBool(blnPmFlashFiles)
	Application(strAppPrefix & "blnPmYouTube") = CBool(blnPmYouTube)
	Application(strAppPrefix & "blnPmIgnoreSpamFilter") = CBool(blnPmIgnoreSpamFilter)
	
	
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

'Read in the forum colours from the database
If NOT rsCommon.EOF Then
	
	'Place into an array for performance
	saryConfiguration = rsCommon.GetRows()
	
	'Read in the colour info from the database
	blnPrivateMessages = CBool(getConfigurationItem("Private_msg", "bool"))
	intPmInbox = CInt(getConfigurationItem("PM_inbox", "numeric"))
	intPmOutbox = CInt(getConfigurationItem("PM_outbox", "numeric"))
	intPmFlood = Cint(getConfigurationItem("PM_Flood", "numeric"))
	strPMoverAction = getConfigurationItem("PM_overusage_action", "string")
	blnPmFlashFiles = CBool(getConfigurationItem("PM_Flash", "bool"))
	blnPmYouTube = CBool(getConfigurationItem("PM_YouTube", "bool"))
	blnPmIgnoreSpamFilter = CBool(getConfigurationItem("PM_spam_ignore", "bool"))
End If



rsCommon.Close



'Initalise the strSQL variable with an SQL statement to query the database
'WHERE cluse added to get round bug in myODBC which won't run an ADO update unless you have a WHERE cluase
strSQL = "SELECT " & strDbTable & "Group.* " & _
"FROM " & strDbTable & "Group " & _
"WHERE " & strDbTable & "Group.Group_ID > 0 " & _
"ORDER BY " & strDbTable & "Group.Group_ID ASC;"
	
'Set the cursor type property of the record set to Forward Only
rsCommon.CursorType = 0

'Set the Lock Type for the records so that the record set is only locked when it is updated
rsCommon.LockType = 3

'Query the database
rsCommon.Open strSQL, adoCon




'Update the db with who can PM
If Request.Form("postBack") Then
	
	'Loop through cats
	Do While NOT rsCommon.EOF
	
		'Update the recordset
		rsCommon.Fields("Private_Messenger") = BoolC(Request.Form("PmGroup" & rsCommon("Group_ID")))

		'Update the database
		rsCommon.Update
   
		'Move to next record in rs
		rsCommon.MoveNext
	Loop
	
End If




	
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Private Messaging Settings</title>
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
  <h1>Private Messaging Settings </h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    From here you can configure the Private Messaging system. <br />
    <br />
</div>
<form action="admin_pm_configure.asp<% = strQsSID1 %>" method="post" name="frmPMCal" id="frmPMCal">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td colspan="2" class="tableLedger">Private Messenger</td>
    </tr>
    <tr>
      <td width="57%" class="tableRow">Private Messenger:<br />
      <span class="smText">If this is disabled your members will no longer be able to use the Private Messenger to send and receive Private Messages.</span></td>
      <td width="43%" valign="top" class="tableRow">Yes
        <input type="radio" name="privateMsg" value="True" <% If blnPrivateMessages = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;No
      <input type="radio" value="False" <% If blnPrivateMessages = False Then Response.Write "checked" %> name="privateMsg"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
    <tr>
     <td class="tableRow">Private Messenger Inbox Storage Size:<br />
      <span class="smText">This is the number of Private Messages a member can have in there 'inbox' at any one time.</span></td>
     <td valign="top" class="tableRow"><select name="pmInboxSize" id="pmInboxSize"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
      <option<% If intPmInbox = 5 Then Response.Write(" selected") %>>5</option>
      <option<% If intPmInbox = 10 Then Response.Write(" selected") %>>10</option>
      <option<% If intPmInbox = 15 Then Response.Write(" selected") %>>15</option>
      <option<% If intPmInbox = 20 Then Response.Write(" selected") %>>20</option>
      <option<% If intPmInbox = 25 Then Response.Write(" selected") %>>25</option>
      <option<% If intPmInbox = 30 Then Response.Write(" selected") %>>30</option>
      <option<% If intPmInbox = 35 Then Response.Write(" selected") %>>35</option>
      <option<% If intPmInbox = 40 Then Response.Write(" selected") %>>40</option>
      <option<% If intPmInbox = 45 Then Response.Write(" selected") %>>45</option>
      <option<% If intPmInbox = 50 Then Response.Write(" selected") %>>50</option>
      <option<% If intPmInbox = 60 Then Response.Write(" selected") %>>60</option>
      <option<% If intPmInbox = 70 Then Response.Write(" selected") %>>70</option>
      <option<% If intPmInbox = 80 Then Response.Write(" selected") %>>80</option>
      <option<% If intPmInbox = 90 Then Response.Write(" selected") %>>90</option>
      <option<% If intPmInbox = 100 Then Response.Write(" selected") %>>100</option>
      <option<% If intPmInbox = 150 Then Response.Write(" selected") %>>150</option>
      <option<% If intPmInbox = 200 Then Response.Write(" selected") %>>200</option>
      <option<% If intPmInbox = 250 Then Response.Write(" selected") %>>250</option>
      <option<% If intPmInbox = 500 Then Response.Write(" selected") %>>500</option>
      <option<% If intPmInbox = 1000 Then Response.Write(" selected") %>>1000</option>
      <option<% If intPmInbox = 2000 Then Response.Write(" selected") %>>2000</option>
      <option<% If intPmInbox = 5000 Then Response.Write(" selected") %>>5000</option>
      <option<% If intPmInbox = 10000 Then Response.Write(" selected") %>>10000</option>
     </select>
     </td>
    </tr>
     <tr>
      <td class="tableRow">Full Private Message Inbox Action:<br />
        <span class="smText">This is the action to take if the member has received more Private Messages than is allowed.</span></td>
      <td valign="top" class="tableRow"><select name="PmOverAction" id="PmOverAction"<% If blnDemoMode  Then Response.Write(" disabled=""disabled""") %>>
       <option value="delete"<% If strPMoverAction = "delete" Then Response.Write(" selected") %>>Delete members oldest Private Messages</option>
       <option value="block"<% If strPMoverAction = "block" Then Response.Write(" selected") %>>Member will not receive any new Private Messages</option>
      </select>
     </td>
    </tr>
     <tr>
     <td class="tableRow">Private Messenger Outbox Storage Size:<br />
      <span class="smText">This is the number of Private Messages a member can have in there 'outbox' at any one time.</span></td>
     <td valign="top" class="tableRow"><select name="pmOutboxSize" id="pmOutboxSize"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
      <option<% If intPmOutbox = 5 Then Response.Write(" selected") %>>5</option>
      <option<% If intPmOutbox = 10 Then Response.Write(" selected") %>>10</option>
      <option<% If intPmOutbox = 15 Then Response.Write(" selected") %>>15</option>
      <option<% If intPmOutbox = 20 Then Response.Write(" selected") %>>20</option>
      <option<% If intPmOutbox = 25 Then Response.Write(" selected") %>>25</option>
      <option<% If intPmOutbox = 30 Then Response.Write(" selected") %>>30</option>
      <option<% If intPmOutbox = 35 Then Response.Write(" selected") %>>35</option>
      <option<% If intPmOutbox = 40 Then Response.Write(" selected") %>>40</option>
      <option<% If intPmOutbox = 45 Then Response.Write(" selected") %>>45</option>
      <option<% If intPmOutbox = 50 Then Response.Write(" selected") %>>50</option>
      <option<% If intPmOutbox = 60 Then Response.Write(" selected") %>>60</option>
      <option<% If intPmOutbox = 70 Then Response.Write(" selected") %>>70</option>
      <option<% If intPmOutbox = 80 Then Response.Write(" selected") %>>80</option>
      <option<% If intPmOutbox = 90 Then Response.Write(" selected") %>>90</option>
      <option<% If intPmOutbox = 100 Then Response.Write(" selected") %>>100</option>
      <option<% If intPmOutbox = 150 Then Response.Write(" selected") %>>150</option>
      <option<% If intPmOutbox = 200 Then Response.Write(" selected") %>>200</option>
      <option<% If intPmOutbox = 250 Then Response.Write(" selected") %>>250</option>
      <option<% If intPmOutbox = 500 Then Response.Write(" selected") %>>500</option>
      <option<% If intPmOutbox = 1000 Then Response.Write(" selected") %>>1000</option>
      <option<% If intPmOutbox = 2000 Then Response.Write(" selected") %>>2000</option>
      <option<% If intPmOutbox = 5000 Then Response.Write(" selected") %>>5000</option>
      <option<% If intPmOutbox = 10000 Then Response.Write(" selected") %>>10000</option>
     </select>
     
     </td>
    </tr>
    <tr>
      <td class="tableRow">Private Messager Flood Control:<br />
        <span class="smText">This is the number of Private Messages a member can send within an hour. This prevents a member sending 100's of spam Private Messages to other members. </span></td>
      <td valign="top" class="tableRow"><select name="PmFlood" id="PmFlood"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
       <option<% If intPmFlood = 1 Then Response.Write(" selected") %>>1</option>
       <option<% If intPmFlood = 2 Then Response.Write(" selected") %>>2</option>
       <option<% If intPmFlood = 3 Then Response.Write(" selected") %>>3</option>
       <option<% If intPmFlood = 4 Then Response.Write(" selected") %>>4</option>
       <option<% If intPmFlood = 5 Then Response.Write(" selected") %>>5</option>
       <option<% If intPmFlood = 6 Then Response.Write(" selected") %>>6</option>
       <option<% If intPmFlood = 7 Then Response.Write(" selected") %>>7</option>
       <option<% If intPmFlood = 8 Then Response.Write(" selected") %>>8</option>
       <option<% If intPmFlood = 9 Then Response.Write(" selected") %>>9</option>
       <option<% If intPmFlood = 10 Then Response.Write(" selected") %>>10</option>
       <option<% If intPmFlood = 15 Then Response.Write(" selected") %>>15</option>
       <option<% If intPmFlood = 20 Then Response.Write(" selected") %>>20</option>
       <option<% If intPmFlood = 25 Then Response.Write(" selected") %>>25</option>
       <option<% If intPmFlood = 30 Then Response.Write(" selected") %>>30</option>
       <option<% If intPmFlood = 35 Then Response.Write(" selected") %>>35</option>
       <option<% If intPmFlood = 40 Then Response.Write(" selected") %>>40</option>
       <option<% If intPmFlood = 50 Then Response.Write(" selected") %>>50</option>
       <option<% If intPmFlood = 75 Then Response.Write(" selected") %>>75</option>
       <option<% If intPmFlood = 100 Then Response.Write(" selected") %>>100</option>
       <option<% If intPmFlood = 150 Then Response.Write(" selected") %>>150</option>
       <option<% If intPmFlood = 200 Then Response.Write(" selected") %>>200</option>
      </select>
       per hour
     </td>
    </tr>
    <tr>
     <td class="tableRow">Adobe Flash:<br />
       <span class="smText">If you enable this then users will be able to display Flash content in Private Messages using Forum BBcode [FLASH]file name here[/FLASH]</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="flash" value="True" <% If blnPmFlashFiles = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="flash" value="False" <% If blnPmFlashFiles = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td></tr>
    <td class="tableRow">YouTube Movies:<br />
       <span class="smText">If you enable this then users will be able to display YouTube movies in Private Messages using Forum BBcode [TUBE]file name here[/TUBE]</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="YouTube" value="True" <% If blnPmYouTube = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="YouTube" value="False" <% If blnPmYouTube = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td></tr>
    <tr>
    <td class="tableRow">Ignore SPAM Filters:<br />
       <span class="smText">Allows you to disable the <a href="admin_spam_filter_configure.asp<% = strQsSID1 %>" class="smLink">SPAM Filters</a> so that Private Messages are not checked for SPAM by the SPAM Filters.</span></td>
     <td valign="top" class="tableRow">Yes
      <input type="radio" name="SpamFilters" value="True" <% If blnPmIgnoreSpamFilter = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      &nbsp;&nbsp;No
      <input type="radio" name="SpamFilters" value="False" <% If blnPmIgnoreSpamFilter = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
    <tr>
     <td  height="2" colspan="2" align="left" class="tableRow">Select Which Groups are Permitted to use the Private Messaging system
      <table width="100%"  border="0" cellspacing="1" cellpadding="1">
      <tr class="tableRow"> 
       <td width="1%" align="right"><input type="checkbox" name="chkAllPmGroup" id="chkAllPmGroup" onclick="checkAll('PmGroup');" /></td>
       <td width="99%"><strong>Check All</strong></td>
      </tr><%
 
'Query the database
rsCommon.MoveFirst        
	
'Loop through groups
Do While NOT rsCommon.EOF
	
	'If not guest group display if they can send PM's
	If rsCommon("Group_ID") <> 2 Then
		Response.Write(vbCrLf & "   <tr class=""tableRow""> " & _
		vbCrLf & "    <td width=""1%"" align=""right""><input type=""checkbox"" name=""PmGroup" & rsCommon("Group_ID") & """ id=""PmGroup" & rsCommon("Group_ID") & """ value=""true""")
		If  CBool(rsCommon("Private_Messenger")) Then Response.Write(" checked")
		If blnDemoMode Then Response.Write(" disabled=""disabled""")
		Response.Write(" /></td>" & _
		vbCrLf & "    <td width=""99%"">" & rsCommon("Name") & "</td>" & _
		vbCrLf & "   </tr>")
	End If
   
	'Move to next record in rs
	rsCommon.MoveNext
Loop

 %>
       </table>     </td>
     </tr>
    <tr>
    <tr align="center">
      <td colspan="2" class="tableBottomRow">
          <input type="hidden" name="postBack" value="true" />
          <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
          <input type="submit" name="Submit" value="Update Settings" />
          <input type="reset" name="Reset" value="Reset Form" />      </td>
    </tr>
  </table>
</form>
<br /><%

'Reset Server Objects
rsCommon.Close
Call closeDatabase()

%>
<!-- #include file="includes/admin_footer_inc.asp" -->
