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
blnChatRoom = BoolC(Request.Form("chatRoom"))


'If the user is changing tthe colours then update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	Call addConfigurationItem("Chat_room", blnChatRoom)
	
	
	'Update variables
	Application.Lock
	
	Application(strAppPrefix & "blnChatRoom") = CBool(blnChatRoom)
	
	
	
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
	blnChatRoom = CBool(getConfigurationItem("Chat_room", "bool"))
	
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
		rsCommon.Fields("Chat_Room") = BoolC(Request.Form("ChatGroup" & rsCommon("Group_ID")))

		'Update the database
		rsCommon.Update
   
		'Move to next record in rs
		rsCommon.MoveNext
	Loop
	
End If



	
	
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Chat Room Settings</title>
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
  <h1>Chat Room Settings </h1> 
  <br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    From here you can configure the Chat Room settings. <br />
    <br />
</div>
<form action="admin_chat_room_settings.asp<% = strQsSID1 %>" method="post" name="frmChat" id="frmChat">
  <table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
      <td colspan="2" class="tableLedger">Chat Room</td>
    </tr>
    <tr>
      <td width="57%" class="tableRow">Chat Room:<br />
      <span class="smText">If this is disabled your members will no longer be able to use the Chat Room.</span></td>
      <td width="43%" valign="top" class="tableRow">Yes
        <input type="radio" name="chatRoom" value="True" <% If blnChatRoom = True Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
        &nbsp;&nbsp;No
      <input type="radio" name="chatRoom" value="False" <% If blnChatRoom = False Then Response.Write "checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />      </td>
    </tr>
    <tr>
     <td  height="2" colspan="2" align="left" class="tableRow">Select Which Groups are Permitted to use the Chat Room
      <table width="100%"  border="0" cellspacing="1" cellpadding="1">
      <tr class="tableRow"> 
       <td width="1%" align="right"><input type="checkbox" name="chkAllChatGroup" id="chkAllChatGroup" onclick="checkAll('ChatGroup');" /></td>
       <td width="99%"><strong>Check All</strong></td>
      </tr><%
 
'Query the database
rsCommon.MoveFirst        
	
'Loop through groups
Do While NOT rsCommon.EOF
	
	'If not guest group display if they can send PM's
	If rsCommon("Group_ID") <> 2 Then
		Response.Write(vbCrLf & "   <tr class=""tableRow""> " & _
		vbCrLf & "    <td width=""1%"" align=""right""><input type=""checkbox"" name=""ChatGroup" & rsCommon("Group_ID") & """ id=""ChatGroup" & rsCommon("Group_ID") & """ value=""true""")
		If  CBool(rsCommon("Chat_Room")) Then Response.Write(" checked")
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
