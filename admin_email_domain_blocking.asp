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






'Set the response buffer to true as we maybe redirecting
Response.Buffer = True


'Dimension variables
Dim rsSelectForum		'Holds the db recordset
Dim strBlockEmail		'Holds the Email address to block
Dim strBlockedEmailList		'Holds the Email addresses in the blocked list
Dim lngBlockedEmailID		'Holds the ID number of the blcoked db record
Dim laryCheckedEmailAddrID	'Holds the array of Email addresses to be ditched





'Run through till all checked IP addresses are deleted
For each laryCheckedEmailAddrID in Request.Form("chkDelete")

	'Here we use the less effiecient ADO to delete from the database this way we can throw in a requery while we wait for slow old MS Access to catch up

	'Delete the IP address from the database	
	strSQL = "SELECT * FROM " & strDbTable & "BanList WHERE " & strDbTable & "BanList.Ban_ID="  & laryCheckedEmailAddrID & ";"
	
	With rsCommon		
		'Set the cursor	type property of the record set	to Forward Only
		.CursorType = 0
		
		'Set the Lock Type for the records so that the record set is only locked when it is updated
		.LockType = 3
		
		'Query the database
		.Open strSQL, adoCon
		
		'Delete from the db
		If NOT .EOF Then .Delete
		
		'Requery
		.Requery
		
		'Close the recordset
		.Close
	End With
	
Next




'Run through till all checked Email addresses are deleted
For each laryCheckedEmailAddrID in Request.Form("chkDelete")

	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))


	'Here we use the less effiecient ADO to delete from the database this way we can throw in a requery while we wait for slow old MS Access to catch up

	'Delete the Email address from the database	
	strSQL = "SELECT * FROM " & strDbTable & "BanList WHERE " & strDbTable & "BanList.Ban_ID="  & laryCheckedEmailAddrID & ";"
	
	With rsCommon		
		'Set the cursor	type property of the record set	to Forward Only
		.CursorType = 0
		
		'Set the Lock Type for the records so that the record set is only locked when it is updated
		.LockType = 3
		
		'Query the database
		.Open strSQL, adoCon
		
		'Delete from the db
		If NOT .EOF Then .Delete
		
		'Requery
		.Requery
		
		'Close the recordset
		.Close
	End With
	
Next



'Read in all the blocked Email address from the database

'Initalise the strSQL variable with an SQL statement to query the database 
strSQL = "SELECT " & strDbTable & "BanList.Ban_ID, " & strDbTable & "BanList.Email " & _
"FROM " & strDbTable & "BanList " & _
"WHERE " & strDbTable & "BanList.Email Is Not Null;"

'Set the cursor	type property of the record set	to Forward Only
rsCommon.CursorType = 0

'Set the Lock Type for the records so that the record set is only locked when it is updated
rsCommon.LockType = 3

'Query the database
rsCommon.Open strSQL, adoCon



'If this is a post back then  update the database
If Request.Form("Email") <> ""  AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))

	'Read in the Email address to block
	strBlockEmail = Trim(Mid(Request.Form("Email"), 1, 30))

	'Update the recordset
	With rsCommon
	
		.AddNew

		'Update	the recorset
		.Fields("Email") = strBlockEmail

		'Update db
		.Update

		'Re-run the query as access needs time to catch up
		.ReQuery

	End With
End If

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Email Address Blocking</title>
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
<script language="JavaScript" type="text/javascript">

//Function to check form is filled in correctly before submitting
function CheckForm () {

	//Check for an Email address
	if (document.frmEmailadd.Email.value==""){
		alert("Please enter an Email address or domain");
		document.frmEmailadd.Email.focus();
		return false;
	}

	return true;
}
</script>
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center"><h1>Email Address Blocking</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br />
  <span class="text">From here you can Block email addresses or email domains.<br />
  <br />
  This function is only really useful if you have email activation enabled as it will prevent anyone with a blocked email address registering on the forum with that email address.<br />
</span><br />
  <br />
</div>
<br />
<form action="admin_email_domain_blocking.asp<% = strQsSID1 %>" method="post" name="frmIPList" id="frmIPList">
  <table align="center" cellpadding="4" cellspacing="1" class="tableBorder" style="width:400px">
    <tr align="left">
      <td colspan="2" class="tableLedger">Blocked Email Address/Domain List</td>
    </tr><%
'Display the email blcok list
If rsCommon.EOF Then 
		
	'Disply no entires forun
	Response.Write("<td colspan=""2"" align=""center"" class=""tableRow"">You have no blocked Email address</td>")
	
'Else disply the IP block list
Else
	
	'Loop through the recordset
	Do While NOT rsCommon.EOF
	
     		'Read in the details
     		lngBlockedEmailID = CLng(rsCommon("Ban_ID"))
		strBlockedEmailList = rsCommon("Email")
     
     %>
    <tr>
      <td width="3%" class="tableRow"><input type="checkbox" name="chkDelete" value="<% = lngBlockedEmailID %>"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
      <td class="tableRow"><% = strBlockedEmailList %></td>
    </tr><%
     

		'Move to the next record in the recordset
		rsCommon.MoveNext
	Loop
End If

'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
    <tr align="center">
      <td colspan="2" valign="top" class="tableRow">
      	<input type="hidden" name="formID" id="formID1" value="<% = getSessionItem("KEY") %>" />
      	<input type="submit" name="Submit" value="Remove Email Address or Email Domain" />
      </td>
    </tr>
  </table>
  <br />
</form>
<form action="admin_email_domain_blocking.asp<% = strQsSID1 %>" method="post" name="frmEmailadd" id="frmEmailadd" onsubmit="return CheckForm();">
<table align="center" cellpadding="4" cellspacing="1" class="tableBorder" style="width: 400px">
  <tr align="left">
    <td colspan="2" class="tableLedger">Block Email Address or Domain</td>
  </tr>
  <tr class="tableRow">
    <td colspan="2" align="center" class="smText">The * wildcard character can be used to block email domains. <br />
      eg. To block users with a yahoo.com email address you would use. eg. *@yahoo.com</td>
  </tr>
  <tr>
    <td align="right" class="tableRow">Email Address or domain:</td>
    <td class="tableRow"><input name="Email" type="text" id="Email" size="25" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
  </tr>
  <tr align="center">
    <td colspan="2" valign="top" class="tableBottomRow">
    	<input type="hidden" name="formID" id="formID2" value="<% = getSessionItem("KEY") %>" />
    	<input type="submit" name="Submit2" value="Block Email Address or Range" />
    </td>
  </tr>
</table>
<br />
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
