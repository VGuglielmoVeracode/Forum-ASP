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
Dim rsSelectForum	'Holds the db recordset
Dim strBlockIP		'Holds the IP address to block
Dim strBlockedIPList	'Holds the IP addresses in the blocked list
Dim lngBlockedIPID	'Holds the ID number of the blcoked db record
Dim laryCheckedIPAddrID	'Holds the array of IP addresses to be ditched
Dim strReason		'Holds the reason for the IP ban


'Get IP address to block
strBlockIP = Trim(Mid(Request.QueryString("IP"), 1, 30))


'Run through till all checked IP addresses are deleted
For each laryCheckedIPAddrID in Request.Form("chkDelete")

	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))


	'Here we use the less effiecient ADO to delete from the database this way we can throw in a requery while we wait for slow old MS Access to catch up

	'Delete the IP address from the database	
	strSQL = "SELECT * " & _
		"FROM " & strDbTable & "BanList " & _
		"WHERE " & strDbTable & "BanList.Ban_ID = "  & laryCheckedIPAddrID & ";"
	
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




'If this is a post back then  update the database
If Request.Form("IP") <> "" AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	'Read in the IP address to block
	strBlockIP = Trim(Mid(Request.Form("IP"), 1, 30))
	strReason = Trim(Mid(Request.Form("Reason"), 1, 40))
	
	strSQL = "SELECT " & strDbTable & "BanList.Ban_ID, " & strDbTable & "BanList.IP, " & strDbTable & "BanList.Reason " & _
	"FROM " & strDbTable & "BanList " & _
	"WHERE " & strDbTable & "BanList.IP = '" & formatSQLInput(strBlockIP) & "';"
	
	'Set the cursor	type property of the record set	to Forward Only
	rsCommon.CursorType = 0
	
	'Set the Lock Type for the records so that the record set is only locked when it is updated
	rsCommon.LockType = 3
	
	'Query the database
	rsCommon.Open strSQL, adoCon


	'Update the recordset
	With rsCommon
	
		'If no record found for IP add new
		If rsCommon.EOF Then .AddNew

		'Update	the recorset
		.Fields("IP") = strBlockIP
		.Fields("Reason") = strReason

		'Update db
		.Update

		'Re-run the query as access needs time to catch up
		.ReQuery

	End With
	
	rsCommon.Close
End If




'Read in all the blocked IP address from the database

'Initalise the strSQL variable with an SQL statement to query the database to count the number of topics in the forums
strSQL = "SELECT " & strDbTable & "BanList.Ban_ID, " & strDbTable & "BanList.IP, " & strDbTable & "BanList.Reason " & _
	"FROM " & strDbTable & "BanList " & _
	"WHERE " & strDbTable & "BanList.IP Is Not Null " & _
	"ORDER BY " & strDbTable & "BanList.IP ASC;"

'Query the database
rsCommon.Open strSQL, adoCon





%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>IP Blocking</title>
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

	//Check for an IP address
	if (document.frmIPadd.IP.value==""){
		alert("Please enter an IP address or range");
		document.frmIPadd.IP.focus();
		return false;
	}

	return true;
}
</script>
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center"><h1>IP Address Blocking</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br />
  <span class="text">From here you can Block individual IP Addresses or Ranges.<br />
  <br />
  Anyone falling into a blocked IP address or range, will find many of the functions of the board disabled, including registering, and posting.<br />
  <br />
  Be careful when blocking IP addresses as you may block legitimate users. Many people share the same IP address, like AOL users, blocking one may block another 500,000 users from using your board.</span><br />
</div>
<br />
<form action="admin_ip_blocking.asp<% = strQsSID1 %>" method="post" name="frmIPList" id="frmIPList">
  <table align="center" cellpadding="4" cellspacing="1" class="tableBorder" style="width:450px">
    <tr align="left">
      <td colspan="3" class="tableLedger">Blocked IP Address List</td>
    </tr><%
'Display the IP block list
If rsCommon.EOF Then 
		
	'Disply no entires forun
	Response.Write("<td colspan=""2"" align=""center"" class=""tableRow"">You have no blocked IP address</td>")
	
'Else disply the IP block list
Else
	
	'Loop through the recordset
	Do While NOT rsCommon.EOF
	
     		'Read in the details
     		lngBlockedIPID = CLng(rsCommon("Ban_ID"))
		strBlockedIPList = rsCommon("IP")
		strReason = rsCommon("Reason")
     
     %>
    <tr class="tableRow">
      <td width="3%"><input type="checkbox" name="chkDelete" value="<% = lngBlockedIPID %>" /></td>
      <td nowrap="nowrap"><% = strBlockedIPList %> <a href="https://network-tools.webwiz.net/ip-information.htm?ip=<% = Server.URLEncode(strBlockedIPList) %>" target="_blank"><img src="<% = strImagePath %>new_window.png" alt="<% = strTxtIP & " " & strTxtInformation %>" title="<% = strTxtIP & " " & strTxtInformation %>" /></a></td>
      <td><% = strReason %></td>
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
      <td colspan="3" valign="top" class="tableRow">
      	<input type="hidden" name="formID" id="formID1" value="<% = getSessionItem("KEY") %>" />
      	<input type="submit" name="Submit" value="Remove IP Address or Range" />
      </td>
    </tr>
  </table>
  <br />
</form>
<form action="admin_ip_blocking.asp<% = strQsSID1 %>" method="post" name="frmIPadd" id="frmIPadd" onsubmit="return CheckForm();">
  <table align="center" cellpadding="4" cellspacing="1" class="tableBorder" style="width:450px">
    <tr align="left" >
      <td colspan="2" class="tableLedger"> Block IP Address or Range </td>
    </tr>
    <tr class="tableRow">
      <td colspan="2" align="center" class="smText"> The * wildcard character can be used to block IP ranges. <br />
      eg. To block the range '200.200.200.0 - 255' you would use '200.200.200.*' </td>
    </tr>
    <tr>
      <td align="right" class="tableRow"  width="35%"> IP Address/Range:</td>
      <td class="tableRow"><input type="text" name="IP" size="20" maxlength="40" value="<% = Server.HTMLEncode(strBlockIP) %>" <% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
     <tr class="tableRow">
      <td align="right"><% = strTxtReason %>:</td>
      <td><input type="text" name="Reason" id="Reason" size="20" maxlength="30" /></td>
    </tr>
    <tr align="center">
      <td colspan="2" valign="top" class="tableBottomRow">
       <input type="hidden" name="formID" id="formID2" value="<% = getSessionItem("KEY") %>" />
       <input type="submit" name="Submit2" value="Block IP Address or Range" />
      </td>
    </tr>
  </table>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
