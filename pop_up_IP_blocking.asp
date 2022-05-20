<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="language_files/admin_language_file_inc.asp" -->
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
Dim lngTopicID 		'Holds the topic ID number to return to
Dim strBlockIP		'Holds the IP address to block
Dim strBlockedIPList	'Holds the IP addresses in the blocked list
Dim lngBlockedIPID	'Holds the ID number of the blcoked db record
Dim laryCheckedIPAddrID	'Holds the array of IP addresses to be ditched
Dim strReason		'Holds the reason for the IP ban


blnSslEnabledPage = True


'Read in forum ID
intForumID = IntC(getSessionItem("FID"))


'Call the moderator function and see if the user is a moderator (if not logged in as an admin)
If blnAdmin = False Then blnModerator = isModerator(intForumID, intGroupID)
	


'If the person is not an admin or a moderator then send them away
If blnAdmin = false AND blnModerator = false Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)



'Only run the following lines if this is a moderator or an admin
Else

	'Run through till all checked IP addresses are deleted
	For each laryCheckedIPAddrID in Request.Form("chkDelete")
	
	
		'Check the form ID to prevent XCSRF
		Call checkFormID(Request.Form("formID"))
	
	
		'Update log file
		If blnLoggingEnabled AND blnModeratorLogging Then 
			
			'Read in the IP address or range being deleted
			strSQL = "SELECT " & strDbTable & "BanList.IP " & _
				"FROM " & strDbTable & "BanList " & _
				"WHERE " & strDbTable & "BanList.Ban_ID = " & CLng(laryCheckedIPAddrID) & ";"
			
			'Open DB
			rsCommon.Open strSQL, adoCon
			
			'Write to log file
			If NOT rsCommon.EOF Then Call logAction(strLoggedInUsername, "Un-Blocked IP Address/Range '" & rsCommon("IP") & "'")
			
			'Close DB
			rsCommon.Close
		End If
	
		'Delete from DB
		strSQL = "DELETE FROM " & strDbTable & "BanList " & strRowLock & " " & _
		"WHERE " & strDbTable & "BanList.Ban_ID = " & CInt(laryCheckedIPAddrID) & ";"
	
		'Delete the threads
		adoCon.Execute(strSQL)	
		
		
		'Update log file
		If blnLoggingEnabled AND blnModeratorLogging Then Call logAction(strLoggedInUsername, "Blocked IP Address/Range '" & strBlockIP & "'")
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
End If
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
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

<script language="JavaScript">
function CheckForm () {

	var errorMsg = "";
	var formArea = document.getElementById('frmIPadd');

	//Check for a subject
	if (formArea.IP.value==""){
		errorMsg += "\n<% = strTxtErrorIPEmpty %>";
	}

	//If there is aproblem with the form then display an error
	if (errorMsg != ""){
		msg = "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine1 %>\n";
		msg += "<% = strTxtErrorDisplayLine2 %>\n";
		msg += "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine3 %>\n";

		errorMsg += alert(msg + errorMsg + "\n\n");
		return false;
	}

	return true;
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body style="margin:0px;" OnLoad="self.focus();">
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="center"><h1><% = strTxtIPBlocking %></h1></td>
 </tr>
</table>
<br />
<form name="frmIPList" id="frmIPList" method="post" action="pop_up_IP_blocking.asp<% = strQsSID1 %>">
<table cellspacing="1" cellpadding="3" class="tableBorder" align="center" style="width: 450px;">
 <tr class="tableLedger">
  <td colspan="3"><% = strTxtBlockedIPList %></td>
 </tr>
 <tr class="tableRow"><%
	'Display the IP blcok list
	If rsCommon.EOF Then 
		
		'Disply no entires forun
		Response.Write(vbCrLf & "  <td colspan=""2"" align=""center""><br /><strong>" & strTxtYouHaveNoBlockedIpAddesses & "</strong><br /><br /></td>")
	
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
  <td width="3%"><input type="checkbox" name="chkDelete" value="<% = lngBlockedIPID %>"></td>
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
          </select></td>
 </tr>
 <tr class="tableBottomRow">
  <td colspan="3" align="center">
   <input type="hidden" name="formID" id="formID1" value="<% = getSessionItem("KEY") %>" />
   <input type="submit" name="Submit" id="Submit" value="<% = strTxtRemoveIP %>" />
  </td>
 </tr>
</table>
</form>
<br />
<form name="frmIPadd" id="frmIPadd" method="post" action="pop_up_IP_blocking.asp<% = strQsSID1 %>" onSubmit="return CheckForm();">
<table cellspacing="1" cellpadding="3" class="tableBorder" align="center" style="width: 450px;">
 <tr class="tableLedger">
  <td colspan="2"><% = strTxtBlockIPAddressOrRange %></td>
 </tr>
 <tr class="tableRow">
  <td colspan="2" align="center" class="smText"><% = strTxtBlockIPRangeWhildcardDescription %></td>
 </tr>
 <tr class="tableRow">
  <td align="right" width="35%"><% = strTxtIpAddressRange %>:</td>
  <td><input type="text" name="IP" id="IP" size="20" maxlength="30" value="<% = Server.HTMLEncode(Request.QueryString("IP")) %>" /> <% If Request.QueryString("IP") <> "" Then %><a href="https://network-tools.webwiz.net/ip-information.htm?ip=<% = Server.HTMLEncode(Request.QueryString("IP")) %>" target="_blank"><img src="<% = strImagePath %>new_window.png" alt="<% = strTxtIP & " " & strTxtInformation %>" title="<% = strTxtIP & " " & strTxtInformation %>" /></a><% End If %></td>
 </tr>
 <tr class="tableRow">
  <td align="right"><% = strTxtReason %>:</td>
  <td><input type="text" name="Reason" id="Reason" size="20" maxlength="40" /></td>
 </tr>
 <tr align="center" class="tableBottomRow">
  <td colspan="2">
   <input type="hidden" name="formID" id="formID2" value="<% = getSessionItem("KEY") %>" />
   <input type="submit" name="Submit" id="Submit" value="<% = strTxtBlockIPAddressRange %>"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
  </td>
 </tr>
</table>
</form> 
<br />
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="center"><input type="button" name="ok" onclick="javascript:window.close();" value="<% = strTxtCloseWindow %>"><br />
  </td>
 </tr>
</table>
</body>