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

Dim strGroupName
Dim intSelGroupID


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Group Permissions</title>
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
  <h1>Member Group Permissions</h1><br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <br />
    From here you can edit permissions for various Member Groups.<br />
    <br />
    Please note that Member Group Permissions can be overridden by setting permissions for individual members on forums.<br />
    <br />
    Select the Member Group that you would like to Edit Permissions for.<br />
  </p>
</div>
<form action="admin_group_permissions.asp" method="get" name="frmSelectForum" id="frmSelectForum">
  <table align="center" cellpadding="4" cellspacing="1" class="tableBorder" style="width:480px">
    <tr>
      <td width="51%" height="2" align="left" valign="top" class="tableLedger">Select the Member Group you would like to  Edit permissions for</td>
    </tr>
    <tr>
      <td  height="12" align="left" class="tableRow"><select name="GID"><%

'Read in the group name from the database
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Group.Group_ID, " & strDbTable & "Group.Name FROM " & strDbTable & "Group ORDER BY " & strDbTable & "Group.Group_ID ASC;"

'Query the database
rsCommon.Open strSQL, adoCon


'Loop through all the categories in the database
Do while NOT rsCommon.EOF

	'Read in the deatils for the category
	strGroupName = rsCommon("Name")
	intSelGroupID = CInt(rsCommon("Group_ID"))

	'Display a link in the link list to the cat
	Response.Write (vbCrLf & "		<option value=""" & intSelGroupID & """")
	Response.Write(">" & strGroupName & "</option>")


	'Move to the next record in the recordset
	rsCommon.MoveNext
Loop

'Reset server objects
rsCommon.Close
Call closeDatabase()

%>
      </select></td>
    </tr>
    <tr>
      <td width="51%"  height="12" align="center" class="tableBottomRow">
        <input type="hidden" name="SID" id="SID" value="<% = strQsSID %>" />
        <input type="submit" name="Submit" value="Edit Member Group Permissions" /></td>
    </tr>
  </table>
  <div align="center"><br />
  </div>
</form>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
