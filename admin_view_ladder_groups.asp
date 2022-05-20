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

'Dimension variables
Dim intLadderGroupID	'Holds the ladder group ID
Dim strLadderGroupName	'Holds the name of the ladder group


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Administer Ladder Groups</title>
<meta name="generator" content="Web Wiz Forums" />
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
<div align="center"><h1>Administer Ladder Groups</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" class="text"><br />
      From here you can create, delete and edit Forum Ladder Groups. Ladder Groups are can be used to allow Forum Members to move up Groups by the number of points they receive in forums.<br />
      <br />
      </td>
  </tr>
</table>
<form action="admin_view_groups.asp<% = strQsSID1 %>" method="post" name="form1" id="form1">
  <br />
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder" style="width: 500px;">
    <tr valign="top">
      <td width="80%" nowrap="nowrap" class="tableLedger">Ladder Group</td>
      <td width="20%" height="12" align="center" nowrap="nowrap" class="tableLedger">Delete</td>
    </tr>
    <%

'Read the various Ladder groups from the database
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "LadderGroup.* " & _
	"FROM " & strDbTable & "LadderGroup " & _
	"ORDER BY " & strDbTable & "LadderGroup.Ladder_Name ASC;"

'Query the database
rsCommon.Open strSQL, adoCon

'Check there are Ladder groups to display
If rsCommon.EOF Then

	'If there are no user groups display then display the appropriate error message
	Response.Write vbCrLf & "<td bgcolor=""#FFFFFF"" colspan=""4""><span class=""text"">There are currently no Ladder Groups to display.</span></td>"

'Else there the are user groups so write the HTML to display them
Else


	'Loop round to read in all the Ladder groups in the database
	Do While NOT rsCommon.EOF

		'Get the category name from the database
		intLadderGroupID = CInt(rsCommon("Ladder_ID"))
		strLadderGroupName = rsCommon("Ladder_Name")

		'Display the Ladder groups

%>
    <tr>
      <td class="tableRow"><a href="admin_ladder_group_details.asp?LID=<% = intLadderGroupID & strQsSID2 %>"><% = strLadderGroupName %></a></td>
      <td width="4%"  align="center" class="tableRow">
        <a href="admin_delete_ladder_group.asp?LID=<% = intLadderGroupID & "&XID=" & getSessionItem("KEY") & strQsSID2 %>" onclick="return confirm('Are you sure you want to Delete this Ladder Group?\n\nWARNING: Deleting this user group will mean all Groups that are part of this Ladder Group will need to be edited to place them in a new Ladder Group!')"><img src="<% = strImagePath %>delete.png" border="0" alt="Delete" /></a>
      </td>
    </tr>
    <%

		'Move to the next database record
		rsCommon.MoveNext
	Loop
End If

'Reset Server Objects
rsCommon.Close
Call closeDatabase()
%>
  </table>
</form>
<br />
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr align="center">
    <td width="50%"><form action="admin_ladder_group_details.asp" method="post" name="form2" id="form2">
        <input type="submit" name="Submit" value="Create New Ladder Group" />
      </form></td>
  </tr>
</table>
<!-- #include file="includes/admin_footer_inc.asp" -->
