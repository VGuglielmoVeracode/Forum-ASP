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
Dim intUserGroupID	'Holds the group ID
Dim strGroupName	'Holds the name of the group
Dim lngMinimumPoints	'Holds the minimum amount of posts to be in that group
Dim blnSpecialGroup	'Set to true if a special group
Dim intStars		'Holds the number of stars for the group
Dim strCustomStars	'Holds the custom stars image if there is one fo0r this group
Dim blnStartingGroup	'Set to true if it is the starting group
Dim strLadderGroup	'Holds the ladder group this group is part of


'If this is a postback update the strating group
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	'Read in the group ID to make the starting group
	intUserGroupID = IntC(Request.Form("start"))

	'Clear the old strating group
	strSQL = "UPDATE " & strDbTable & "Group "
	strSQL = strSQL & "SET " & strDbTable & "Group.Starting_group = " & strDBFalse & " "
	strSQL = strSQL & " WHERE " & strDbTable & "Group.Starting_group = " & strDBTrue & ";"


	'Write the updated number of posts to the database
	adoCon.Execute(strSQL)
	

	'Read the various groups from the database
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Group.* FROM " & strDbTable & "Group WHERE " & strDbTable & "Group.Group_ID = " & intUserGroupID & ";"
	
	'Set the Lock Type for the records so that the record set is only locked when it is updated
	rsCommon.LockType = 3
	
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'If a record is returned update it
	If not rsCommon.EOF Then	
		
		rsCommon.Fields("Starting_group") = True
		rsCommon.Update
		
		'Requery to let access catch up
		rsCommon.Requery
	End If
	
	
	'Close the recordset
	rsCommon.Close
End If
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Administer Member Groups</title>
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
<div align="center"><h1>Administer Member Groups</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a></div>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" class="text"><br />
      From here you can create, delete and edit Member Group. Member Groups can be used to grant permissions to forums for Groups of Forums Members.<br />
      <br />
      Click  the Member Group Name to change the details and settings for that Member Group.<br />
      <br />
      To change the 'New Members Group', select which group you want as the starting group and click the 'Update New Members Group' button. </td>
  </tr>
</table>
<form action="admin_view_groups.asp<% = strQsSID1 %>" method="post" name="form1" id="form1">
  <br />
  <table border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
    <tr valign="top">
      <td width="30%" nowrap="nowrap" class="tableLedger">Member Group</td>
      <td width="10%" nowrap="nowrap" class="tableLedger">Stars</td>
      <td width="30%" height="12" nowrap="nowrap" class="tableLedger">Ladder Group</td>
      <td width="10%" height="12" align="center" nowrap="nowrap" class="tableLedger">Group Members</td>
      <td width="5%" height="12" align="center" nowrap="nowrap" class="tableLedger">New Members Group</td>
      <td width="5%" height="12" align="center" nowrap="nowrap" class="tableLedger">Delete</td>
    </tr>
    <%

'Read the various groups from the database
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Group.*, " & strDbTable & "LadderGroup.* " & _
	"FROM " & strDbTable & "Group " & _
	"LEFT JOIN " & strDbTable & "LadderGroup ON " & strDbTable & "Group.Ladder_ID = " & strDbTable & "LadderGroup.Ladder_ID " & _
	"ORDER BY " & strDbTable & "LadderGroup.Ladder_Name ASC, " & strDbTable & "Group.Minimum_posts ASC, " & strDbTable & "Group.Group_ID ASC;"

'Query the database
rsCommon.Open strSQL, adoCon

'Check there are user groups to display
If rsCommon.EOF Then

	'If there are no user groups display then display the appropriate error message
	Response.Write vbCrLf & "<td bgcolor=""#FFFFFF"" colspan=""4""><span class=""text"">There are currently no User Groups to display.</span></td>"

'Else there the are user groups so write the HTML to display them
Else


	'Loop round to read in all the groups in the database
	Do While NOT rsCommon.EOF

		'Get the category name from the database
		intUserGroupID = CInt(rsCommon("Group_ID"))
		strGroupName = rsCommon("Name")
		lngMinimumPoints = CLng(rsCommon("Minimum_posts"))
		blnSpecialGroup = CBool(rsCommon("Special_rank"))
		intStars = CInt(rsCommon("Stars"))
		strCustomStars = rsCommon("Custom_stars")
		blnStartingGroup = CBool(rsCommon("Starting_group"))
		strLadderGroup = rsCommon("Ladder_Name")

		'Display the groups

%>
    <tr>
      <td class="tableRow"><a href="admin_group_details.asp?GID=<% = intUserGroupID %><% = strQsSID2 %>"><% = strGroupName %></a>
      <% If intUserGroupID = 2 Then Response.Write(" - <span class=""smText"">(Un-registered users)</span>") %></td>
      <td class="tableRow"><img src="<% If strCustomStars <> "" Then Response.Write(strCustomStars) Else Response.Write(strImagePath & "/" & intStars & "_star_rating.png") %>" alt="<% = intStars %> stars" /></td>
      <td class="tableRow"><%

     	'If this is a special group then disply a message that it is not a rank group
     	If blnSpecialGroup Then
     		Response.Write("Non Ladder Group")

     	'If this is a rank group disply the minimum number of posts to be in this group
     	Else
     		Response.Write("Ladder Group '" & strLadderGroup & "' Min. Points: " & lngMinimumPoints)

	End If

     %>
      </td>
      
       <td  align="center" class="tableRow" class="smText">
        <a href="admin_select_members.asp?GID=<% = intUserGroupID & "&XID=" & getSessionItem("KEY") & strQsSID2 %>"  class="smText">View Members</a>
      </td>
      
      <td align="center" class="tableRow"><%
     
     	'If this is the admin group or guest group then don't let em change the starting group
	If intUserGroupID <> 1 AND intUserGroupID <> 2 Then 
     
     %>
        <input type="radio" name="start" value="<% = intUserGroupID %>"<% If blnStartingGroup Then Response.Write(" checked") %> />
        <%
     
	End If
       %>
      </td>
      
     
      
      <td align="center" class="tableRow"><%

     		'If this is not the admin group, guest group, or starting group let them delete the group
     		If intUserGroupID <> 1 AND intUserGroupID <> 2 AND blnStartingGroup = False Then

     %>
        <a href="admin_delete_group.asp?GID=<% = intUserGroupID & "&XID=" & getSessionItem("KEY") & strQsSID2 %>" onclick="return confirm('Are you sure you want to Delete this User Group?\n\nWARNING: Deleting this user group will mean all members of this user group will be transfered to the Starting Group!')"><img src="<% = strImagePath %>delete.png" border="0" alt="Delete" /></a>
        <%

	End If

     %>
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
  <table width="100%" border="0" cellspacing="0" cellpadding="4">
    <tr>
      <td align="right"><input name="postBack" type="hidden" id="postBack" value="true" />
      	<input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
        <input type="submit" name="Submit" value="Update New Members Group"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      </td>
    </tr>
  </table>
</form>
<table width="98%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr align="center">
    <td width="50%"><form action="admin_group_details.asp" method="post" name="form2" id="form2">
        <input type="submit" name="Submit" value="Create New Member Group" />
      </form></td>
  </tr>
</table>
<br />
<table border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
  <tr>
    <td align="center" class="tableLedger">Please Read the Following</td>
  </tr>
  <tr>
    <td class="tableSubLedger">Admin Group</td>
  </tr>
  <tr>
    <td class="tableRow">The Admin Group, (the first group in the list), cannot be deleted. Any member placed in this user group has admin powers over the whole board and can also use this admin area, so be careful of the users you place in this group! For security it is recomended you only ever have 1 member in this Group.</td>
  </tr>
  <tr>
    <td class="tableSubLedger">Guest Group</td>
  </tr>
  <tr>
    <td class="tableRow">The Guest Group, (the second group in the list), cannot be deleted. Visitors to your forums that have not registered come under this 'Guest Group'.</td>
  </tr>
  <tr>
    <td class="tableSubLedger">New Members Group</td>
  </tr>
  <tr>
    <td class="tableRow">The New Members Group is a special group that new members automatically become part of when they first register. You can only have one 'New Members Group', if you want to delete the New Members Group you must first select another group to be the New Members Group. You change the New Members Group by editing the details of a group and selecting it to become the New Members Group.<br />
    </td>
  </tr>
</table>
<!-- #include file="includes/admin_footer_inc.asp" -->
