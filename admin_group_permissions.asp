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
Dim rsCommon2		'Holds a secound recordset for the page
Dim intUserGroupID	'Holds the group ID
Dim strGroupName	'Holds the name of the group
Dim intCatID
Dim sarryForums
Dim intCurrentRecord
Dim sarrySubForums
Dim intCurrentRecord2
Dim intSubForumID



'Initlise variables
intCatID = 0


'Read in the details
intUserGroupID = IntC(Request.QueryString("GID"))





'Intialise the ADO recordset object
Set rsCommon2 = Server.CreateObject("ADODB.Recordset")


'If this is a post back update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	'Read in the groups from db
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Forum.Forum_ID FROM " & strDbTable & "Forum ORDER BY " & strDbTable & "Forum.Forum_Order ASC;"
	
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'Loop through all the categories in the database
	Do while NOT rsCommon.EOF
	
	
		'Get the group ID
		intForumID = CInt(rsCommon("Forum_ID"))
	
		'Read in the permssions from the db for this group (not very efficient doing it this way, but this page won't be run often)
		'Initalise the strSQL variable with an SQL statement to query the database
		strSQL = "SELECT " & strDbTable & "Permissions.* FROM " & strDbTable & "Permissions WHERE " & strDbTable & "Permissions.Group_ID = " & intUserGroupID & " AND " & strDbTable & "Permissions.Forum_ID = " & intForumID & ";"
		
		'Set the cursor type property of the record set to Forward Only
		rsCommon2.CursorType = 0
		
		'Set the Lock Type for the records so that the record set is only locked when it is updated
		rsCommon2.LockType = 3
		
		'Query the database
		rsCommon2.Open strSQL, adoCon
		
		With rsCommon2
			'If no records are returned then add a new record to the database
			If .EOF Then .AddNew
			
			'Update the recordset
			.Fields("Group_ID") = intUserGroupID
			.Fields("Forum_ID") = intForumID
			.Fields("View_Forum") = BoolC(Request.Form("read" & intForumID))
			.Fields("Post") = BoolC(Request.Form("topic" & intForumID))
			.Fields("Priority_posts") = BoolC(Request.Form("sticky" & intForumID))
			.Fields("Reply_posts") = BoolC(Request.Form("reply" & intForumID))
			.Fields("Edit_posts") = BoolC(Request.Form("edit" & intForumID))
			.Fields("Delete_posts") = BoolC(Request.Form("delete" & intForumID))
			.Fields("Poll_create") = BoolC(Request.Form("polls" & intForumID))
			.Fields("Vote") = BoolC(Request.Form("vote" & intForumID))
			.Fields("Display_post") = BoolC(Request.Form("approve" & intForumID))
			.Fields("Moderate") = BoolC(Request.Form("moderator" & intForumID))
			.Fields("Calendar_event") = BoolC(Request.Form("calEvent" & intForumID))
			.Fields("Attachments") = False
			.Fields("Image_upload") = False
	
			'Update the database
			.Update
		End With
		
		
		
		'Close rsCommon2
		rsCommon2.Close
	            
		'Move to the next record in the recordset
		rsCommon.MoveNext
	Loop
	
	rsCommon.Close
	
End If 





	
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Group.* FROM " & strDbTable & "Group WHERE " & strDbTable & "Group.Group_ID = " & intUserGroupID & ";"
	
'Query the database
rsCommon.Open strSQL, adoCon

If NOT rsCommon.EOF Then

	'Get the category name from the database
	strGroupName = rsCommon("Name")
End If
	
'Close the rs
rsCommon.Close


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Group Permissions</title>
<meta name="generator" content="Web Wiz Forums" />
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
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center"><h1> <% = strGroupName %> Group Permissions</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <a href="admin_group_permissions_form.asp<% = strQsSID1 %>">Select an alternate Member Group </a></div>
<form action="admin_group_permissions.asp?GID=<% = intUserGroupID %><% = strQsSID2 %>" method="post" name="frmGroup" id="frmGroup">
  <table width="100%" height="58" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td align="center" class="text"><span class="lgText"><br />
        <% = strGroupName %>
        Group Permissions</span><br />
        Use the grid below to set Permissions for the this Member Group on various forums.<br />
        <a href="#per" class="smLink">What do the different permissions mean?</a><br />
        <br />
        <table border="0" cellpadding="2" cellspacing="1" class="tableBorder">
          <tr>
            <td width="194" align="left" class="tableLedger">Member Group</td>
            <td width="43" align="center" class="tableLedger">Access</td>
            <td width="43" align="center" class="tableLedger"> New Topics</td>
            <td width="43" align="center" class="tableLedger">Sticky Topics</td>
            <td width="43" align="center" class="tableLedger">Post Reply</td>
            <td width="43" align="center" class="tableLedger">Edit Posts</td>
            <td width="43" align="center" class="tableLedger">Delete Posts</td>
            <td width="43" align="center" class="tableLedger">New Polls </td>
            <td width="43" align="center" class="tableLedger">Poll Vote </td>
            <td width="43" align="center" class="tableLedger">Calendar Event</td>
            <td width="43" align="center" class="tableLedger">Post Approval </td>
            <td width="43" align="center" class="tableLedger">Forum Moderator</td>
          </tr>
          <tr class="tableSubLedger">
           <td align="left">Check All</td>
           <td align="center"><input type="checkbox" name="chkAllread" id="chkAllread" onclick="checkAll('read');" /></td>
           <td align="center"><input type="checkbox" name="chkAlltopic" id="chkAlltopic" onclick="checkAll('topic');" /></td>
           <td align="center"><input type="checkbox" name="chkAllsticky" id="chkAllsticky" onclick="checkAll('sticky');" /></td>
           <td align="center"><input type="checkbox" name="chkAllreply" id="chkAllreply" onclick="checkAll('reply');" /></td>
           <td align="center"><input type="checkbox" name="chkAlledit" id="chkAlledit" onclick="checkAll('edit');" /></td>
           <td align="center"><input type="checkbox" name="chkAlldelete" id="chkAlldelete" onclick="checkAll('delete');" /></td>
           <td align="center"><input type="checkbox" name="chkAllpolls" id="chkAllpolls" onclick="checkAll('polls');" /></td>
           <td align="center"><input type="checkbox" name="chkAllvote" id="chkAllvote" onclick="checkAll('vote');" /></td>
           <td align="center"><input type="checkbox" name="chkAllcalEvent" id="chkAllcalEvent" onclick="checkAll('calEvent');" /></td>
           <td align="center"><input type="checkbox" name="chkAllapprove" id="chkAllapprove" onclick="checkAll('approve');" /></td>
           <td align="center">&nbsp;</td>
          </tr>
          <%

'Read in the groups from db
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Category.Cat_ID, " & strDbTable & "Category.Cat_name, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Forum_URL " & _
"FROM " & strDbTable & "Category, " & strDbTable & "Forum " & _
"WHERE " & strDbTable & "Category.Cat_ID=" & strDbTable & "Forum.Cat_ID AND " & strDbTable & "Forum.Sub_ID=0 " & _
"ORDER BY " & strDbTable & "Category.Cat_order ASC, " & strDbTable & "Category.Cat_ID ASC, " & strDbTable & "Forum.Forum_Order ASC;"

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the row from the db using getrows for better performance
If NOT rsCommon.EOF Then 
	sarryForums = rsCommon.GetRows()
End If

'close
rsCommon.Close

'If no record returned display so
If NOT isArray(sarryForums) Then
	%>
	  <tr>
            <td align="left" class="tableRow" colspan="14">There are presently no Forums created to set permissions on</td>
          </tr><%

'If records returned
Else
    
	'Loop round to read in all the forums in the database
	Do While intCurrentRecord =< Ubound(sarryForums,2)
  
		'Get the forum ID
		intForumID = CInt(sarryForums(2,intCurrentRecord))
		
		'If this is a different cat display the cat ID
		If intCatID <> CInt(sarryForums(0,intCurrentRecord)) Then 
				
			'Change the cat ID
			intCatID = CInt(sarryForums(0,intCurrentRecord))
			
			%>
          <tr>
            <td align="left" class="tableSubLedger" colspan="14"><% = sarryForums(1,intCurrentRecord) %></td>
          </tr>
          <%
       
	End If
	

		'Read in the permssions from the db for this group (not very efficient doing it this way, but this page won't be run often)
		'Initalise the strSQL variable with an SQL statement to query the database
		strSQL = "SELECT " & strDbTable & "Permissions.* " & _
		"FROM " & strDbTable & "Permissions " & _
		"WHERE " & strDbTable & "Permissions.Group_ID = " & intUserGroupID & " AND " & strDbTable & "Permissions.Forum_ID = " & intForumID & ";"
			
		'Query the database
		rsCommon.Open strSQL, adoCon
			
		'If no records are returned use default values
		If rsCommon.EOF Then		

%>
          <tr>
            <td align="left" class="tableRow"><% = sarryForums(3,intCurrentRecord) %></td>
            <td align="center" class="tableRow"><input name="read<% = intForumID %>" type="checkbox" id="read" value="true" checked /></td>
            <td align="center" class="tableRow"><input name="topic<% = intForumID %>" type="checkbox" id="topic" value="true"<% If intUserGroupID <> 2 Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="sticky<% = intForumID %>" type="checkbox" id="sticky" value="true"<% If intUserGroupID = 1 Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="reply<% = intForumID %>" type="checkbox" id="reply" value="true"<% If intUserGroupID <> 2 Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="edit<% = intForumID %>" type="checkbox" id="edit" value="true"<% If intUserGroupID <> 2 Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="delete<% = intForumID %>" type="checkbox" id="delete" value="true"<% If intUserGroupID <> 2 Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="polls<% = intForumID %>" type="checkbox" id="polls" value="true" /></td>
            <td align="center" class="tableRow"><input name="vote<% = intForumID %>" type="checkbox" id="vote" value="true" /></td>
            <td align="center" class="tableRow"><input name="calEvent<% = intForumID %>" type="checkbox" id="calEvent" value="true" /></td>
            <td align="center" class="tableRow"><input name="approve<% = intForumID %>" type="checkbox" id="approve" value="true" /></td>
            <td align="center" class="tableRow"><input name="moderator<% = intForumID %>" type="checkbox" id="moderator" value="true" /></td>
          </tr>
          <%
      
		'Else display the values for this group
		 Else
%>
          <tr>
            <td align="left" class="tableRow"><% = sarryForums(3,intCurrentRecord) %></td>
            <td align="center" class="tableRow"><input name="read<% = intForumID %>" type="checkbox" id="read" value="true"<% If CBool(rsCommon("View_Forum")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="topic<% = intForumID %>" type="checkbox" id="topic" value="true"<% If CBool(rsCommon("Post")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="sticky<% = intForumID %>" type="checkbox" id="sticky" value="true"<% If CBool(rsCommon("Priority_posts")) Then Response.Write(" checked") %><% If intUserGroupID = 2 Then Response.Write(" disabled=true") %> /></td>
            <td align="center" class="tableRow"><input name="reply<% = intForumID %>" type="checkbox" id="reply" value="true"<% If CBool(rsCommon("Reply_posts")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="edit<% = intForumID %>"<% If intUserGroupID = 2 Then Response.Write(" disabled=true") %> type="checkbox" id="edit" value="true"<% If CBool(rsCommon("Edit_posts")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="delete<% = intForumID %>"<% If intUserGroupID = 2 Then Response.Write(" disabled=true") %> type="checkbox" id="delete" value="true"<% If CBool(rsCommon("Delete_posts")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="polls<% = intForumID %>" type="checkbox" id="polls" value="true"<% If CBool(rsCommon("Poll_create")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="vote<% = intForumID %>" type="checkbox" id="vote" value="true"<% If CBool(rsCommon("Vote")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="calEvent<% = intForumID %>" type="checkbox" id="calEvent" value="true"<% If CBool(rsCommon("Calendar_event")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="approve<% = intForumID %>" type="checkbox" id="approve" value="true"<% If CBool(rsCommon("Display_post")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="moderator<% = intForumID %>" type="checkbox" id="moderator" value="true"<% If CBool(rsCommon("Moderate")) Then Response.Write(" checked") %><% If intUserGroupID = 2 Then Response.Write(" disabled=true") %> /></td>
          </tr>
          <%
		End If
		
		'Close rsCommon
		rsCommon.Close
	        
	        
	        
	        '********* check for sub forums *****************
	        
	        'Reset intCurrentRecord2
		intCurrentRecord2 = 0
	        
	        'Read in the groups from db
		'Initalise the strSQL variable with an SQL statement to query the database
		strSQL = "SELECT " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Forum_URL " & _
		"FROM " & strDbTable & "Forum " & _
		"WHERE " & strDbTable & "Forum.Sub_ID=" & intForumID & " " & _
		"ORDER BY " & strDbTable & "Forum.Forum_Order ASC;"
		
		'Query the database
		rsCommon.Open strSQL, adoCon
		
		'Place rs in array
		If NOT rsCommon.EOF Then 
			sarrySubForums = rsCommon.GetRows()
		Else
			sarrySubForums = null
		End If
			
		'close
		rsCommon.Close
		
		'Read in the row from the db using getrows for better performance
		If isArray(sarrySubForums) Then 
			
			'Loop round to read in all the forums in the database
			Do While intCurrentRecord2 =< Ubound(sarrySubForums,2)
	
				'Get the forum ID
				intSubForumID = CInt(sarrySubForums(0,intCurrentRecord2))
			
			
				'Read in the permssions from the db for this group (not very efficient doing it this way, but this page won't be run often)
				'Initalise the strSQL variable with an SQL statement to query the database
				strSQL = "SELECT " & strDbTable & "Permissions.* " & _
				"FROM " & strDbTable & "Permissions " & _
				"WHERE " & strDbTable & "Permissions.Group_ID = " & intUserGroupID & " AND " & strDbTable & "Permissions.Forum_ID = " & intSubForumID & ";"
					
				'Query the database
				rsCommon.Open strSQL, adoCon
					
				'If no records are returned use default values
				If rsCommon.EOF Then		

%>
          <tr>
            <td align="left" class="tableRow">&nbsp;<img src="<% = strImagePath %>arrow.gif" />&nbsp;<% = sarrySubForums(1,intCurrentRecord2) %></td>
            <td align="center" class="tableRow"><input name="read<% = intSubForumID %>" type="checkbox" id="read" value="true" checked /></td>
            <td align="center" class="tableRow"><input name="topic<% = intSubForumID %>" type="checkbox" id="topic" value="true"<% If intUserGroupID <> 2 Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="sticky<% = intSubForumID %>" type="checkbox" id="sticky" value="true"<% If intUserGroupID = 1 Then Response.Write(" checked") %><% If intUserGroupID = 2 Then Response.Write(" disabled=true") %> /></td>
            <td align="center" class="tableRow"><input name="reply<% = intSubForumID %>" type="checkbox" id="reply" value="true"<% If intUserGroupID <> 2 Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="edit<% = intSubForumID %>" type="checkbox" id="edit" value="true"<% If intUserGroupID <> 2 Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="delete<% = intSubForumID %>" type="checkbox" id="delete" value="true"<% If intUserGroupID <> 2 Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="polls<% = intSubForumID %>" type="checkbox" id="polls" value="true" /></td>
            <td align="center" class="tableRow"><input name="vote<% = intSubForumID %>" type="checkbox" id="vote" value="true" /></td>
            <td align="center" class="tableRow"><input name="calEvent<% = intSubForumID %>" type="checkbox" id="calEvent" value="true" /></td>
            <td align="center" class="tableRow"><input name="approve<% = intSubForumID %>" type="checkbox" id="approve" value="true" /></td>
            <td align="center" class="tableRow"><input name="moderator<% = intSubForumID %>" type="checkbox" id="moderator" value="true"<% If intUserGroupID = 2 Then Response.Write(" disabled=true") %> /></td>
          </tr>
          <%
      
				'Else display the values for this group
		 		Else
%>
          <tr>
            <td align="left"class="tableRow">&nbsp;<img src="<% = strImagePath %>arrow.gif" />&nbsp;<% = sarrySubForums(1,intCurrentRecord2) %></td>
            <td align="center" class="tableRow"><input name="read<% = intSubForumID %>" type="checkbox" id="read" value="true"<% If CBool(rsCommon("View_Forum")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="topic<% = intSubForumID %>" type="checkbox" id="topic" value="true"<% If CBool(rsCommon("Post")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="sticky<% = intSubForumID %>" type="checkbox" id="sticky" value="true"<% If CBool(rsCommon("Priority_posts")) Then Response.Write(" checked") %><% If intUserGroupID = 2 Then Response.Write(" disabled=true") %> /></td>
            <td align="center" class="tableRow"><input name="reply<% = intSubForumID %>" type="checkbox" id="reply" value="true"<% If CBool(rsCommon("Reply_posts")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="edit<% = intSubForumID %>"<% If intUserGroupID = 2 Then Response.Write(" disabled=true") %> type="checkbox" id="edit" value="true"<% If CBool(rsCommon("Edit_posts")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="delete<% = intSubForumID %>"<% If intUserGroupID = 2 Then Response.Write(" disabled=true") %> type="checkbox" id="delete" value="true"<% If CBool(rsCommon("Delete_posts")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="polls<% = intSubForumID %>" type="checkbox" id="polls" value="true"<% If CBool(rsCommon("Poll_create")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="vote<% = intSubForumID %>" type="checkbox" id="vote" value="true"<% If CBool(rsCommon("Vote")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="calEvent<% = intSubForumID %>" type="checkbox" id="calEvent" value="true"<% If CBool(rsCommon("Calendar_event")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="approve<% = intSubForumID %>" type="checkbox" id="approve" value="true"<% If CBool(rsCommon("Display_post")) Then Response.Write(" checked") %> /></td>
            <td align="center" class="tableRow"><input name="moderator<% = intSubForumID %>" type="checkbox" id="moderator" value="true"<% If CBool(rsCommon("Moderate")) Then Response.Write(" checked") %><% If intUserGroupID = 2 Then Response.Write(" disabled=true") %> /></td>
          </tr>
          <%
				End If
		
				'Close rsCommon
				rsCommon.Close
			
				'Move to the next record in the recordset
				intCurrentRecord2 = intCurrentRecord2 + 1
			Loop
		End If
	        
		'Move to the next record in the recordset
		intCurrentRecord = intCurrentRecord + 1
	Loop
End If

'Reset Server Objects
Set rsCommon2 = Nothing
Call closeDatabase()

%>
        </table></td>
    </tr>
  </table>
  <div align="center"><br />
    <input type="hidden" name="postBack" value="true" />
    <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
    <input type="submit" name="Submit" value="Update Member Group Permissions"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
    <input type="reset" name="Reset" value="Reset Form" />
    <br />
  </div>
</form>
<br />
  <a name="per" id="per"></a>
  <br />
  <br />
<table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
     <td colspan="2" class="tableLedger">Forum Permissions Table </td>
    </tr>
    <tr>
     <td width="24%" align="right" valign="top"  class="tableRow"><strong>Access:</strong></td>
     <td width="76%" valign="top"  class="tableRow">Allows the Group access to the forum </td>
    </tr>
    <tr>
     <td width="24%" align="right" valign="top"  class="tableRow"><strong>New Topics:</strong></td>
     <td width="76%" valign="top"  class="tableRow">Allows the Group to post new topics </td>
    </tr>
    <tr>
     <td align="right" valign="top"  class="tableRow"><strong>Sticky Topics:</strong></td>
     <td valign="top"  class="tableRow">Allows the Group to post sticky topics that remain at the top of the forum </td>
    </tr>
    <tr>
     <td width="24%" align="right" valign="top"  class="tableRow"><strong>Post Reply:<br />
      </strong></td>
     <td width="76%" valign="top"  class="tableRow">Allows the Group to reply to posts </td>
    </tr>
    <tr>
     <td align="right" valign="top"  class="tableRow"><strong>Edit Posts:</strong></td>
     <td valign="top"  class="tableRow">Allows the Group to edit their posts </td>
    </tr>
    <tr>
     <td align="right" valign="top"  class="tableRow"><strong>Delete Posts:</strong></td>
     <td valign="top"  class="tableRow">Allows the Group to delete their posts, but only if no-one has posted a reply </td>
    </tr>
    <tr>
     <td align="right" valign="top"  class="tableRow"><strong>New Polls:</strong></td>
     <td valign="top"  class="tableRow">Allows the Group to create new polls </td>
    </tr>
    <tr>
     <td align="right" valign="top"  class="tableRow"><strong>Poll Vote:</strong></td>
     <td valign="top"  class="tableRow">Allows the Group to vote in polls <br />
      <span class="smText">If you allow Guest Groups to vote in Polls, only cookies prevent Guests from multiple voting.</span></td>
    </tr>
    <tr>
     <td align="right" valign="top"  class="tableRow"><strong>Calendar Event:</strong></td>
     <td valign="top"  class="tableRow">Allows the Group to enter Topics into the Calendar system as an event to be displayed in the Calendar.<br />
      <span class="smText">The Calendar System needs to be enabled from the '<a href="admin_calendar_configuration.asp<% = strQsSID1 %>" class="smLink">Calendar Settings</a>' Page </span></td>
    </tr>
    <tr>
     <td align="right" valign="top"  class="tableRow"><strong>Post Approval:<br /></strong></td>
     <td valign="top" class="tableRow"><span class="text">Requires that posts for this Group need to be  approved before they are displayed <span class="smText"><br />
      If you choose to not let users have there posts displayed, then their posts will first need to be approved by the forum admin/moderator.</span> </span> </td>
    </tr>
    <tr>
     <td align="right" valign="top"  class="tableRow"><strong>Forum Moderator:<br />
      </strong></td>
     <td valign="top" class="tableRow"><span class="text">Allows the Group to have Moderator rights in this forum<br />
      </span><span class="smText">This will allow the group to be able to delete, edit, move, etc. all posts in this forum, and edit user profiles etc. across the entire board </span></td>
    </tr>
   </table>
<div align="center"><br />
  <span class="text">Please be aware that the Group Permissions can be over ridden by setting permissions on this forum for individual members.<br />
  <br />
  <br />
  <a name="lad" id="lad"></a><br />
  </span>
  <table width="100%" border="0" cellpadding="2" cellspacing="1" class="tableBorder">
    <tr>
      <td align="center" class="tableLedger">What is the Ladder System?</td>
    </tr>
    <tr>
      <td class="tableRow">The Ladder system enables your members to move up forum groups automatically depending on the number of posts they make. Once a member has made the minimum amount of posts for a Ladder User Group that member will be moved up to that user group.<br />
        <br />
        If you select that a user group is a Non Ladder Group, any member of the group will not be effected by the ladder system, this is useful if you wish not to use the Ladder System or for special groups like moderator groups.</td>
    </tr>
  </table>
  <br />
  <span class="text"> </span></div>
<!-- #include file="includes/admin_footer_inc.asp" -->
