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
Dim strForumName	'Holds the name of the forum
Dim strMemberName	'Holds the name of the forum member
Dim lngMemberID		'Holds the ID number of the member
Dim intSelGroupID	'Holds the group ID to select
Dim iaryForumID		'Holds the forum ID array
Dim intCatID		'Holds the cat ID
Dim sarryForums
Dim intCurrentRecord
Dim sarrySubForums
Dim intCurrentRecord2
Dim intSubForumID


'Read in the details
lngMemberID = LngC(Request("UID"))


'Don't let em edit the Guests member, as some people change the qurystring and then get 'em selfs in trouble!!
If lngMemberID = 2 Then
	
	'Close DB
	Call closeDatabase()
	
	'Redirect back to previous page
	Response.Redirect("admin_find_user.asp")
End If







'Read in the member name
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Author.Username From " & strDbTable & "Author WHERE " & strDbTable & "Author.Author_ID=" & lngMemberID & ";"

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the forum name form the recordset
If NOT rsCommon.EOF Then

	'Read in the forums from the recordset
	strMemberName = rsCommon("Username")
End If

'Release server varaibles
rsCommon.Close




'If this is a post back update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
	
	'Run through till all checked forums are added
	For each iaryForumID in Request.Form("chkFID")


		'Initalise the strSQL variable with an SQL statement to query the database
		strSQL = "SELECT " & strDbTable & "Permissions.* From " & strDbTable & "Permissions WHERE " & strDbTable & "Permissions.Forum_ID=" & iaryForumID & " AND " & strDbTable & "Permissions.Author_ID = " & lngMemberID & ";"
		
		'Set the Lock Type for the records so that the record set is only locked when it is updated
		rsCommon.LockType = 3
		
		'Query the database
		rsCommon.Open strSQL, adoCon
	
		With rsCommon
			'If this is a new one add new
			If rsCommon.EOF Then .AddNew
	
			'Update the recordset
			.Fields("Forum_ID") = iaryForumID
			.Fields("Author_ID") = lngMemberID
			.Fields("View_Forum") = BoolC(Request.Form("read"))
			.Fields("Post") = BoolC(Request.Form("post"))
			.Fields("Reply_posts") = BoolC(Request.Form("reply"))
			.Fields("Edit_posts") = BoolC(Request.Form("edit"))
			.Fields("Delete_posts") = BoolC(Request.Form("delete"))
			.Fields("Priority_posts") = BoolC(Request.Form("priority"))
			.Fields("Poll_create") = BoolC(Request.Form("poll"))
			.Fields("Vote") = BoolC(Request.Form("vote"))
			.Fields("Attachments") = BoolC(Request.Form("files"))
			.Fields("Image_upload") = BoolC(Request.Form("images"))
			.Fields("Moderate") = BoolC(Request.Form("moderate"))
			.Fields("Display_post") = BoolC(Request.Form("display"))
			.Fields("Calendar_event") = BoolC(Request.Form("calEvent"))
			.Fields("Attachments") = False
			.Fields("Image_upload") = False
	
			'Update the database with the new user's details
			.Update
			
			'Close recordset
			.close
		End With
	Next
	

	'Release server varaibles
	Call closeDatabase()

	'Redirect back to permissions page
	Response.Redirect("admin_user_permissions.asp?UID=" & lngMemberID & strQsSID3)
End If


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Create Member Permissions</title>
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
<div align="center"><h1>Create Member Permissions for <% = strMemberName %></h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <a href="admin_find_user.asp<% = strQsSID1 %>">Select another Member to Create, Edit, or Delete  Permissions for </a><br />
  <br />
  <span class="text">Use the form below to create permissions on Forums for the member <b>
  <% = strMemberName %>
  </b>.<br />
  <b>These member permissions override any Group Permissions the member would have on  forums. </b> </span></div>
<form action="admin_user_permissions.asp?UID=<% = lngMemberID %><% = strQsSID2 %>" method="post" name="frmNewForum" id="frmNewForum">
  <div align="center"><span class="text"><span class="lgText"><br />
    Permissions</span><br />
    Select what permissions you require for this user </span><br />
    <span class="text"><a href="#per" class="smLink">What do the different permissions mean?</a></span> <br />
    <br />
    <table border="0" align="center" cellpadding="2" cellspacing="1" class="tableBorder">
      <tr>
        <td align="left" width="194" class="tableLedger">Member Name </td>
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
        <td width="43" align="center" class="tableLedger"> Forum Moderator</td>
      </tr>
      <tr>
        <td align="left" class="tableRow"><% = strMemberName %></td>
        <td align="center" class="tableRow"><input name="read" type="checkbox" value="true" checked="checked" /></td>
        <td align="center" class="tableRow"><input name="post" type="checkbox" value="true" checked="checked" /></td>
        <td align="center" class="tableRow"><input type="checkbox" name="priority" value="true" /></td>
        <td align="center" class="tableRow"><input name="reply" type="checkbox" value="true" checked="checked" /></td>
        <td align="center" class="tableRow"><input name="edit" type="checkbox" value="true" checked="checked" /></td>
        <td align="center" class="tableRow"><input name="delete" type="checkbox" value="true" checked="checked" /></td>
        <td align="center" class="tableRow"><input type="checkbox" name="poll" value="true" /></td>
        <td align="center" class="tableRow"><input type="checkbox" name="vote" value="true" /></td>
        <td align="center" class="tableRow"><input type="checkbox" name="calEvent" value="true" /></td>
        <td align="center" class="tableRow"><input type="checkbox" name="display" value="true" /></td>
        <td align="center" class="tableRow"><input type="checkbox" name="moderate" value="true" /></td>
      </tr>
    </table>
    <br />
    <span class="text"><span class="lgText">Forums</span><br />
    Select which forums you would like to apply/modify these permissions on</span><br />
    <br />
    <table border="0" align="center" cellpadding="2" cellspacing="1" class="tableBorder">
      <tr>
        <td width="37" class="tableLedger">&nbsp;</td>
        <td class="tableLedger">Member Group</td>
      </tr>
      <%

'Read in the groups from db
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Category.Cat_ID, " & strDbTable & "Category.Cat_name, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name FROM " & strDbTable & "Category, " & strDbTable & "Forum WHERE " & strDbTable & "Category.Cat_ID=" & strDbTable & "Forum.Cat_ID AND " & strDbTable & "Forum.Sub_ID=0 ORDER BY " & strDbTable & "Category.Cat_order ASC, " & strDbTable & "Category.Cat_ID ASC, " & strDbTable & "Forum.Forum_Order ASC;"

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the row from the db using getrows for better performance
If NOT rsCommon.EOF Then 
	sarryForums = rsCommon.GetRows()
End If

'close
rsCommon.Close

'If no forums to set permisisons on display a message saying so
If NOT isArray(sarryForums) Then
%>
	  <tr>
            <td align="left" class="tableRow" colspan="14">There are presently no Forums created to set permissions on</td>
          </tr><%

'If there are results show them
Else

	'Loop round to read in all the forums in the database
	Do While intCurrentRecord <= Ubound(sarryForums,2)

		'Get the forum ID
		intForumID = CInt(sarryForums(2,intCurrentRecord))
		
		'If this is a different cat display the cat ID
		If intCatID <> CInt(sarryForums(0,intCurrentRecord)) Then 
				
			'Change the cat ID
			intCatID = CInt(sarryForums(0,intCurrentRecord))
			
			%>
      <tr>
        <td align="left" class="tableSubLedger" colspan="13"><% = sarryForums(1,intCurrentRecord) %></td>
      </tr>
      <%
       
	End If
	
	%>
      <tr>
        <td align="center" class="tableRow"><input type="checkbox" name="chkFID" id="chkFID" value="<% = intForumID %>" /></td>
        <td align="left" class="tableRow"><% = sarryForums(3,intCurrentRecord) %></td>
      </tr>
      <%
	
        
	        '********* check for sub forums *****************
	        
	        'Reset intCurrentRecord2
		intCurrentRecord2 = 0
	        
	        'Read in the groups from db
		'Initalise the strSQL variable with an SQL statement to query the database
		strSQL = "SELECT " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name FROM " & strDbTable & "Forum WHERE " & strDbTable & "Forum.Sub_ID= " & intForumID & " ORDER BY " & strDbTable & "Forum.Forum_Order ASC;"
		
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
			Do While NOT intCurrentRecord2 > Ubound(sarrySubForums,2)
		
				'Get the forum ID
				intSubForumID = CInt(sarrySubForums(0,intCurrentRecord2))
		
%>
      <tr>
        <td align="center" class="tableRow"><input type="checkbox" name="chkFID" id="chkFID" value="<% = intSubForumID %>" /></td>
        <td align="left" class="tableRow">&nbsp;<img src="<% = strImagePath %>arrow.gif" />&nbsp;<% = sarrySubForums(1,intCurrentRecord2) %></td>
      </tr>
      <%
		
				'Move to the next record in the recordset
				intCurrentRecord2 = intCurrentRecord2 + 1
			Loop
		End If
	        
		'Move to the next record in the recordset
		intCurrentRecord = intCurrentRecord + 1
	Loop

End If
%>
    </table>
    <br />
    <input type="hidden" name="postBack" value="true" />
    <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
    <input type="submit" name="Submit" value="Create/Update Member Permissions"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
    <input type="reset" name="Reset" value="Reset Form" />
    <br />
  </div>
</form>
<div align="center" class="text"><br />
  <h1><br />
  Present Member Permissions For <% = strMemberName %></h1> <br />
  Below are the present member forum permissions for <b>
  <% = strMemberName %>
  .<br />
  These member permissions override any Group Permissions the member would have on these forums. </b><br />
  <a href="#per" class="smLink">What do the different permissions mean?</a><br />
  <br />
  <table border="0" cellpadding="2" cellspacing="1" class="tableBorder">
    <tr>
      <td align="left" width="144" class="tableLedger">Forum</td>
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
      <td width="43" align="center" class="tableLedger">&nbsp;</td>
    </tr>
    <%

'Reset record position holders
intCurrentRecord = 0
intCurrentRecord2 = 0

'Read in the groups from db
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Category.Cat_ID, " & strDbTable & "Category.Cat_name, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name " & _
"FROM " & strDbTable & "Category, " & strDbTable & "Forum " & _
"WHERE " & strDbTable & "Category.Cat_ID=" & strDbTable & "Forum.Cat_ID " & _
	"AND " & strDbTable & "Forum.Sub_ID=0 " & _
"ORDER BY " & strDbTable & "Category.Cat_order ASC, " & strDbTable & "Forum.Forum_Order ASC;"

'Query the database
rsCommon.Open strSQL, adoCon

'Read in the row from the db using getrows for better performance
If NOT rsCommon.EOF Then sarryForums = rsCommon.GetRows()


'close
rsCommon.Close

'If there are results show them
If isArray(sarryForums) Then

	'Loop round to read in all the forums in the database
	Do While intCurrentRecord <= Ubound(sarryForums,2)

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
		strSQL = "SELECT " & strDbTable & "Permissions.* FROM " & strDbTable & "Permissions WHERE " & strDbTable & "Permissions.Author_ID = " & lngMemberID & " AND " & strDbTable & "Permissions.Forum_ID = " & intForumID & ";"
			
		'Query the database
		rsCommon.Open strSQL, adoCon
			
		'If no records are returned use default values
		If rsCommon.EOF Then		

%>
    <tr>
      <td align="left" class="tableRow"><% = sarryForums(3,intCurrentRecord) %></td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
    </tr>
    <%
      
		'Else display the values for this group
		 Else
%>
    <tr>
      <td align="left" class="tableRow"><% = sarryForums(3,intCurrentRecord) %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("View_Forum")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Post")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Priority_posts")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Reply_posts")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Edit_posts")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Delete_posts")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Poll_create")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Vote")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Calendar_event")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Display_post")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Moderate")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><a href="admin_remove_permissions.asp?FID=<% = intForumID %>&amp;UID=<% = lngMemberID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 %>" onclick="return confirm('Are you sure you want to Remove these Forum Permissions for this member?')"><img src="<% = strImagePath %>delete.png" width="15" height="16" border="0" alt="Remove" /></a></td>
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
		strSQL = "SELECT " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name FROM " & strDbTable & "Forum WHERE " & strDbTable & "Forum.Sub_ID= " & intForumID & " ORDER BY " & strDbTable & "Forum.Forum_Order ASC;"
		
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
			Do While NOT intCurrentRecord2 > Ubound(sarrySubForums,2)
		
				'Get the forum ID
				intSubForumID = CInt(sarrySubForums(0,intCurrentRecord2))
			
			
				'Read in the permssions from the db for this group (not very efficient doing it this way, but this page won't be run often)
				'Initalise the strSQL variable with an SQL statement to query the database
				strSQL = "SELECT " & strDbTable & "Permissions.* FROM " & strDbTable & "Permissions WHERE " & strDbTable & "Permissions.Author_ID = " & lngMemberID & " AND " & strDbTable & "Permissions.Forum_ID = " & intSubForumID & ";"
					
				'Query the database
				rsCommon.Open strSQL, adoCon
					
				'If no records are returned use default values
				If rsCommon.EOF Then		

%>
    <tr>
      <td align="left" class="tableRow">&nbsp;<img src="<% = strImagePath %>arrow.gif" />&nbsp;<% = sarrySubForums(1,intCurrentRecord2) %></td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
      <td align="center" class="tableRow">&nbsp;</td>
    </tr>
    <%
      
				'Else display the values for this group
	 			Else
%>
    <tr>
      <td align="left" class="tableRow">&nbsp;<img src="<% = strImagePath %>arrow.gif" />&nbsp;<% = sarrySubForums(1,intCurrentRecord2) %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("View_Forum")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Post")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Priority_posts")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Reply_posts")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Edit_posts")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Delete_posts")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Poll_create")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Vote")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Calendar_event")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Display_post")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><% If CBool(rsCommon("Moderate")) Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""Yes"" width=""13"" height=""14"">") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""No"" width=""13"" height=""14"">") %></td>
      <td align="center" class="tableRow"><a href="admin_remove_permissions.asp?FID=<% = intSubForumID %>&amp;UID=<% = lngMemberID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 %>" onclick="return confirm('Are you sure you want to Remove these Forum Permissions for this member?')"><img src="<% = strImagePath %>delete.png" width="15" height="16" border="0" alt="Remove" /></a></td>
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

%>
  </table>
  <br />
  <br />
  <br />
  <%


'Reset Server Objects
Call closeDatabase()
%>
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
      <td width="76%" valign="top" align="left" class="tableRow">Allows the Member access to the forum </td>
    </tr>
    <tr>
      <td width="24%" align="right" valign="top"  class="tableRow"><strong>New Topics:</strong></td>
      <td width="76%" valign="top" align="left" class="tableRow">Allows the Member to post new topics </td>
    </tr>
    <tr>
      <td align="right" valign="top"  class="tableRow"><strong>Sticky Topics:</strong></td>
      <td valign="top" align="left" class="tableRow">Allows the Member to post sticky topics that remain at the top of the forum </td>
    </tr>
    <tr>
      <td width="24%" align="right" valign="top"  class="tableRow"><strong>Post Reply:<br />
        </strong></td>
      <td width="76%" valign="top" align="left" class="tableRow">Allows the Member to reply to posts </td>
    </tr>
    <tr>
      <td align="right" valign="top"  class="tableRow"><strong>Edit Posts:</strong></td>
      <td valign="top" align="left" class="tableRow">Allows the Member to edit their posts </td>
    </tr>
    <tr>
      <td align="right" valign="top"  class="tableRow"><strong>Delete Posts:</strong></td>
      <td valign="top" align="left" class="tableRow">Allows the User to delete their posts, but only if no-one has posted a reply </td>
    </tr>
    <tr>
      <td align="right" valign="top"  class="tableRow"><strong>New Polls:</strong></td>
      <td valign="top" align="left" class="tableRow">Allows the Member to create new polls </td>
    </tr>
    <tr>
      <td align="right" valign="top"  class="tableRow"><strong>Poll Vote:</strong></td>
      <td valign="top" align="left" class="tableRow">Allows the Member to vote in polls </td>
    </tr>
    <tr>
    <tr>
     <td align="right" valign="top"  class="tableRow"><strong>Calendar Event:</strong></td>
     <td valign="top"  class="tableRow">Allows the Group to enter Topics into the Calendar system as an event to be displayed in the Calendar.<br />
      <span class="smText">The Calendar System needs to be enabled from the '<a href="admin_calendar_configuration.asp<% = strQsSID1 %>" class="smLink">Calendar Settings</a>' Page </span></td>
    </tr>
    <tr>
      <td align="right" valign="top"  class="tableRow"><strong>Post Approval:<br />
        </strong></td>
      <td valign="top" align="left" class="tableRow">Requires that posts for this Member need to be  approved before they are displayed <span class="smText"><br />
      If you choose to not let users have there posts displayed, then their posts will first need to be approved by the forum admin/moderator.</span> </td>
    </tr>
    <tr>
      <td align="right" valign="top"  class="tableRow"><strong>Forum Moderator:<br />
        </strong></td>
      <td valign="top" align="left" class="tableRow">Allows the Member to have Moderator rights in this forum<br />
      <span class="smText">This will allow the group to be able to delete, edit, move, etc. all posts in this forum, and edit user profiles etc. across the entire board </span></td>
    </tr>
  </table>
  <br />
</div>
<!-- #include file="includes/admin_footer_inc.asp" -->
