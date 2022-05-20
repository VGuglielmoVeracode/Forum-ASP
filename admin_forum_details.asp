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
Dim strForumName	'Holds the name of the forum
Dim strForumDescription	'Holds the discription of the forum
Dim strForumPassword	'Holds the forum password
Dim strForumCode	'Holds a security code for the forum if it is password protected
Dim strCatName		'Holds the name of the category
Dim intCatID		'Holds the ID number of the category
Dim intSelCatID		'Holds the selected cat id
Dim intSubID		'Holds the sub forum ID
Dim blnLocked
Dim blnHide
Dim intShowTopicsFrom	'Holds the amount of time to show topics in
Dim intUserGroupID	'Holds the group ID
Dim intMainForumID	'Holds the ID of the main forum is sub forum mode
Dim strUserCode		'Holds user code
Dim intForumOrderNum	'Holds the forum order number
Dim blnSub
Dim strForumImage





'Initilise variables
intCatID = 0
intShowTopicsFrom = 0
intForumID = 0
blnLocked = False
blnHide = False
intForumOrderNum = 0


'Read in the details
intForumID = IntC(Request.QueryString("FID"))
strForumPassword = LCase(Request.Form("password"))



blnSub = BoolC(Request("sub"))



'Intialise the ADO recordset object
Set rsCommon2 = Server.CreateObject("ADODB.Recordset")




'If this is a post back update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))


	'If this is a sub forum do things a little different
	If blnSub Then

		'Read in the sub ID
		intMainForumID = IntC(Request.Form("mainForum"))
		intSubID = intMainForumID

		'Get the Cat ID for this sub forum from the database
		strSQL = "SELECT " & strDbTable & "Forum.Cat_ID From " & strDbTable & "Forum WHERE " & strDbTable & "Forum.Forum_ID= " & intMainForumID & ";"

		'Query the database
		rsCommon.Open strSQL, adoCon

		'If there is a record get the Cat ID
		If NOT rsCommon.EOF Then intCatID = CInt(rsCommon("Cat_ID"))

		'Reset rsCommon
		rsCommon.Close

	'Else main forums
	Else
		
		'Set the sub ID to 0
		intSubID = 0

		'Read in the cat ID for main forums
		intCatID = IntC(Request.Form("cat"))
		
		'If this is an edit then update any sub forums
		If intForumID > 0 Then
			'If there are any sub forums we need to update the Cat ID for them
			strSQL = "UPDATE " & strDbTable & "Forum " & _
			"SET Cat_ID = " & intCatID & " " & _
			"WHERE (Sub_ID = "  & intForumID & ");"
				
			'Write to database
			adoCon.Execute(strSQL)
		End If
	End If



	'If this is new we need a different query and also need to get the number of forum in cat for forum order
	If intForumID = 0 Then
		
		'SQL to colunt the number of forum in this cat
		strSQL = "SELECT Count(" & strDbTable & "Forum.Forum_ID) AS forumCount " & _
		"FROM " & strDbTable & "Forum " & _
		"WHERE " & strDbTable & "Forum.Cat_ID = " & intCatID & ";"
		
		'Query the database
		rsCommon.Open strSQL, adoCon
		
		'Get the number of forums in cat
		If NOT rsCommon.EOF Then intForumOrderNum = CInt(rsCommon("forumCount")) + 1
		
		'Close recordset
		rsCommon.Close
		
		'Initalise the strSQL variable with an SQL statement to query the database
		strSQL = "SELECT " & strDbTable & "Forum.* " & _
		"From " & strDbTable & "Forum " & _
		"ORDER BY " & strDbTable & "Forum.Forum_ID DESC;"
	
	'Else use a different query if we are updating
	Else
		strSQL = "SELECT " & strDbTable & "Forum.Cat_ID, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Forum_description, " & strDbTable & "Forum.Locked, " & strDbTable & "Forum.Hide, " & strDbTable & "Forum.Show_topics, " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Forum_code, " & strDbTable & "Forum.Forum_icon " & _
		"From " & strDbTable & "Forum " & _
		"WHERE " & strDbTable & "Forum.Forum_ID = " & intForumID & ";"
	End If

	'Set the cursor type property of the record set to Forward Only
	rsCommon.CursorType = 0

	'Set the Lock Type for the records so that the record set is only locked when it is updated
	rsCommon.LockType = 3

	'Query the database
	rsCommon.Open strSQL, adoCon

	With rsCommon
		'If this is a new one add new
		If intForumID = 0 Then .AddNew

		'Update the recordset
		.Fields("Cat_ID") = intCatID
		.Fields("Sub_ID") = intSubID
		.Fields("Forum_name") = Request.Form("forumName")
		.Fields("Forum_icon") = Request.Form("forumIcon")
		.Fields("Forum_description") = Request.Form("description")
		.Fields("Locked") = BoolC(Request.Form("locked"))
		.Fields("Hide") = BoolC(Request.Form("hide"))
		.Fields("Show_topics") = IntC(Request.Form("showTopics"))
		
		'See if there is a password if not the filed must be null
		If Request.Form("remove") Then
			.Fields("Password") = null
			.Fields("Forum_code") = null
		
		'Add the new or updated password and usercode to the database
		ElseIf strForumPassword <> "" Then

			'Encrypt the forum password
			strForumPassword = HashEncode(strForumPassword)

			'Calculate a code for the forum
			strForumCode = LCase(hexValue(32))

			'Place in recordset
			.Fields("Password") = strForumPassword
			.Fields("Forum_code") = strForumCode
		End If
		
		'If new database set some default values
		If intForumID = 0 Then 
			.Fields("Forum_Order") = intForumOrderNum
			.Fields("Last_post_author_ID") = lngLoggedInUserID   'Changed to use the admins logged in ID number to prevent errors of forums not displaying if the built in admin account is deleted
			.Fields("Last_post_date") = internationalDateTime("2001-01-01 00:00:00")
			.Fields("No_of_topics") = 0
			.Fields("No_of_posts") = 0
		End If
		

		'Update the database with the new user's details
		.Update
	End With


	'Re-run the query to read in the updated recordset from the database
	'We need to do this to get the new forum ID
	rsCommon.Requery

	'Read in the new forum ID
	intForumID = CInt(rsCommon("Forum_ID"))


	'Close RS
	rsCommon.Close



	'Set the permissions for this forum

	'Read in the groups from db
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Group.Group_ID, " & strDbTable & "Group.Name, " & strDbTable & "Group.Starting_group FROM " & strDbTable & "Group ORDER BY " & strDbTable & "Group.Group_ID ASC;"

	'Query the database
	rsCommon.Open strSQL, adoCon

	'Loop through all the categories in the database
	Do while NOT rsCommon.EOF

		'Get the group ID
		intUserGroupID = CInt(rsCommon("Group_ID"))
		
		
		'Due to some issues when updating from previous versions delete permission before reseting
		strSQL = "DELETE " & strDbTable & "Permissions FROM " & strDbTable & "Permissions WHERE " & strDbTable & "Permissions.Group_ID = " & intUserGroupID & " AND " & strDbTable & "Permissions.Forum_ID = " & intForumID & ";"

		'Write to database
		adoCon.Execute(strSQL)
			

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
			.Fields("View_Forum") = BoolC(Request.Form("read" & intUserGroupID))
			.Fields("Post") = BoolC(Request.Form("topic" & intUserGroupID))
			.Fields("Priority_posts") = BoolC(Request.Form("sticky" & intUserGroupID))
			.Fields("Reply_posts") = BoolC(Request.Form("reply" & intUserGroupID))
			.Fields("Edit_posts") = BoolC(Request.Form("edit" & intUserGroupID))
			.Fields("Delete_posts") = BoolC(Request.Form("delete" & intUserGroupID))
			.Fields("Poll_create") = BoolC(Request.Form("polls" & intUserGroupID))
			.Fields("Vote") = BoolC(Request.Form("vote" & intUserGroupID))
			.Fields("Display_post") = BoolC(Request.Form("approve" & intUserGroupID))
			.Fields("Moderate") = BoolC(Request.Form("moderator" & intUserGroupID))
			.Fields("Calendar_event") = BoolC(Request.Form("calEvent" & intUserGroupID))
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


	'If this is a new forum go back to the main forums page
	If intForumID = 0 Then

		'Release server varaibles
		Set rsCommon2 = Nothing
		Call closeDatabase()

		Response.Redirect("admin_view_forums.asp" & strQsSID1)
	End If
End If




'If this is an edit read in the forum details
If intForumID > 0 Then

	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Forum.* " & _
	"From " & strDbTable & "Forum " & _
	"WHERE " & strDbTable & "Forum.Forum_ID=" & intForumID & ";"

	'Query the database
	rsCommon.Open strSQL, adoCon

	If NOT rsCommon.EOF Then

		'Read in the forums from the recordset
		intCatID = CInt(rsCommon("Cat_ID"))
		intSubID = CInt(rsCommon("Sub_ID"))
		strForumName = rsCommon("Forum_name")
		strForumDescription = rsCommon("Forum_description")
		intForumID = CInt(rsCommon("Forum_ID"))
		strForumPassword = rsCommon("Password")
		blnLocked = CBool(rsCommon("Locked"))
		blnHide = CBool(rsCommon("Hide"))
		intShowTopicsFrom = CInt(rsCommon("Show_topics"))
		strForumImage = rsCommon("Forum_icon")
	End If

	'Close the rs
	rsCommon.Close
End If




'See if there is a main forum for this forum to be in or if there is a category to place a main forum within
If blnSub Then
	'Read in the main forum name from the database
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Category.Cat_ID, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Forum_name " & _
	"FROM " & strDbTable & "Category, " & strDbTable & "Forum " & _
	"WHERE " & strDbTable & "Category.Cat_ID=" & strDbTable & "Forum.Cat_ID " & _
		"AND " & strDbTable & "Forum.Sub_ID=0 " & _
		"ORDER BY " & strDbTable & "Category.Cat_order ASC, " & strDbTable & "Forum.Forum_Order ASC;"
Else
	'Read in the category name from the database
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Category.Cat_name, " & strDbTable & "Category.Cat_ID " & _
	"FROM " & strDbTable & "Category " & _
	"ORDER BY " & strDbTable & "Category.Cat_order ASC;"
End If


'Query the database
rsCommon.Open strSQL, adoCon

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Forum Details</title>
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
<!-- Check the from is filled in correctly before submitting -->
<script  language="JavaScript" type="text/javascript">

//Function to check form is filled in correctly before submitting
function CheckForm () {

<%

'If this is a sub forum check a forum is selected
If blnSub Then
	%>
	//Check for a forum
	if (document.frmNewForum.mainForum.value==""){
		alert("Please select a Parent Forum for this Sub Forum is to be in");
		return false;
	}<%

'Else we need a category
Else
%>
	//Check for a category
	if (document.frmNewForum.cat.value==""){
		alert("Please select the Category this Forum is to be in");
		return false;
	}<%
End If

%>
	//Check for a forum name
	if (document.frmNewForum.forumName.value==""){
		alert("Please enter a Name for the Forum");
		document.frmNewForum.forumName.focus();
		return false;
	}

	//Check for a pforum description
	if (document.frmNewForum.description.value==""){
		alert("Please enter a Description for the Forum");
		document.frmNewForum.description.focus();
		return false;
	}

	return true
}
</script>
<!-- #include file="includes/admin_header_inc.asp" -->
   <div align="center"><h1>Forum Details</h1><br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <a href="admin_view_forums.asp<% = strQsSID1 %>">Return to the Forum Administration page</a><br />
    <br />
    <%

'If there is no main forum or cat then display a message
If rsCommon.EOF Then
 	%>
    <table width="98%" border="0" cellspacing="0" cellpadding="1" height="135">
     <tr>
      <td align="center" class="text"><span class="lgText">
       <%

	'Sub forum mode
	If blnSub Then
		%>
       You must first create a New Forum to place your new Sub Forum in.</span><br />
       <br />
       <a href="admin_forum_details.asp?mode=new<% = strQsSID2 %>">Create a New Forum</a>
       <%

	'Main forum mode
	Else
		%>
       You must first create a Forum Category to place your new Forum in.</span><br />
       <br />
       <a href="admin_category_details.asp?mode=new<% = strQsSID2 %>">Create a Forum Category</a>
       <%
	End If

%></td>
     </tr>
    </table>
    <%
Else
%>
   </div>
   <form action="admin_forum_details.asp?FID=<% = intForumID %><% = strQsSID2 %>" method="post" name="frmNewForum" id="frmNewForum" onsubmit="return CheckForm();">
    <%

	'If this is a sub forum check a forum is selected
	If blnSub Then
	%>
    <table align="center" cellpadding="4" cellspacing="1" class="tableBorder">
     <tr>
      <td colspan="2" class="tableLedger">Select Parent Forum </td>
     </tr>
     <tr>
      <td colspan="2" class="tableRow">Select the Parent Forum from the drop down list below that you would like this Sub Forum to be in.<br />
       <select name="mainForum">
        <option value=""<% If intSubID = 0 Then Response.Write(" selected") %>>-- Select Parent Forum --</option>
        <%

		'Loop through all the main forums in the database
		Do while NOT rsCommon.EOF

			'Read in the deatils for the main forum
			strMainForumName = rsCommon("Forum_name")
			intMainForumID = CInt(rsCommon("Forum_ID"))

			'Display a link in the link list to the cat
			Response.Write (vbCrLf & "		<option value=""" & intMainForumID & """")
			If intMainForumID = intSubID Then Response.Write(" selected")
			Response.Write(">" & strMainForumName & "</option>")


			'Move to the next record in the recordset
			rsCommon.MoveNext
		Loop

		'Close Rs
		rsCommon.Close
%>
       </select>
      </td>
     </tr>
    </table>
    <%

	'Else this is a main forum so select a category
	Else

%>
    <table align="center" cellpadding="4" cellspacing="1" class="tableBorder">
     <tr>
      <td colspan="2" class="tableLedger">Select Forum Category</td>
     </tr>
     <tr>
      <td colspan="2" class="tableRow">Select the Category from the drop down list below that you would like this Forum to be in.<br />
       <select name="cat">
        <option value=""<% If intCatID = 0 Then Response.Write(" selected") %>>-- Select Forum Category --</option>
        <%



		'Loop through all the categories in the database
		Do while NOT rsCommon.EOF

			'Read in the deatils for the category
			strCatName = rsCommon("Cat_name")
			intSelCatID = CInt(rsCommon("Cat_ID"))

			'Display a link in the link list to the cat
			Response.Write (vbCrLf & "		<option value=""" & intSelCatID & """")
			If intCatID = intSelCatID Then Response.Write(" selected")
			Response.Write(">" & strCatName & "</option>")


			'Move to the next record in the recordset
			rsCommon.MoveNext
		Loop

		'Close Rs
		rsCommon.Close
%>
       </select>
      </td>
     </tr>
    </table>
    <%
	End If

%>
    <br />
    <table align="center" cellpadding="4" cellspacing="1" class="tableBorder">
     <tr>
      <td colspan="2" class="tableLedger">Forum Details</td>
     </tr>
     <tr>
      <td width="50%" class="tableRow">Forum Name*:</td>
      <td width="50%" valign="top" class="tableRow"><input type="text" name="forumName" maxlength="60" size="40" value="<% = strForumName %>"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      </td>
     </tr>
     <tr>
      <td class="tableRow">Forum Description*:<br />
       <span class="smText">Give a brief description of the forum.</span></td>
      <td valign="top" class="tableRow"><input type="text" name="description" maxlength="190" size="70" value="<% = strForumDescription %>"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
      <tr>
      <td class="tableRow">Forum Image Icon:<br />
       <span class="smText">Use this to add a link to a custom image icon displayed on the Forum Index page, in place of the standard icons displayed next to forums.</span></td>
      <td valign="top" class="tableRow"><input type="text" id="forumIcon" name="forumIcon" maxlength="70" size="40" value="<% = strForumImage %>"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr>
      <td class="tableRow">Show Topics in the Last:<br />
       <span class="smText">This is the default time span in which topics containing new posts are shown in the topic list.</span></td>
      <td valign="top" class="tableRow"><select name="showTopics">
        <option value="0"<% If intShowTopicsFrom = 0 OR intShowTopicsFrom = "" Then Response.Write " selected" %>>Any Date</option>
        <option value="1"<% If intShowTopicsFrom = 1 Then Response.Write " selected" %>>User Last Login Date</option>
        <option value="2"<% If intShowTopicsFrom = 2 Then Response.Write " selected" %>>Yesterday</option>
        <option value="3"<% If intShowTopicsFrom = 3 Then Response.Write " selected" %>>Last Two Days</option>
        <option value="4"<% If intShowTopicsFrom = 4 Then Response.Write " selected" %>>Last Week</option>
        <option value="5"<% If intShowTopicsFrom = 5 Then Response.Write " selected" %>>Last Month</option>
        <option value="6"<% If intShowTopicsFrom = 6 Then Response.Write " selected" %>>Last Two Months</option>
        <option value="7"<% If intShowTopicsFrom = 7 Then Response.Write " selected" %>>Last Six Months</option>
        <option value="8"<% If intShowTopicsFrom = 8 Then Response.Write " selected" %>>Last Year</option>
       </select>
       </select></td>
     </tr>
     <tr>
      <td class="tableRow">Forum Locked:<br />
        <span class="smText">If the forum is locked posts can not be made in the forum. Useful for maintenance.</span></td>
      <td valign="top" class="tableRow"><input name="locked" type="checkbox" id="locked2" value="true"<% If blnLocked Then Response.Write " checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr>
      <td class="tableRow">Hide Forum if no access:<br />
       <span class="smText">Hide this forum on the boards main page if the user can not access the forum.</span></td>
      <td valign="top" class="tableRow"><input name="hide" type="checkbox" id="hide" value="true"<% If blnHide Then Response.Write " checked" %><% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr>
      <td  class="tableRow"><% If strForumPassword <> "" Then Response.Write("Change ")%>
       Password:<br />
       <span class="smText">If you want this forum password protected place the password here. Otherwise leave it blank.</span><br />
       <br />
       <span class="smText"><strong>Please note</strong>: A password is required to view posts in a password protected forum, but Topic Subjects will still display in Searches and the Active Users Page, for better security consider using forum permissions in the table below. </span> </td>
      <td valign="top" class="tableRow"><input type="text" name="password" maxlength="20" size="20"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
       <%

      	'If there is a password ask if they want to remove it
      	If strForumPassword <> "" Then
      	%>
       <br />
       <input name="remove" type="checkbox" id="remove" value="true"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
       Forum is password protected check box to 'Remove Forum Password'
       <%

	End If

      	%></td>
     </tr>
    </table>
    <br />
    <br />
    <table width="100%" align="center" cellpadding="0" cellspacing="0">
     <tr>
      <td align="center" class="text"><span class="lgText">Forum Permissions</span><br />
       Use the grid below to set  Permissions for the various Member Groups on this forum.<br />
       <a href="#per" class="smLink">What do the different permissions mean?</a><br />
       <br />
       <table cellpadding="2" cellspacing="1" class="tableBorder">
        <tr>
         <td width="194" align="left" class="tableLedger">Member Group</td>
         <td width="43" align="center" class="tableLedger">Access</td>
         <td width="43" align="center" class="tableLedger">New Topics</td>
         <td width="43" align="center" class="tableLedger">Sticky Topics</td>
         <td width="43" align="center" class="tableLedger">Post Reply</td>
         <td width="43" align="center" class="tableLedger">Edit Posts</td>
         <td width="43" align="center" class="tableLedger">Delete Posts</td>
         <td width="43" align="center" class="tableLedger">New Polls</td>
         <td width="43" align="center" class="tableLedger">Poll Vote</td>
         <td width="43" align="center" class="tableLedger">Calendar Event</td>
         <td width="43" align="center" class="tableLedger">Post Approval</td>
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
	strSQL = "SELECT " & strDbTable & "Group.Group_ID, " & strDbTable & "Group.Name, " & strDbTable & "Group.Starting_group FROM " & strDbTable & "Group ORDER BY " & strDbTable & "Group.Group_ID ASC;"

	'Query the database
	rsCommon.Open strSQL, adoCon

	'Loop through all the categories in the database
	Do while NOT rsCommon.EOF

		'Get the group ID
		intUserGroupID = CInt(rsCommon("Group_ID"))

		'Read in the permssions from the db for this group (not very efficient doing it this way, but this page won't be run often)
		'Initalise the strSQL variable with an SQL statement to query the database
		strSQL = "SELECT " & strDbTable & "Permissions.* FROM " & strDbTable & "Permissions WHERE " & strDbTable & "Permissions.Group_ID = " & intUserGroupID & " AND " & strDbTable & "Permissions.Forum_ID = " & intForumID & ";"

		'Query the database
		rsCommon2.Open strSQL, adoCon

		'If no records are returned use default values
		If rsCommon2.EOF Then

%>
        <tr>
         <td align="left" class="tableRow"><% = rsCommon("Name") %>
          <% If CBool(rsCommon("Starting_group")) Then Response.Write("<br /><span class=""smText"">(New members group)</span>") %>
          <% If intUserGroupID = 2 Then Response.Write("<br /><span class=""smText"">(Un-registered users)</span>") %></td>
         <td align="center" class="tableRow"><input name="read<% = intUserGroupID %>" type="checkbox" id="read" value="true" checked /></td>
         <td align="center" class="tableRow"><input name="topic<% = intUserGroupID %>" type="checkbox" id="topic" value="true"<% If intUserGroupID <> 2 Then Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="sticky<% = intUserGroupID %>" type="checkbox" id="sticky" value="true"<% If intUserGroupID = 2 Then Response.Write(" disabled=true") %><% If intUserGroupID = 1 Then Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="reply<% = intUserGroupID %>" type="checkbox" id="reply" value="true"<% If intUserGroupID <> 2 Then Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="edit<% = intUserGroupID %>" type="checkbox" id="edit" value="true"<% If intUserGroupID = 2 Then Response.Write(" disabled=true") Else Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="delete<% = intUserGroupID %>" type="checkbox" id="delete" value="true"<% If intUserGroupID = 2 Then Response.Write(" disabled=true") Else Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="polls<% = intUserGroupID %>" type="checkbox" id="polls" value="true" /></td>
         <td align="center" class="tableRow"><input name="vote<% = intUserGroupID %>" type="checkbox" id="vote" value="true" /></td>
         <td align="center" class="tableRow"><input name="calEvent<% = intUserGroupID %>" type="checkbox" id="calEvent" value="true" /></td>
         <td align="center" class="tableRow"><input name="approve<% = intUserGroupID %>" type="checkbox" id="approve" value="true" /></td>
         <td align="center" class="tableRow"><input name="moderator<% = intUserGroupID %>" type="checkbox" id="moderator" value="true"<% If intUserGroupID = 2 Then Response.Write(" disabled=true") %> /></td>
        </tr>
        <%

      		'Else display the values for this group
      		Else
%>
        <tr>
         <td align="left" class="tableRow"><% = rsCommon("Name") %>
          <% If CBool(rsCommon("Starting_group")) Then Response.Write("<br /><span class=""smText"">(New members group)</span>") %>
          <% If intUserGroupID = 2 Then Response.Write("<br /><span class=""smText"">(Un-registered users)</span>") %></td>
         <td align="center" class="tableRow"><input name="read<% = intUserGroupID %>" type="checkbox" id="read" value="true"<% If CBool(rsCommon2("View_Forum")) Then Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="topic<% = intUserGroupID %>" type="checkbox" id="topic" value="true"<% If CBool(rsCommon2("Post")) Then Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="sticky<% = intUserGroupID %>" type="checkbox" id="sticky" value="true"<% If intUserGroupID = 2 Then Response.Write(" disabled=true") %><% If CBool(rsCommon2("Priority_posts")) Then Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="reply<% = intUserGroupID %>" type="checkbox" id="reply" value="true"<% If CBool(rsCommon2("Reply_posts")) Then Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="edit<% = intUserGroupID %>" type="checkbox" id="edit"<% If intUserGroupID = 2 Then Response.Write(" disabled=true") %> value="true"<% If CBool(rsCommon2("Edit_posts")) Then Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="delete<% = intUserGroupID %>" type="checkbox" id="delete"<% If intUserGroupID = 2 Then Response.Write(" disabled=true") %> value="true"<% If CBool(rsCommon2("Delete_posts")) Then Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="polls<% = intUserGroupID %>" type="checkbox" id="polls" value="true"<% If CBool(rsCommon2("Poll_create")) Then Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="vote<% = intUserGroupID %>" type="checkbox" id="vote" value="true"<% If CBool(rsCommon2("Vote")) Then Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="calEvent<% = intUserGroupID %>" type="checkbox" id="calEvent" value="true"<% If CBool(rsCommon2("Calendar_event")) Then Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="approve<% = intUserGroupID %>" type="checkbox" id="approve" value="true"<% If CBool(rsCommon2("Display_post")) Then Response.Write(" checked") %> /></td>
         <td align="center" class="tableRow"><input name="moderator<% = intUserGroupID %>" type="checkbox" id="moderator" value="true"<% If intUserGroupID = 2 Then Response.Write(" disabled=true") %><% If CBool(rsCommon2("Moderate")) Then Response.Write(" checked") %> /></td>
        </tr>
        <%

		End If

		'Close rsCommon2
		rsCommon2.Close

		'Move to the next record in the recordset
		rsCommon.MoveNext
	Loop

	'Close Rs
	rsCommon.Close

%>
       </table></td>
     </tr>
    </table>
    <div align="center"><br />
     <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
     <input type="hidden" name="postBack" value="true" />
     <input type="hidden" name="sub" value="<% = blnSub %>" />
     <input type="submit" name="Submit" value="Submit Forum Details"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
     <input type="reset" name="Reset" value="Reset Form" />
     <br />
    </div>
   </form>
   <%
End If


'Reset Server Objects
Set rsCommon2 = Nothing
Call closeDatabase()
%>
   <br />
   <a name="per" id="per"></a> <br />
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
    <span class="text">Please be aware that the Group Permissions can be over ridden by setting permissions on this forum for individual members.</span><br />
    <br />
   </div>
   <!-- #include file="includes/admin_footer_inc.asp" -->
