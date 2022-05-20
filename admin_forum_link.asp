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
Dim intUserGroupID	'Holds the group ID
Dim intMainForumID	'Holds the ID of the main forum is sub forum mode
Dim intForumOrderNum	'Holds the forum order number
Dim strForumImage
Dim strForumURL





'Initilise variables
intCatID = 0
intForumID = 0
intForumOrderNum = 0
strForumURL = "http://"


'Read in the details
intForumID = IntC(Request.QueryString("FID"))



'Intialise the ADO recordset object
Set rsCommon2 = Server.CreateObject("ADODB.Recordset")




'If this is a post back update the database
If Request.Form("postBack") AND blnDemoMode = False Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))


	

	'Read in the cat ID for main forums
	intCatID = IntC(Request.Form("cat"))
		

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
		strSQL = "SELECT " & strDbTable & "Forum.Cat_ID, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Forum_description, " & strDbTable & "Forum.Locked, " & strDbTable & "Forum.Hide, " & strDbTable & "Forum.Show_topics, " & strDbTable & "Forum.Password, " & strDbTable & "Forum.Forum_code, " & strDbTable & "Forum.Forum_icon, " & strDbTable & "Forum.Forum_URL " & _
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
		.Fields("Sub_ID") = 0
		If intForumID = 0 Then .Fields("Forum_Order") = intForumOrderNum
		If intForumID = 0 Then .Fields("Last_post_author_ID") = lngLoggedInUserID   'Changed to use the admins logged in ID number to prevent errors of forums not displaying if the built in admin account is deleted
		If intForumID = 0 Then .Fields("Last_post_date") = internationalDateTime("2001-01-01 00:00:00")
		.Fields("Forum_URL") = Request.Form("forumURL")
		.Fields("Forum_name") = Request.Form("forumName")
		.Fields("Forum_icon") = Request.Form("forumIcon")
		.Fields("Forum_description") = Request.Form("description")
		.Fields("Locked") = False
		.Fields("Hide") = True
		.Fields("Show_topics") = 0
		.Fields("Password") = null
		.Fields("Forum_code") = null

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
			.Fields("Post") = False
			.Fields("Priority_posts") = False
			.Fields("Reply_posts") = False
			.Fields("Edit_posts") = False
			.Fields("Delete_posts") = False
			.Fields("Poll_create") = False
			.Fields("Vote") = False
			.Fields("Display_post") = False
			.Fields("Moderate") = False
			.Fields("Calendar_event") = False
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
		strForumName = rsCommon("Forum_name")
		strForumURL = rsCommon("Forum_URL")
		strForumDescription = rsCommon("Forum_description")
		intForumID = CInt(rsCommon("Forum_ID"))
		strForumImage = rsCommon("Forum_icon")
	End If

	'Close the rs
	rsCommon.Close
End If


'If strForumURL is blank inistyalise it
If strForumURL = "" Then strForumURL = "http://"
	

'See if there is a main forum for this forum to be in or if there is a category to place a main forum within
'Read in the category name from the database
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Category.Cat_name, " & strDbTable & "Category.Cat_ID " & _
"FROM " & strDbTable & "Category " & _
"ORDER BY " & strDbTable & "Category.Cat_order ASC;"


'Query the database
rsCommon.Open strSQL, adoCon

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Forum External Link</title>
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


	//Check for a category
	if (document.frmNewForum.cat.value==""){
		alert("Please select the Category this Forum is to be in");
		return false;
	}
	
	//Check for a forum URL
	if (document.frmNewForum.forumURL.value=="" || document.frmNewForum.forumURL.value=="http://"){
		alert("Please enter a URL for the link");
		document.frmNewForum.forumURL.focus();
		return false;
	}
	
	//Check for a forum name
	if (document.frmNewForum.forumName.value==""){
		alert("Please enter a Name for the Link");
		document.frmNewForum.forumName.focus();
		return false;
	}

	//Check for a pforum description
	if (document.frmNewForum.description.value==""){
		alert("Please enter a Description for the Link");
		document.frmNewForum.description.focus();
		return false;
	}

	return true
}
</script>
<!-- #include file="includes/admin_header_inc.asp" -->
   <div align="center"><h1>Forum External Link</h1><br />
    <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
    <a href="admin_view_forums.asp<% = strQsSID1 %>">Return to the Forum Administration page</a><br />
    <br />
    From this page you can create a Forum External Link. A Forum External Link is a link that is displayed on your Forum Index in the list of forums, but is a link to an external web page.
    <br /><br />
    <%

'If there is no main forum or cat then display a message
If rsCommon.EOF Then
 	%>
    <table width="98%" border="0" cellspacing="0" cellpadding="1" height="135">
     <tr>
      <td align="center" class="text"><span class="lgText">
       You must first enter a Forum Category to place your new External Link in.</span><br />
       <br />
       <a href="admin_category_details.asp?mode=new<% = strQsSID2 %>">Enter a Forum Category</a>
       </td>
     </tr>
    </table>
    <%
Else
%>
   </div>
   <form action="admin_forum_link.asp?FID=<% = intForumID %><% = strQsSID2 %>" method="post" name="frmNewForum" id="frmNewForum" onsubmit="return CheckForm();">
    <table align="center" cellpadding="4" cellspacing="1" class="tableBorder">
     <tr>
      <td colspan="2" class="tableLedger">Select Forum Category</td>
     </tr>
     <tr>
      <td colspan="2" class="tableRow">Select the Forum Category from the drop down list below that you would like this External Link to be in.<br />
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
    <br />
    <table align="center" cellpadding="4" cellspacing="1" class="tableBorder">
     <tr>
      <td colspan="2" class="tableLedger">External Link</td>
     </tr>
     <tr>
      <td width="50%" class="tableRow">Link URL*:</td>
      <td width="50%" valign="top" class="tableRow"><input type="text" name="forumURL" maxlength="60" size="40" value="<% = strForumURL %>"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      </td>
     </tr>
     <tr>
      <td width="50%" class="tableRow">Link Name*:</td>
      <td width="50%" valign="top" class="tableRow"><input type="text" name="forumName" maxlength="60" size="40" value="<% = strForumName %>"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
      </td>
     </tr>
     <tr>
      <td class="tableRow">Link Description*:<br />
       <span class="smText">Give a brief description of the forum.</span></td>
      <td valign="top" class="tableRow"><input type="text" name="description" maxlength="190" size="70" value="<% = strForumDescription %>"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
      <tr>
      <td class="tableRow">Link Image Icon:<br />
       <span class="smText">Use this to add a link to a custom image icon displayed on the Forum Index page, in place of the standard icons displayed next to External Links.</span></td>
      <td valign="top" class="tableRow"><input type="text" id="forumIcon" name="forumIcon" maxlength="70" size="40" value="<% = strForumImage %>"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
    </table>
    <br />
    <br />
    <table width="100%" align="center" cellpadding="0" cellspacing="0">
     <tr>
      <td align="center" class="text"><span class="lgText">External Link Permissions</span><br />
       Select which Forum Groups are able to see this External Link in your Forums List.<br />
       <br />
       <table cellpadding="2" cellspacing="1" class="tableBorder">
        <tr>
         <td width="20%" align="left" class="tableLedger">Member Group</td>
         <td width="80%" align="left" class="tableLedger">View Forum Link</td>
         
        </tr>
        <tr class="tableSubLedger">
         <td align="left">Check All</td>
         <td align="left"><input type="checkbox" name="chkAllread" id="chkAllread" onclick="checkAll('read');" /></td>
         
         <td align="left">&nbsp;</td>
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
         <td align="left" class="tableRow"><input name="read<% = intUserGroupID %>" type="checkbox" id="read" value="true" checked /></td>
        
        </tr>
        <%

      		'Else display the values for this group
      		Else
%>
        <tr>
         <td align="left" class="tableRow"><% = rsCommon("Name") %>
          <% If CBool(rsCommon("Starting_group")) Then Response.Write("<br /><span class=""smText"">(New members group)</span>") %>
          <% If intUserGroupID = 2 Then Response.Write("<br /><span class=""smText"">(Un-registered users)</span>") %></td>
         <td align="left" class="tableRow"><input name="read<% = intUserGroupID %>" type="checkbox" id="read" value="true"<% If CBool(rsCommon2("View_Forum")) Then Response.Write(" checked") %> /></td>
         
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
     <input type="submit" name="Submit" value="Submit External Link"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
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
   <!-- #include file="includes/admin_footer_inc.asp" -->
