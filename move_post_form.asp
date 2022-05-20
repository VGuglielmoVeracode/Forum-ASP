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
Dim rsSelectForum		'Holds the recordset for the forum
Dim strCatName			'Holds the name of the category
Dim intCatID			'Holds the ID number of the category
Dim strForumName		'Holds the name of the forum to jump to
Dim lngPostID			'Holds the post ID
Dim sarryForumSelect		'Holds the array with all the forums
Dim intSubForumID		'Holds if the forum is a sub forum
Dim intTempRecord		'Temporay record store
Dim blnHideForum		'Holds if the jump forum is hidden or not
Dim intCurrentRecord		'Holds the recordset array position
Dim strForumURL		'Holds if the forum is a external link


'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)
End If



'Read in the post ID
lngPostID = LngC(Request.QueryString("PID"))



'Query the datbase to get the forum ID for this post
strSQL = "SELECT " & strDbTable & "Topic.Forum_ID " & _
"FROM " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Topic.Topic_ID = " & strDbTable & "Thread.Topic_ID AND " & strDbTable & "Thread.Thread_ID = " & lngPostID & ";"

'Query the database
rsCommon.Open strSQL, adoCon


'If there is a record returened read in the forum ID
If NOT rsCommon.EOF Then
	intForumID = CInt(rsCommon("Forum_ID"))
End If

'Clean up
rsCommon.Close


'Call the moderator function and see if the user is a moderator
If blnAdmin = False Then blnModerator = isModerator(intForumID, intGroupID)


'If the user is not a moderator or admin then keck em
If blnAdmin = false AND  blnModerator = false Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If



'DB hit to get forums with cats and permissions, for the forum select drop down

'Initlise the sql statement
strSQL = "" & _
"SELECT " & strDbTable & "Category.Cat_ID, " & strDbTable & "Category.Cat_name, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Hide, " & strDbTable & "Permissions.View_Forum, " & strDbTable & "Forum.Forum_URL " & _
"FROM " & strDbTable & "Category" & strDBNoLock & ", " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Permissions" & strDBNoLock & "  " & _
"WHERE " & strDbTable & "Category.Cat_ID=" & strDbTable & "Forum.Cat_ID " & _
	"AND " & strDbTable & "Forum.Forum_ID=" & strDbTable & "Permissions.Forum_ID  " & _
	"AND (" & strDbTable & "Permissions.Author_ID=" & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID=" & intGroupID & ") " & _
"ORDER BY " & strDbTable & "Category.Cat_order, " & strDbTable & "Forum.Forum_Order;"

'Query the database
rsCommon.Open strSQL, adoCon		

'Place the subscribed topics into an array
If NOT rsCommon.EOF Then
	
	'Read in the row from the db using getrows for better performance
	sarryForumSelect = rsCommon.GetRows()
End If

'Clean up
rsCommon.Close
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Web Wiz Fourms Move Post</title>

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
<script language="JavaScript">

//Function to check form is filled in correctly before submitting
function CheckForm () {

	//Check for a forum to move Post to
	if (document.getElementById('frmMovePost').forum.value==""){

		msg = "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine1 %>\n";
		msg += "<% = strTxtErrorDisplayLine2 %>\n";
		msg += "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine3 %>\n";

		alert(msg + "\n<% = strTxtMovePostErrorMsg %>\n\n");
		return false;
	}

	return true
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" OnLoad="self.focus();">
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="tableTopRow">
 <form method="post" name="frmMovePost" id="frmMovePost" action="move_post_form_to.asp<% = strQsSID1 %>" onSubmit="return CheckForm();">
  <tr class="tableTopRow">
   <td colspan="2"><h1><% = strTxtMovePost %></h1></td>
    </tr>
    <tr class="tableRow" height="195">
      <td colspan="2">
        <br />
        <% = strTxtSelectForumClickNext %>
        <br />
        <br />
        <% = strTxtSelectTheForumYouWouldLikePostIn %>
        <br />
        <select name="forum" id="forum"><%


'Show forum select
If isArray(sarryForumSelect) Then
	
	'Loop round to show all the categories and forums
	Do While intCurrentRecord <= Ubound(sarryForumSelect,2)
		
		'Loop through the array looking for forums that are to be shown
		'if a forum is found to be displayed then show the category and the forum, if not the category is not displayed as there are no forums the user can access
		Do While intCurrentRecord <= Ubound(sarryForumSelect,2)
		
			'Read in details
			blnHideForum = CBool(sarryForumSelect(5,intCurrentRecord))
			blnRead = CBool(sarryForumSelect(6,intCurrentRecord))
					
			'If this forum is to be shown then leave the loop and display the cat and the forums
			If blnHideForum = False OR blnRead = True Then Exit Do
			
			'Move to next record
			intCurrentRecord = intCurrentRecord + 1
		Loop
				
		'If we have run out of records jump out of loop
		If intCurrentRecord > Ubound(sarryForumSelect,2) Then Exit Do
	 
		
		
		'Read in the deatils for the category
		intCatID = CInt(sarryForumSelect(0,intCurrentRecord))
		strCatName = sarryForumSelect(1,intCurrentRecord)		
		
		
		'Display a link in the link list to the forum
		Response.Write vbCrLf & "		<optgroup label=""&nbsp;&nbsp;" & strCatName & """>"
		
		
		
		'Loop round to display all the forums for this category
		Do While intCurrentRecord <= Ubound(sarryForumSelect,2)
		
			'Read in the forum details from the recordset
			intForumID = CInt(sarryForumSelect(2,intCurrentRecord))
			intSubForumID = CInt(sarryForumSelect(3,intCurrentRecord))
			strForumName = sarryForumSelect(4,intCurrentRecord)
			blnHideForum = CBool(sarryForumSelect(5,intCurrentRecord))
			blnRead = CBool(sarryForumSelect(6,intCurrentRecord))
			strForumURL = sarryForumSelect(7,intCurrentRecord)
				
			If strForumURL = "http://" OR isNull(strForumURL) Then strForumURL = ""
			
			
			'If this forum is to be hidden but the user is allowed access to it set the hidden boolen back to false
			If blnHideForum = True AND blnRead = True Then blnHideForum = False

			'If the forum is not a hidden forum to this user, display it
			If blnHideForum = False AND intSubForumID = 0 AND strForumURL = "" Then
				'Display a link in the link list to the forum
				Response.Write (vbCrLf & "		<option value=""" & intForumID & """>&nbsp;" & strForumName & "</option>")	
			End If
			
			
			
			'See if this forum has any sub forums
			'Initilise variables
			intTempRecord = 0
					
			'Loop round to read in any sub forums in the stored array recordset
			Do While intTempRecord <= Ubound(sarryForumSelect,2)
			
				'Becuase the member may have an individual permission entry in the permissions table for this forum, 
				'it maybe listed twice in the array, so we need to make sure we don't display the same forum twice
				If intSubForumID = CInt(sarryForumSelect(2,intTempRecord)) Then intTempRecord = intTempRecord + 1
				
				'If there are no records left exit loop
				If intTempRecord > Ubound(sarryForumSelect,2) Then Exit Do
				
				'If this is a subforum of the main forum then get the details
				If CInt(sarryForumSelect(3,intTempRecord)) = intForumID Then
				
					'Read in the forum details from the recordset
					intSubForumID = CInt(sarryForumSelect(2,intTempRecord))
					strForumName = sarryForumSelect(4,intTempRecord)
					blnHideForum = CBool(sarryForumSelect(5,intTempRecord))
					blnRead = CBool(sarryForumSelect(6,intTempRecord))
					
					
					'If this forum is to be hidden but the user is allowed access to it set the hidden boolen back to false
					If blnHideForum = True AND blnRead = True Then blnHideForum = False
		
					'If the forum is not a hidden forum to this user, display it
					If blnHideForum = False Then
						'Display a link in the link list to the forum
						Response.Write (vbCrLf & "		<option value=""" & intSubForumID & """>&nbsp&nbsp;-&nbsp;" & strForumName & "</option>")	
					End If
				End If
				
				'Move to next record 
				intTempRecord = intTempRecord + 1
				
			Loop
			
			
					
			'Move to the next record in the array
			intCurrentRecord = intCurrentRecord + 1
			
			
			'If there are more records in the array to display then run some test to see what record to display next and where				
			If intCurrentRecord <= Ubound(sarryForumSelect,2) Then

				'Becuase the member may have an individual permission entry in the permissions table for this forum, 
				'it maybe listed twice in the array, so we need to make sure we don't display the same forum twice
				If intForumID = CInt(sarryForumSelect(2,intCurrentRecord)) Then intCurrentRecord = intCurrentRecord + 1
				
				'If there are no records left exit loop
				If intCurrentRecord > Ubound(sarryForumSelect,2) Then Exit Do
				
				'See if the next forum is in a new category, if so jump out of this loop to display the next category
				If intCatID <> CInt(sarryForumSelect(0,intCurrentRecord)) Then Exit Do
			End If
		Loop
		
		
		Response.Write(vbCrLf & "		</optgroup>")
	Loop
End If
%>
          </select>
          <br />
          <br />
          <br />
          <br />
          <br />
          <br />
          <input type="hidden" id="PID" name="PID" value="<% = lngPostID %>" />
       </td>
    </tr>
    <tr class="tableBottomRow">
      <td width="38%" valign="top"><%

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode Then
	Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******      
      
      %></td>
      <td width="24%" align="right"><input type="submit" name="Submit" id="Submit" value="<% = strTxtNext %> &gt;&gt;" />&nbsp;<input type="button" name="cancel" id="cancel" value=" <% = strTxtCancel %> " onclick="window.close()" />        
       <input type="hidden" name="postBack" id="postBack" value="true" /></td>
    </tr>
  </form>
</table>
</body>
</html>