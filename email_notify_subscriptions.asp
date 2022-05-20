<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
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





'Set the buffer to true
Response.Buffer = True

'Declare variables
Dim intForumColourNumber	'Holds the number to calculate the table row colour
Dim rsForumSelect		'Recordset for teh forum selection	
Dim strCatName			'Holds the category name
Dim intCatID			'Holds the cat ID
Dim strForumName		'Holds the forum  name
Dim lngEmailUserID		'Holds the user ID to look at email notification for
Dim strMode			'Holds the mode of the page
Dim dtmLastEntryDate		'Holds the last entry date
Dim intCurrentRecord		'Holds the recordset array position
Dim sarrySubscribedForums	'Holds the subscribed forums
Dim sarrySubscribedTopics	'Holds the subscribed topics
Dim sarryForumSelect		'Holds the array with all the forums
Dim intSubForumID		'Holds if the forum is a sub forum
Dim intTempRecord		'Temporay record store
Dim blnHideForum		'Holds if the jump forum is hidden or not
Dim strXID 			'Holds the session key

'Initialise variable
blnSslEnabledPage = True
intForumColourNumber = 0
intCurrentRecord = 0


'If emial notify is not on then send them away
If blnEmail = False Then 
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
End If


'If the user is not allowed then send them away
If intGroupID = 2 OR blnActiveMember = False Then 
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If



'Read in the mode of the page
strMode = Trim(Mid(Request.QueryString("M"), 1, 1))


'If this is not an admin but in admin mode then see if the user is a moderator
If blnAdmin = False AND strMode = "A"  AND blnModeratorProfileEdit Then
	
	'Initalise the strSQL variable with an SQL statement to query the database
        strSQL = "SELECT " & strDbTable & "Permissions.Moderate " & _
        "FROM " & strDbTable & "Permissions" & strDBNoLock & " " & _
        "WHERE (" & strDbTable & "Permissions.Group_ID = " & intGroupID & " OR " & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & ") AND  " & strDbTable & "Permissions.Moderate=" & strDBTrue & ";"
               

        'Query the database
         rsCommon.Open strSQL, adoCon

        'If a record is returned then the user is a moderator in one of the forums
        If NOT rsCommon.EOF Then blnModerator = True

        'Clean up
        rsCommon.Close
End If


'Get the user ID of the email notifications to look at
If (blnAdmin OR (blnModerator AND LngC(Request.QueryString("PF")) > 2)) AND strMode = "A" AND LngC(Request.QueryString("PF")) <> 2 Then
	
	lngEmailUserID = LngC(Request.QueryString("PF"))

'Get the logged in ID number
Else
	lngEmailUserID = lngLoggedInUserID
End If






'DB hit to get subscribed forums

'Initlise the sql statement
strSQL = "SELECT " & strDbTable & "Forum.Forum_name, " & strDbTable & "EmailNotify.Forum_ID, " & strDbTable & "EmailNotify.Watch_ID " & _
"FROM " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "EmailNotify" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Forum.Forum_ID=" & strDbTable & "EmailNotify.Forum_ID AND " & strDbTable & "EmailNotify.Author_ID=" & lngEmailUserID & " " & _
"ORDER BY " & strDbTable & "Forum.Forum_Order;"

'Query the database
rsCommon.Open strSQL, adoCon

'Place the subscribed forums into an array
If NOT rsCommon.EOF Then
	
	'Read in the row from the db using getrows for better performance
	sarrySubscribedForums = rsCommon.GetRows()
End If

'Clean up
rsCommon.Close




'DB hit to get subscribed topics

'Initlise the sql statement
strSQL = "SELECT " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Last_Thread_ID, " & strDbTable & "EmailNotify.Topic_ID, " & strDbTable & "EmailNotify.Watch_ID, " & strDbTable & "Thread.Message_date " & _
"FROM " & strDbTable & "Topic" & strDBNoLock & ", " & strDbTable & "EmailNotify" & strDBNoLock & ", " & strDbTable & "Thread" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Topic.Last_Thread_ID=" & strDbTable & "Thread.Thread_ID AND " & strDbTable & "Topic.Topic_ID=" & strDbTable & "EmailNotify.Topic_ID AND " & strDbTable & "EmailNotify.Author_ID=" & lngEmailUserID & " " & _
"ORDER BY " & strDbTable & "Thread.Message_date DESC;"

'Query the database
rsCommon.Open strSQL, adoCon		

'Place the subscribed topics into an array
If NOT rsCommon.EOF Then
	
	'Read in the row from the db using getrows for better performance
	sarrySubscribedTopics = rsCommon.GetRows()
End If

'Clean up
rsCommon.Close




'DB hit to get forums with cats and permissions, for the forum select drop down

'Initlise the sql statement
strSQL = "" & _
"SELECT " & strDbTable & "Category.Cat_ID, " & strDbTable & "Category.Cat_name, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Hide, " & strDbTable & "Permissions.View_Forum " & _
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


'get session key
strXID = getSessionItem("KEY")


Call closeDatabase()


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtEmailNotificationSubscriptions


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtEmailNotificationSubscriptions %></title>

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

<script  language="JavaScript">
//Funtion to check or uncheck all forum delete boxes
function checkAllForum(){

	var formArea = document.getElementById('frmForumDel');

 	if (formArea.chkDelete.length > 0) {
  		for (i=0; i < formArea.chkDelete.length; i++){
   			formArea.chkDelete[i].checked = formArea.chkAll.checked;
  		}
 	}
 	else {
  		formArea.chkDelete.checked = formArea.chkAll.checked;
 	}
}

//Funtion to check or uncheck all topic delete boxes
function checkAllTopic(){

	var formArea = document.getElementById('frmTopicDel');

 	if (formArea.chkDelete.length > 0) {
  		for (i=0; i < formArea.chkDelete.length; i++){
   				formArea.chkDelete[i].checked = formArea.chkAll.checked;
  		}
 	}
 	else {
  		formArea.chkDelete.checked = formArea.chkAll.checked;
 	}
}

//Function to check form is filled in correctly before submitting
function CheckForm () {
	
	var errorMsg = "";
	var formArea = document.getElementById('frmAddForum');
	
	//Check for a forum
	if (formArea.FID.value==""){
		errorMsg += "\n<% = strTxtSelectForumErrorMsg %>";
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
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtEmailNotificationSubscriptions %></h1></td>
</tr>
</table>  
<br />
<table class="basicTable" cellspacing="0" cellpadding="0" align="center"> 
 <tr> 
  <td class="tabTable">
   <a href="member_control_panel.asp<% If strMode = "A" Then Response.Write("?PF=" & lngEmailUserID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtControlPanel %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>member_control_panel.<% = strForumImageType %>" border="0" alt="<% = strTxtControlPanel %>" /> <% = strTxtControlPanel %></a>
   <a href="register.asp<% If strMode = "A" Then Response.Write("?PF=" & lngEmailUserID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtProfile2 %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>profile.<% = strForumImageType %>" border="0" alt="<% = strTxtProfile2 %>" /> <% = strTxtProfile2 %></a><%
 
If blnEmail Then

%>
   <a href="email_notify_subscriptions.asp<% If strMode = "A" Then Response.Write("?PF=" & lngEmailUserID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtSubscriptions %>" class="tabButtonActive">&nbsp;<img src="<% = strImagePath %>subscriptions.<% = strForumImageType %>" border="0" alt="<% = strTxtSubscriptions %>" /> <% = strTxtSubscriptions %></a><%
End If


'Only disply other links if not in admin mode
If strMode <> "A" AND blnActiveMember AND blnPrivateMessages Then

%>
   <a href="pm_buddy_list.asp<% = strQsSID1 %>" title="<% = strTxtBuddyList %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>buddy_list.<% = strForumImageType %>" border="0" alt="<% = strTxtBuddyList %>" /> <% = strTxtBuddyList %></a><%

End If


'If the user is user is using a banned IP redirect to an error page
If blnAttachments OR blnImageUpload Then

%>
   <a href="file_manager.asp<% If strMode = "A" Then Response.Write("?PF=" & lngEmailUserID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtFileManager %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>file_manager.<% = strForumImageType %>" border="0" alt="<% = strTxtFileManager %>" /> <% = strTxtFileManager %></a><%

End If



%>
  </td>
 </tr>
</table>
<br />
<form name="frmForumDel" id="frmForumDel" method="post" action="email_notify_remove.asp<% If strMode = "A" Then Response.Write("?PF=" & lngEmailUserID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" OnSubmit="return confirm('<% = strTxtAreYouWantToUnsubscribe & " " & strTxtForums %>')">
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td width="97%" align="left"><% = strTxtForums & " " & strTxtThatYouHaveSubscribedTo %></td>
  <td width="3%" align="center"><input type="checkbox" name="chkAll" onclick="checkAllForum();"></td>
 </tr><%
	
    
'Check there are email subscriptions to show
If isArray(sarrySubscribedForums) = false Then

	'If there are no email notifies to display
	Response.Write(vbCrLf & " <tr class=""tableRow""><td colspan=""2"" align=""center""><br />" &  strTxtYouHaveNoSubToEmailNotify & "<input type=""hidden"" name=""chkDelete"" value=""-1""><br /><br /></tr></td>")

'Else there the are email subs so show em
Else 	
	'Loop round to read in all the email notifys
	Do WHILE intCurrentRecord =< UBound(sarrySubscribedForums, 2)
	
		'SQL Query Array Look Up table
		'0 = Forum_name
		'1 = Forum_ID
		'2 = Watch_ID
	
		'Get the row number
		intForumColourNumber = intForumColourNumber + 1
	
	%>
 <tr class="<% If (intForumColourNumber MOD 2 = 0 ) Then Response.Write("evenTableRow") Else Response.Write("oddTableRow") %>"> 
  <td><a href="forum_topics.asp?FID=<% = sarrySubscribedForums(1,intCurrentRecord) %><% = strQsSID2 %>"><% = sarrySubscribedForums(0,intCurrentRecord) %></a></td>
  <td align="center"><input type="checkbox" name="chkDelete" value="<% = sarrySubscribedForums(2,intCurrentRecord) %>"></td>
 </tr><%
		
		'Move to the next record
		intCurrentRecord = intCurrentRecord + 1
	Loop
End If

'reset
intCurrentRecord = 0
%>
</table>
<table class="basicTable" cellspacing="0" cellpadding="2" align="center">
 <tr align="right"><td><input type="hidden" name="formID" id="formID1" value="<% = strXID %>" /><input type="submit" name="Submit" id="Submit1" value="<% = strTxtUnsusbribe %>" /></td></tr>
</table>
</form>
<form name="frmTopicDel" id="frmTopicDel" method="post" action="email_notify_remove.asp<% If strMode = "A" Then Response.Write("?PF=" & lngEmailUserID & "&M=A") %><% = strQsSID2 %>" OnSubmit="return confirm('<% = strTxtAreYouWantToUnsubscribe & " " & strTxtTopics %>')">
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td width="67%" align="left"><% = strTxtTopics & " " & strTxtThatYouHaveSubscribedTo %></td>
  <td width="30%"><% = strTxtLastPost %></td>
  <td width="3%" align="center"><input type="checkbox" name="chkAll" onclick="checkAllTopic();"></td>
 </tr><%
	
    
'Check there are email subscriptions to show
If isArray(sarrySubscribedTopics) = false Then

	'If there are no pm messages to display then display the appropriate error message
	Response.Write(vbCrLf & " <tr class=""tableRow""><td colspan=""3"" align=""center""><br />" & strTxtYouHaveNoSubToEmailNotify & "<input type=""hidden"" name=""chkDelete"" value=""-1""><br /><br /></tr></td>")

'Else there the are email subs so show em
Else 	
	'Loop round to read in all the email notifys
	Do WHILE intCurrentRecord =< UBound(sarrySubscribedTopics, 2)
	
		'SQL Query Array Look Up table
		'0 = Subject
		'1 = Last_Thread_ID
		'2 = Topic_ID
		'3 = Watch_ID
		'4 = Message_date
	
		'Get the date of the last entry
		dtmLastEntryDate = CDate(sarrySubscribedTopics(4,intCurrentRecord))
	
		'Get the row number
		intForumColourNumber = intForumColourNumber + 1
	%>
 <tr class="<% If (intForumColourNumber MOD 2 = 0 ) Then Response.Write("evenTableRow") Else Response.Write("oddTableRow") %>"> 
  <td><a href="forum_posts.asp?TID=<% = sarrySubscribedTopics(2,intCurrentRecord) %><% = strQsSID2 %>"><% = formatInput(sarrySubscribedTopics(0,intCurrentRecord)) %></a></td>
  <td nowrap><% Response.Write(DateFormat(dtmLastEntryDate) & "&nbsp;" & strTxtAt & "&nbsp;" & TimeFormat(dtmLastEntryDate)) %></td>
  <td align="center"><input type="checkbox" name="chkDelete" value="<% = sarrySubscribedTopics(3,intCurrentRecord) %>"></td>
 </tr><%
		
		'Move to the next record
		intCurrentRecord = intCurrentRecord + 1
	Loop
End If

'reset
intCurrentRecord = 0
%>
</table>
<table class="basicTable" cellspacing="0" cellpadding="2" align="center">
 <tr align="right"><td><input type="hidden" name="formID" id="formID2" value="<% = strXID %>" /><input type="submit" name="Submit" id="Submit2" value="<% = strTxtUnsusbribe %>" /></td></tr>
</table>
</form><%

'If this is not in admin mode then see if the user wants email notification of a forum
If strMode <> "A" Then 
%>
<form method="post" name="frmAddForum" id="frmAddForum" action="email_notify.asp?M=SP&XID=<% = strXID & strQsSID2 %>" onSubmit="return CheckForm();">
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td><% = strTxtSubscribeToForum %></td>
 </tr>
 <tr class="tableRow"> 
  <td> <% = strTxtSelectForumToSubscribeTo %> <%
            
Response.Write(vbCrLf & "   <select name=""FID"">")
Response.Write(vbCrLf & "    <option value="""" selected>-- " & strTxtSelectForum & " --</option>")

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
		Response.Write vbCrLf & "    <optgroup label=""&nbsp;&nbsp;" & strCatName & """>"
		
		
		
		'Loop round to display all the forums for this category
		Do While intCurrentRecord <= Ubound(sarryForumSelect,2)
		
			'Read in the forum details from the recordset
			intForumID = CInt(sarryForumSelect(2,intCurrentRecord))
			intSubForumID = CInt(sarryForumSelect(3,intCurrentRecord))
			strForumName = sarryForumSelect(4,intCurrentRecord)
			blnHideForum = CBool(sarryForumSelect(5,intCurrentRecord))
			blnRead = CBool(sarryForumSelect(6,intCurrentRecord))
			
			
			'If this forum is to be hidden but the user is allowed access to it set the hidden boolen back to false
			If blnHideForum = True AND blnRead = True Then blnHideForum = False

			'If the forum is not a hidden forum to this user, display it
			If blnHideForum = False AND intSubForumID = 0 Then
				'Display a link in the link list to the forum
				Response.Write (vbCrLf & "     <option value=""" & intForumID & """>&nbsp;" & strForumName & "</option>")	
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
						Response.Write (vbCrLf & "     <option value=""" & intSubForumID & """>&nbsp&nbsp;-&nbsp;" & strForumName & "</option>")	
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
		
		
		Response.Write(vbCrLf & "    </optgroup>")
	Loop
End If

Response.Write(vbCrLf & "    </select>")

%>
   <input type="submit" name="Submit" id="subscribe" value="<% = strTxtSubscribe %>" />
  </td>
 </tr>
</table>
</form><%
End If
%>
<div align="center">
<br /><%

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	
	If blnTextLinks = True Then 
		Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
	Else
  		Response.Write("<a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank"" ><img src=""webwizforums_image.asp"" border=""0"" title=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ alt=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ /></a>")
	End If
	
	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

'Display the process time
If blnShowProcessTime Then Response.Write "<span class=""smText""><br /><br />" & strTxtThisPageWasGeneratedIn & " " & FormatNumber(Timer() - dblStartTime, 3) & " " & strTxtSeconds & "</span>"


Response.Write("</div>")

%>
<!-- #include file="includes/footer.asp" -->