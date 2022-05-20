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
Dim intNewForumID	'Holds the new forum ID if the topic is to be moved
Dim strSubject		'Holds the topic subject
Dim intTopicPriority	'Holds the priority of the topic
Dim blnLockedStatus	'Holds the lock status of the topic
Dim strCatName		'Holds the name of the category
Dim intCatID		'Holds the ID number of the category
Dim strForumName	'Holds the name of the forum to jump to
Dim lngFID		'Holds the forum id to jump to
Dim lngPollID		'Holds the poll ID
Dim lngMovedNum		'Holds the moved number if topic has been moved
Dim blnMoved		'Set to true if moved icon is to be shown
Dim blnHidden		'set to true if post is hidden
Dim sarryForumSelect	'Holds the array with all the forums
Dim intSubForumID	'Holds if the forum is a sub forum
Dim intTempRecord	'Temporay record store
Dim blnHideForum	'Holds if the jump forum is hidden or not
Dim intCurrentRecord	'Holds the recordset array position
Dim intoldForumID	'Holds the forum ID of the forum the topic was moved from
Dim strTopicIcon	'Holds the topic icon
Dim intLoop		'Loop counter
Dim dtmEventDate	'Holds the event date
Dim intEventDay		'Holds the event day
Dim intEventMonth	'Holds the event month
Dim intEventYear	'Holds the event year
Dim intEventYearEnd	'Holds the year of Calendar event
Dim intEventMonthEnd	'Holds the month of Calendar event
Dim intEventDayEnd	'Holds the day of Calendar event
Dim dtmEventDateEnd	'Holds the Calendar event date
Dim strForumURL		'Holds if the forum is a external link


blnSslEnabledPage = True


'Read in the topic ID number
lngTopicID = LngC(Request("TID"))


'If the person is not an admin or a moderator then send them away
If lngTopicID = "" OR bannedIP() OR  blnActiveMember = False OR blnBanned Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If


'Initliase the SQL query to get the topic details from the database
strSQL = "SELECT " & strDbTable & "Topic.Forum_ID, " & strDbTable & "Topic.Subject, " & strDbTable & "Topic.Priority, " & strDbTable & "Topic.Locked, " & strDbTable & "Topic.Poll_ID, " & strDbTable & "Topic.Moved_ID, " & strDbTable & "Topic.Hide, " & strDbTable & "Topic.Icon, " & strDbTable & "Topic.Event_date, " & strDbTable & "Topic.Event_date_end " & _
"FROM " & strDbTable & "Topic" & strRowLock & " " & _
"WHERE " & strDbTable & "Topic.Topic_ID = " & lngTopicID & ";"

'Set the cursor	type property of the record set	to Forward only
rsCommon.CursorType = 0

'Set the Lock Type for the records so that the record set is only locked when it is updated
rsCommon.LockType = 3

'Query the database
rsCommon.Open strSQL, adoCon

'If there is a record returened read in the forum ID
If NOT rsCommon.EOF Then
	intForumID = CInt(rsCommon("Forum_ID"))
End If


'Call the moderator function and see if the user is a moderator
If blnAdmin = False Then blnModerator = isModerator(intForumID, intGroupID)



'If the user is not a moderator or admin then kick em
If blnAdmin = false AND  blnModerator = false Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If


'If this is a post back then  update the database
If (Request.Form("postBack")) AND (blnAdmin OR blnModerator) Then
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))

	strSubject = Trim(Mid(Request.Form("subject"), 1, 50))
	intTopicPriority = IntC(Request.Form("priority"))
	blnLockedStatus = BoolC(Request.Form("locked"))
	intNewForumID = IntC(Request.Form("forum"))
	blnMoved = BoolC(Request.Form("moveIco"))
	blnHidden = BoolC(Request.Form("hidePost"))
	strTopicIcon = Request.Form("icon")
	'Read in Calendar event date
	If Request.Form("eventDay") <> 0 AND Request.Form("eventMonth") <> 0 AND Request.Form("eventYear") <> 0 Then
		dtmEventDate = internationalDateTime(DateSerial(Request.Form("eventYear"), Request.Form("eventMonth"), Request.Form("eventDay")))
	End If
	'Read in event end date
	If Request.Form("eventDayEnd") <> 0 AND Request.Form("eventMonthEnd") <> 0 AND Request.Form("eventYearEnd") <> 0 Then
		dtmEventDateEnd = internationalDateTime(DateSerial(Request.Form("eventYearEnd"), Request.Form("eventMonthEnd"), Request.Form("eventDayEnd")))
		
		'If the end date is before the start date don't add it to the database
		If dtmEventDate => dtmEventDateEnd OR dtmEventDate = "" Then dtmEventDateEnd = null
	End If
	
	'Get rid of scripting tags in the subject
	strSubject = removeAllTags(strSubject)
	
	'If the topic icon is not selected don't fill the db with crap and leave field empty
	If strTopicIcon = strImagePath & "blank_smiley.gif" Then strTopicIcon = ""
	
	'Clean up user input
	strTopicIcon = formatInput(strTopicIcon)
	strTopicIcon = removeAllTags(strTopicIcon)
	
	
	
	'If logging is enabled then update log files
	If blnLoggingEnabled AND blnModeratorLogging Then
		
		If strSubject <> rsCommon("Subject") Then Call logAction(strLoggedInUsername, "Updated Topic Subject to '" & decodeString(strSubject) & "' - TopicID " & lngTopicID)
		If intNewForumID <> CInt(rsCommon("Forum_ID")) Then Call logAction(strLoggedInUsername, "Moved Topic '" & decodeString(strSubject) & "' to new forum - TopicID " & lngTopicID)
		If blnLockedStatus <> CBool(rsCommon("Locked")) Then 
			If blnLockedStatus Then 
				Call logAction(strLoggedInUsername, "Locked Topic '" & decodeString(strSubject) & "' - TopicID " & lngTopicID)
			Else
				Call logAction(strLoggedInUsername, "Unlocked Topic '" & decodeString(strSubject) & "' - TopicID " & lngTopicID)
			End If
		End If
		If intTopicPriority <> CInt(rsCommon("Priority")) Then Call logAction(strLoggedInUsername, "Changed Topic Priority '" & decodeString(strSubject) & "' - TopicID " & lngTopicID)
		If blnHidden <> CBool(rsCommon("Hide")) Then 
			If blnHidden Then 
				Call logAction(strLoggedInUsername, "Hide Topic '" & decodeString(strSubject) & "' - TopicID " & lngTopicID)
			Else
				Call logAction(strLoggedInUsername, "Approved Topic '" & decodeString(strSubject) & "' - TopicID " & lngTopicID)
			End If
		End If
	End If
	
	
	
	
	'Update the recordset
	With rsCommon
	
		'Update	the recorset
		If intNewForumID <> 0 Then .Fields("Forum_ID") = intNewForumID
		If blnMoved = false Then 
			.Fields("Moved_ID") = 0
		ElseIf intNewForumID <> 0 AND intNewForumID <> intForumID AND blnMoved Then 
			.Fields("Moved_ID") = intForumID
		End If
		.Fields("Subject") = strSubject
		.Fields("Priority") = intTopicPriority
		.Fields("Locked") = blnLockedStatus
		.Fields("Hide") = blnHidden
		.Fields("Icon") = strTopicIcon
		If blnCalendar Then .Fields("Event_date") = dtmEventDate
		If blnCalendar Then .Fields("Event_date_end") = dtmEventDateEnd
	
		'Update db
		.Update
	
		'Re-run the query as access needs time to catch up
		.ReQuery
			
		'Update topic and post count
		Call updateForumStats(intForumID)
		
		'Re-run the query (AGAIN!!) as crappy Access needs time to catch up
		.ReQuery
	
	End With
	
	'Update topic and post count
	Call updateForumStats(intForumID)

End If


'Read in the topic details
If NOT rsCommon.EOF Then
	
	'Read in the topic details
	intoldForumID = CInt(rsCommon("Forum_ID"))
	strSubject = rsCommon("Subject")
	intTopicPriority = CInt(rsCommon("Priority"))
	blnLockedStatus = CBool(rsCommon("Locked"))
	lngPollID = CLng(rsCommon("Poll_ID"))
	lngMovedNum = CLng(rsCommon("Moved_ID"))
	blnHidden = CBool(rsCommon("Hide"))
	strTopicIcon = rsCommon("Icon")
	dtmEventDate = rsCommon("Event_date")
	dtmEventDateEnd = rsCommon("Event_date_end")
	
	'Split the start date of event into the various parts
	If isDate(dtmEventDate) Then
		intEventYear = Year(dtmEventDate)
		intEventMonth = Month(dtmEventDate)
		intEventDay = Day(dtmEventDate)
	End If
	
	'Split the end date of event into the various parts
	If isDate(dtmEventDateEnd) Then
		intEventYearEnd = Year(dtmEventDateEnd)
		intEventMonthEnd = Month(dtmEventDateEnd)
		intEventDayEnd = Day(dtmEventDateEnd)
	End If

End If


'Close the rs
rsCommon.Close





'If this is a post back then  update the topic stats in the tblTopic table
If (Request.Form("postBack")) AND (blnAdmin OR blnModerator) Then
	
	'Update the stats for this topic in the tblTopic table
	'This isn't really neccisary, but could be useful for those people whos topic stats have gone astray
	Call updateTopicStats(lngTopicID)
End If





'DB hit to get forums with cats and permissions, for the forum select drop down

'Initlise the sql statement
strSQL = "" & _
"SELECT " & strDbTable & "Category.Cat_ID, " & strDbTable & "Category.Cat_name, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Hide, " & strDbTable & "Permissions.View_Forum, " & strDbTable & "Forum.Forum_URL " & _
"FROM " & strDbTable & "Category" & strDBNoLock & ", " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Permissions" & strDBNoLock & "  " & _
"WHERE " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID " & _
	"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Permissions.Forum_ID  " & _
	"AND (" & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intGroupID & ") " & _
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
<title>Topic Admin</title>

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
//Function to check form is filled in correctly before submitting
function CheckForm () {

	var errorMsg = "";
	var formArea = document.getElementById('frmTopicAdmin');

	//Check for a subject
	if (formArea.subject.value==""){
		errorMsg += "\n<% = strTxtErrorTopicSubject %>";
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
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="tableTopRow">
  <tr class="tableTopRow"> 
   <td colspan="2">
    <div style="float:left;"><h1><% = strTxtTopicAdmin %></h1></div>
    <div style="float:right;">
     <form name="frmDeleteTopic" id="frmDeleteTopic" method="post" action="delete_topic.asp" OnSubmit="return confirm('<% = strTxtDeleteTopicAlert %>')">
      <input type="hidden" name="XID" id="XID" value="<% = getSessionItem("KEY") %>" />
      <input type="hidden" name="TID" id="TID2" value="<% = lngTopicID %>" />
      <input type="submit" name="Submit" id="Submit2" value="<% = strTxtDeleteTopic %>"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /> 
     </form>
    </div>
   </td> 
  </tr>
 <form name="frmTopicAdmin" id="frmTopicAdmin" method="post" action="pop_up_topic_admin.asp<% = strQsSID1 %>" onSubmit="return CheckForm();">
  <tr class="tableRow" colspan="2">
    <td colspan="2">*<% = strTxtRequiredFields %></td>
  </tr>
  <tr class="tableRow">
   <td align="right" width="20%"><% = strTxtSubjectFolder %>*:</td>
   <td width="80%"><input type="text" name="subject" size="30" maxlength="50" value="<% = strSubject %>" /></td>
  </tr><%

'Display message icon drop down
If blnTopicIcon Then

	'Get the topic icon array
	%><!--#include file="includes/topic_icon_inc.asp" -->
  <tr class="tableRow">
   <td align="right" width="20%"><% = strTxtMessageIcon %>:</td>
   <td align="left" width="80%">
    <select name="icon" id="icon" onchange="(T_icon.src = icon.options[icon.selectedIndex].value)" >
     <option value="<% = strImagePath %>blank_smiley.gif"<% If strTopicIcon = "" Then Response.Write(" selected") %>><% = strTxtNoneSelected %></option><%

	'Loop through to display topic icons
	For intLoop = 1 TO Ubound(saryTopicIcon)

		Response.Write(vbCrLf & "     <option value=""" & saryTopicIcon(intLoop,2) & """")
		If strTopicIcon = saryTopicIcon(intLoop,2) Then Response.Write(" selected")
		Response.Write(">" & saryTopicIcon(intLoop,1) & "</option>")
	Next

	'If no topic Icon then get the default one
	If strTopicIcon = "" Then strTopicIcon = strImagePath & "blank_smiley.gif"
%>
    </select>
    &nbsp;&nbsp;<img src="<% = strTopicIcon %>" border="0" id="T_icon" alt"<% = strTxtMessageIcon %>" />
   </td>
  </tr><%

End If

%>
  <tr class="tableRow">
   <td align="right"><% = strTxtPinnedTopic %>:</td>
    <td>
    <select name="priority">
     <option value="0"<% If intTopicPriority = 0 Then Response.Write(" selected") %>><% = strTxtNormal %></option>
     <option value="1"<% If intTopicPriority = 1 Then Response.Write(" selected") %>><% = strTxtPinnedTopic %></option>
     <option value="2"<% If intTopicPriority = 2 Then Response.Write(" selected") %>><% = strTopThisForum %></option><%

'If this is the forum admin let them post a priority post to all forums
If blnAdmin Then 
         		%>
     <option value="3"<% If intTopicPriority = 3 Then Response.Write(" selected") %>><% = strTxtTopAllForums %></option><%

End If
        	%>
    </select></td>
   </tr><%




'Display Calendar event date input (if this is an event)
If blnCalendar AND (isDate(dtmEventDate) OR isDate(dtmEventDateEnd)) Then
%>
  <tr class="tableRow">
   <td align="right" valign="top"><% = strTxtCalendarEvent %>:</td>
   <td align="left">
    <% = strTxtStartDate %>:
    <br />
    &nbsp;&nbsp;&nbsp;&nbsp;<% = strTxtDay %>
    <select name="eventDay" id="eventDay">
     <option value="0"<% If intEventDay = 0 Then Response.Write(" selected") %>>----</option><%

	'Create lists day's for birthdays
	For intLoop = 1 to 31
		Response.Write(vbCrLf & "     <option value=""" & intLoop & """")
		If intEventDay = intLoop Then Response.Write(" selected")
		Response.Write(">" & intLoop & "</option>")
	Next

%>
    </select>
    <% = strTxtCMonth %>
    <select name="eventMonth" id="eventMonth">
     <option value="0"<% If intEventMonth = 0 Then Response.Write(" selected") %>>---</option><%

	'Create lists of days of the month for birthdays
	For intLoop = 1 to 12
		Response.Write(vbCrLf & "     <option value=""" & intLoop & """")
		If intEventMonth = intLoop Then Response.Write(" selected")
		Response.Write(">" & intLoop & "</option>")
	Next

%>
    </select>
    <% = strTxtCYear %>
    <select name="eventYear" id="eventYear">
     <option value="0"<% If intEventYear = 0 Then Response.Write(" selected") %>>-----</option><%

	'If this is an old event and the date is from a previous year, display that year
	If intEventYear <> 0 AND intEventYear < CInt(Year(Now())) Then Response.Write(VbCrLf & "     <option value=""" & intEventYear & """ selected>" & intEventYear & "</option>")

	'Create lists of years for birthdays
	For intLoop = CInt(Year(Now())) to CInt(Year(Now()))+1
		Response.Write(vbCrLf & "     <option value=""" & intLoop & """")
		If intEventYear = intLoop Then Response.Write(" selected")
		Response.Write(">" & intLoop & "</option>")
	Next

%>
    </select>
    <br />
    <% = strTxtEndDate %>: 
    <br />
    &nbsp;&nbsp;&nbsp;&nbsp;<% = strTxtDay %>
    <select name="eventDayEnd" id="eventDayEnd">
     <option value="0"<% If intEventDayEnd = 0 Then Response.Write(" selected") %>>----</option><%

		'Create lists day's for birthdays
		For intLoop = 1 to 31
			Response.Write(vbCrLf & "     <option value=""" & intLoop & """")
			If intEventDayEnd = intLoop Then Response.Write(" selected")
			Response.Write(">" & intLoop & "</option>")
		Next

%>
    </select>
    <% = strTxtCMonth %>
    <select name="eventMonthEnd" id="eventMonthEnd">
     <option value="0"<% If intEventMonthEnd = 0 Then Response.Write(" selected") %>>---</option><%

		'Create lists of days of the month for birthdays
		For intLoop = 1 to 12
			Response.Write(vbCrLf & "     <option value=""" & intLoop & """")
			If intEventMonthEnd = intLoop Then Response.Write(" selected")
			Response.Write(">" & intLoop & "</option>")
		Next

%>
    </select>
    <% = strTxtCYear %>
    <select name="eventYearEnd" id="eventYearEnd">
     <option value="0"<% If intEventYear = 0 Then Response.Write(" selected") %>>-----</option><%

		'If this is an old event and the date is from a previous year, display that year
		If intEventYearEnd <> 0 AND intEventYearEnd < CInt(Year(Now())) Then Response.Write(VbCrLf & "     <option value=""" & intEventYearEnd & """ selected>" & intEventYearEnd & "</option>")

		'Create lists of years for birthdays
		For intLoop = CInt(Year(Now())) to CInt(Year(Now()))+1
			Response.Write(vbCrLf & "     <option value=""" & intLoop & """")
			If intEventYearEnd = intLoop Then Response.Write(" selected")
			Response.Write(">" & intLoop & "</option>")
		Next

%>
    </select>
   </td>
  </tr><%

End If




	
	
%>
   <tr class="tableRow">
    <td align="right"><% = strTxtLockedTopic %>:</td>
    <td><input type="checkbox" name="locked" value="true" <% If blnLockedStatus = True Then Response.Write(" checked") %> /></td>
   </tr>
   <tr class="tableRow">
    <td align="right" valign="top"><% = strTxtMoveTopic %>:</td>
    <td><select name="forum" id="forum"><%


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
		Response.Write vbCrLf & "     <optgroup label=""&nbsp;&nbsp;" & strCatName & """>"
		
		
		
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
				Response.Write (vbCrLf & "      <option value=""" & intForumID & """")
				If intoldForumID = intForumID Then Response.Write(" selected")
				Response.Write(">&nbsp;" & strForumName & "</option>")	
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
						Response.Write (vbCrLf & "      <option value=""" & intSubForumID & """")
						If intoldForumID = intSubForumID Then Response.Write(" selected")
						Response.Write (">&nbsp&nbsp;-&nbsp;" & strForumName & "</option>")	
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
		
		
		Response.Write(vbCrLf & "     </optgroup>")
	Loop
End If
%>
     </select><br />
     <input type="checkbox" name="moveIco" value="true" checked /><% = strTxtShowMovedIconInLastForum %></td>
   </tr><%

'If there is a poll then let the admin moderator edit or delete the poll
If lngPollID <> 0 Then

	%>
   <tr class="tableRow">
    <td align="right"><% = strTxtPoll %></td>
    <td><a href="delete_poll.asp?TID=<% = lngTopicID  & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 %>" onclick="return confirm('<% = strTxtAreYouSureYouWantToDeleteThisPoll %>');"><% = strTxtDeletePoll %></a></td>
   </tr><%
End If

%>
   <tr class="tableRow">
    <td align="right" valign="top"><% = strTxtHideTopic %>:</td>
    <td><input type="checkbox" name="hidePost" value="true" <% If blnHidden = True Then Response.Write(" checked") %> /> <span class="smText"><% = strTxtIfYouAreShowingTopic %></span></td>
   </tr>
   <tr align="center" class="tableRow">
    <td valign="top" colspan="2" />
     <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
     <input type="hidden" name="TID" id="TID" value="<% = lngTopicID %>" />
     <input type="hidden" name="postBack" id="postBack" value="true" />
    </td>
   </tr>
   <tr class="tableBottomRow">
    <td align="right" colspan="2">
     <input type="submit" name="Submit" id="Submit" value="     <% = strTxtOK %>     " <% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />&nbsp;<input type="button" name="cancel" value=" <% = strTxtClose %> " onclick="window.close()"></td>
   </tr>
  </form>
</table>
</body>
</html>