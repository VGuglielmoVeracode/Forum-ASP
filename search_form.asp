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




'Set the response buffer to true as we maybe redirecting
Response.Buffer = True


'Dimension variables
Dim strReturnPage		'Holds the page to return to
Dim strForumName 		'Holds the forum name
Dim intReadPermission		'holds the forums read permisisons
Dim strSearchKeywords		'Holds the keywords to search for
Dim strSearchMode		'Holds the search mode
Dim strSearchUser		'Holds the user to search for
Dim intCurrentRecord		'Holds the recordset array position
Dim sarrySubscribedForums	'Holds the subscribed forums
Dim sarrySubscribedTopics	'Holds the subscribed topics
Dim sarryForumSelect		'Holds the array with all the forums
Dim intSubForumID		'Holds if the forum is a sub forum
Dim intTempRecord		'Temporay record store
Dim blnHideForum		'Holds if the jump forum is hidden or not
Dim strCatName			'Holds the category name
Dim intCatID			'Holds the cat ID
Dim intForumID2			'Holds the read in forum id	
Dim lngTopicID			'Holds post id for searching posts
Dim strForumURL 		'Holds the forum URL if a link	





intCurrentRecord = 0

'Read in values passed to this form
blnSslEnabledPage = True
intForumID2 = IntC(Request("FID"))
strSearchKeywords = Trim(Mid(Request("KW"), 1, 35))
strSearchUser = Trim(Mid(Request.QueryString("USR"), 1, 25))
strSearchMode = Trim(Mid(Request.QueryString("SM"), 1, 3))
lngTopicID = LngC(Request("TID"))




'First see if the user is a in a moderator group for any forum
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Permissions.Moderate " & _
"FROM " & strDbTable & "Permissions " & _
"WHERE (" & strDbTable & "Permissions.Group_ID = " & intGroupID & " OR " & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & ") AND  " & strDbTable & "Permissions.Moderate=" & strDBTrue & ";"
	

'Query the database
rsCommon.Open strSQL, adoCon

'If a record is returned then the user is a moderator in one of the forums
If NOT rsCommon.EOF Then blnModerator = True

'Clean up
rsCommon.Close




'DB hit to get forums with cats and permissions, for the forum select drop down
If lngTopicID = 0 Then
	'Initlise the sql statement
	strSQL = "" & _
	"SELECT " & strDbTable & "Category.Cat_ID, " & strDbTable & "Category.Cat_name, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Hide, " & strDbTable & "Permissions.View_Forum, " & strDbTable & "Forum.Forum_URL " & _
	"FROM " & strDbTable & "Category" & strDBNoLock & ", " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Permissions" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID " & _
		"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Permissions.Forum_ID " & _
		"AND (" & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intGroupID & ") " & _
	"ORDER BY " & strDbTable & "Category.Cat_order, " & strDbTable & "Forum.Forum_Order, " & strDbTable & "Permissions.Author_ID DESC;"
	
	'Query the database
	rsCommon.Open strSQL, adoCon		
	
	'Place the subscribed topics into an array
	If NOT rsCommon.EOF Then
		
		'Read in the row from the db using getrows for better performance
		sarryForumSelect = rsCommon.GetRows()
	End If
	
	'Clean up
	rsCommon.Close

End If




'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers("", strTxtSearchingForums, "search_form.asp", 0)
End If


'Set bread crumb trail
If lngTopicID <> 0 Then
	strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtTopic & " " & strTxtSearch
Else
	strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtSearchTheForum
End If

Call closeDatabase()

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% If lngTopicID <> 0 Then Response.Write(strTxtTopic & " " & strTxtSearch) Else Response.Write(strTxtSearchTheForum) %></title>
<meta name="generator" content="Web Wiz Forums" />
<meta name="description" content="<% = strBoardMetaDescription %>" />
<meta name="keywords" content="<% = strBoardMetaKeywords %>" />

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

<script  language="JavaScript">
//Function to check form is filled in correctly before submitting
function CheckForm () {

	var formArea = document.getElementById('frmSearch');

	//Check for a somthing to search for
	if ((formArea.KW.value=="") && (formArea.USR.value=="")){

		msg = "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine1 %>\n";
		msg += "<% = strTxtErrorDisplayLine2 %>\n";
		msg += "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine3 %>\n";

		alert(msg + "\n<% = strTxtSearchFormError %>\n\n");
		formArea.KW.focus();
		return false;
	}	
	
	//Disable submit button
	//document.getElementById('Submit').disabled=true;

	//Show progress bar
	var progressWin = document.getElementById('progressBar');
	var progressArea = document.getElementById('progressFormArea');
	progressWin.style.left = progressArea.offsetLeft + (progressArea.offsetWidth-210)/2 + 'px';
	progressWin.style.top = progressArea.offsetTop + (progressArea.offsetHeight-140)/2 + 'px';
	progressWin.style.display='inline'
	return true;
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<iframe width="200" height="110" id="progressBar" src="includes/progress_bar.asp" style="display:none; position:absolute; left:0px; top:0px;" frameborder="0" scrolling="no"></iframe>
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% If lngTopicID <> 0 Then Response.Write(strTxtTopic & " " & strTxtSearch) Else Response.Write(strTxtSearchTheForum) %></h1></td>
 </tr>
</table> 
<br />
<div id="progressFormArea">
<form method="post" name="frmSearch" id="frmSearch" action="search_process.asp<% = strQsSID1 %>" onSubmit="return CheckForm();" onReset="return confirm('<% = strResetFormConfirm %>');">
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td colspan="2" align="left"><% = strTxtSearchKeywords %></td>
 </tr>
 <tr class="tableRow">
  <td align="left" width="50%" valign="top"><% = strTxtSearchbyKeyWord %>
   <br />
   <input name="KW" id="KW" type="text" value="<% = Server.HTMLEncode(strSearchKeywords) %>" size="25" maxlength="35" tabindex="1" />
   <select name="searchType" id="searchType" tabindex="2">
    <option value="allWords" selected="selected"><% = strTxtMatch & " " & strTxtAllWords %></option>
    <option value="anyWords"><% = strTxtMatch & " " & strTxtAnyWords %></option>
    <option value="phrase"><% = strTxtMatch & " " & strTxtPhrase %></option><%

'If an admin or moderator let them search on IP address
If blnAdmin OR blnModerator Then
%>   	
    <option value="IP"><% = strTxtMatch & " " & strTxtIPAddress %></option><%

End If   
%>    
   </select><%
   
'Include the follwoing if doing a topic search
If lngTopicID <> 0 Then
%>
   <input name="TID" type="hidden" id="TID" value="<% = lngTopicID %>" />
   <input name="qTopic" type="hidden" id="qTopic" value="1" /><%
End If

%>    
  </td>
  <td height="2" width="50%" valign="top"><% = strTxtSearchbyUserName %><br />
   <input name="USR" id="USR" type="text" value="<% = Server.HTMLEncode(strSearchUser) %>" maxlength="20" tabindex="3" /> 
   <a href="javascript:winOpener('pop_up_member_search.asp?RP=SEARCH','memSearch',0,1,580,355)"><img src="<% = strImagePath %>member_search.png" alt="<% = strTxtFindMember %>" title="<% = strTxtFindMember %>" border="0" align="absbottom"></a>
   <br />
   <input name="UsrMatch" id="UsrMatch" type="checkbox" value="true" tabindex="4"> <% = strTxtMemberName2 & " " & strTxtExactMatch %><%

'If seraching a topic don't include the other options
If lngTopicID = 0 Then
	
%>
   <br />
   <input name="UsrTopicStart" id="UsrTopicStart" type="checkbox" value="true" tabindex="5"> <% = strTxtSearchForTopicsThisMemberStarted %><%

End If
%>
  </td>
 </tr><%

'If seraching a topic don't include the other options
If lngTopicID = 0 Then
	
%>
 <tr class="tableLedger">
  <td colspan="2" align="left"><% = strTxtSearchOptions %></td>
 </tr>
 <tr class="tableRow">
  <td align="left" width="50%"><% = strTxtSearchForum %>
   <br /><span class="smText"><% = strTxtCtrlApple %></span>
   <br />
   <select name="forumID" id="forumID" size="13" multiple="multiple" tabindex="6">
    <option value="0"<% If intForumID2 = 0 Then Response.Write(" selected=""selected""") %>><% = strTxtAllForums %></option><%


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
			Response.Write(vbCrLf & "    <optgroup label=""&nbsp;&nbsp;" & strCatName & """>")
			
			
			
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
					Response.Write (vbCrLf & "     <option value=""" & intForumID & """")
					If intForumID2 = intForumID Then Response.Write(" selected=""selected""")
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
						strForumURL = sarryForumSelect(7,intCurrentRecord)
				
						'Remove any parts that could be mistaken for a forum URL
						If strForumURL = "http://" OR isNull(strForumURL) Then strForumURL = ""
						
						'If this forum is to be hidden but the user is allowed access to it set the hidden boolen back to false
						If blnHideForum = True AND blnRead = True Then blnHideForum = False
			
						'If the forum is not a hidden forum to this user, display it
						If blnHideForum = False AND strForumURL = "" Then
							'Display a link in the link list to the forum
							Response.Write (vbCrLf & "     <option value=""" & intSubForumID & """")
							If intForumID2 = intSubForumID Then Response.Write(" selected=""selected""")
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
    </select>
  </td>
  <td width="50%" valign="top">
   <table width="100%" border="0" cellspacing="4" cellpadding="2">
    <tr>
     <td><% = strTxtSearchIn %>
      <br />
      <select name="searchIn" id="searchIn" tabindex="7">
       <option value="body" selected="selected"><% = strTxtMessageBody %></option>
       <option value="subject"><% = strTxtTopicSubject %></option>
      </select>
     </td>
    </tr>
    <tr>
     <td><% = strTxtFindPosts %>
      <br />
      <select name="AGE" id="AGE" tabindex="8">
       <option value="0"<% If intSearchTimeDefault = 0 Then Response.Write(" selected=""selected""")%>><% = strTxtAnyDate %></option>
       <option value="1"<% If intSearchTimeDefault = 1 Then Response.Write(" selected=""selected""")%>><% = DateFormat(dtmLastVisitDate) & " " & strTxtAt & " " & TimeFormat(dtmLastVisitDate) %></option>
       <option value="2"<% If intSearchTimeDefault = 2 Then Response.Write(" selected=""selected""")%>><% = strTxtYesterday %></option>
       <option value="3"<% If intSearchTimeDefault = 3 Then Response.Write(" selected=""selected""")%>><% = strTxtLastWeek %></option>
       <option value="4"<% If intSearchTimeDefault = 4 Then Response.Write(" selected=""selected""")%>><% = strTxtLastMonth %></option>
       <option value="5"<% If intSearchTimeDefault = 5 Then Response.Write(" selected=""selected""")%>><% = strTxtLastTwoMonths %></option>
       <option value="6"<% If intSearchTimeDefault = 6 Then Response.Write(" selected=""selected""")%>><% = strTxtLastSixMonths %></option>
       <option value="7"<% If intSearchTimeDefault = 7 Then Response.Write(" selected=""selected""")%>><% = strTxtLastYear %></select>
      <select name="DIR" id="DIR" tabindex="9">
       <option value="newer"><% = strTxtAndNewer %></option>
       <option value="older"><% = strTxtAndOlder %></option>
      </select>
     </td>
    </tr>
    <tr>
     <td><% = strTxtSortResultsBy %>
      <br />
      <select name="OrderBy" id="OrderBy" tabindex="10">
       <option value="LastPost" selected="selected"><% = strTxtLastPostTime %></option>
       <option value="StartDate"><% = strTxtTopicStartDate %></option>
       <option value="Subject"><% = strTxtSubjectAlphabetically %></option>
       <option value="Replies"><% = strTxtNumberReplies %></option>
       <option value="Views"><% = strTxtNumberViews %></option>
       <option value="Username"><% = strTxtUsername %></option>
       <option value="ForumName"><% = strTxtForumName %></option>
      </select>
     </td>
    </tr>
    <tr>
     <td><% = strTxtDisplayResultsAs %>
      <br />
      <select name="resultType" id="resultType" tabindex="11">
       <option value="posts" selected="selected"><% = strTxtPosts %></option>
       <option value="topics"><% = strTxtTopics %></option>
      </select>
     </td>
    </tr>
   </table>
  </td>
 </tr><%

End If

%>
 <tr class="tableBottomRow">
  <td colspan="2" align="center">
   <input type="submit" name="Submit" id="Submit" value="<% = strTxtStartSearch %>" tabindex="12" />
   <input type="reset" name="Reset" id="Reset" value="<% = strTxtResetForm %>" tabindex="13" />
  </td>
 </tr>
 </table>
</form>
</div>
<br />
<div align="center"><%

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
%>
</div>
<!-- #include file="includes/footer.asp" -->
