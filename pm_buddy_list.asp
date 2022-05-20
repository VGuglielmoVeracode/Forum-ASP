<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="language_files/pm_language_file_inc.asp" -->
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
Dim intRowColourNumber		'Holds the number to calculate the table row colour	
Dim blnIsUserOnline		'Set to true if the user is online
Dim intPageSize			'Holds the number of memebrs shown per page
Dim intStartPosition		'Holds the start poition for records to be shown
Dim intEndPosition		'Holds the end poition for records to be shown
Dim intCurrentRecord		'Holds the current record position
Dim lngTotalRecords		'Holds the total number of therads in this topic
Dim lngTotalRecordsPages	'Holds the total number of pages
Dim sarryPmBuddy		'Holds the buddy list array
Dim intArrayPass		'Loop variable for online users	
Dim strFormID			'Holds the ID for the form	


'Initialise variable
blnSslEnabledPage = True
intRowColourNumber = 0
intCurrentRecord = 0


'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)

End If



'If Priavte messages are not on then send them away
If blnPrivateMessages = False Then 
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
End If


'If the user is not allowed then send them away
If intGroupID = 2 OR blnActiveMember = False OR blnBanned Then 
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If




'Get the users buddy detals from the db
	
'Initlise the sql statement
strSQL = "SELECT " & strDbTable & "BuddyList.Buddy_ID, " & strDbTable & "BuddyList.Address_ID, " & strDbTable & "BuddyList.Description, " & strDbTable & "BuddyList.Block, " & strDbTable & "Author.Username, " & strDbTable & "Author.Author_ID " & _
"FROM " & strDbTable & "Author" & strDBNoLock & ", " & strDbTable & "BuddyList" & strDBNoLock & " "& _
"WHERE " & strDbTable & "Author.Author_ID=" & strDbTable & "BuddyList.Buddy_ID " & _
	"AND " & strDbTable & "BuddyList.Author_ID=" & lngLoggedInUserID & " " & _
	"AND " & strDbTable & "BuddyList.Buddy_ID <> 2 " & _
"ORDER BY " & strDbTable & "BuddyList.Block ASC, " & strDbTable & "Author.Username ASC;" 
	
	
'Query the database
rsCommon.Open strSQL, adoCon

'If not eof then get some details
If NOT rsCommon.EOF Then 
	
	'Read in the row from the db using getrows for better performance
	sarryPmBuddy = rsCommon.GetRows()
End If	
	
'Close rs
rsCommon.Close


'If active users is enabled update the active users application array
If blnActiveUsers Then
	'Call active users function
	saryActiveUsers = activeUsers(strTxtPrivateMessenger & " " & strTxtBuddyList, "", "", 0)
End If


'get form ID
strFormID = getSessionItem("KEY")


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtBuddyList

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtBuddyList %></title>

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
	
	var errorMsg = "";
	var formArea = document.getElementById('frmBuddy');
	
	//Check for a buddy
	if (formArea.username.value==""){
		errorMsg += "\n<% = strTxtNoBuddyErrorMsg %>";
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
	
	document.getElementById('formID').value='<% = strFormID %>';
	
	return true;
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtBuddyList %></h1></td>
 </tr>
</table>
<br />
<table class="basicTable" cellspacing="0" cellpadding="0" align="center"> 
 <tr> 
  <td class="tabTable">
   <a href="member_control_panel.asp<% = strQsSID1 %>" title="<% = strTxtControlPanel %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>member_control_panel.<% = strForumImageType %>" border="0" alt="<% = strTxtControlPanel %>" /> <% = strTxtControlPanel %></a>
   <a href="register.asp<% = strQsSID1 %>" title="<% = strTxtProfile2 %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>profile.<% = strForumImageType %>" border="0" alt="<% = strTxtProfile2 %>" /> <% = strTxtProfile2 %></a><%
 
If blnEmail Then

%>
   <a href="email_notify_subscriptions.asp<% = strQsSID1 %>" title="<% = strTxtSubscriptions %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>subscriptions.<% = strForumImageType %>" border="0" alt="<% = strTxtSubscriptions %>" /> <% = strTxtSubscriptions %></a><%
End If

%>
   <a href="pm_buddy_list.asp<% = strQsSID1 %>" title="<% = strTxtBuddyList %>" class="tabButtonActive">&nbsp;<img src="<% = strImagePath %>buddy_list.<% = strForumImageType %>" border="0" alt="<% = strTxtBuddyList %>" /> <% = strTxtBuddyList %></a><%

'If file/image uploading enabled
If blnAttachments OR blnImageUpload Then

%>
   <a href="file_manager.asp<% = strQsSID1 %>" title="<% = strTxtFileManager %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>file_manager.<% = strForumImageType %>" border="0" alt="<% = strTxtFileManager %>" /> <% = strTxtFileManager %></a><%

End If



%>
  </td>
 </tr>
</table>
<br />
<form method="post" name="frmBuddy" id="frmBuddy" action="pm_add_buddy.asp<% = strQsSID1 %>" onSubmit="return CheckForm();" onReset="return ResetForm();">
  <table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
    <tr class="tableLedger">
         <td width="23%"><% = strTxtMemberName %></td>
        <td width="32%"><% = strTxtDescription %></td>
        <td width="24%"><% = strTxtAllowThisMemberTo %></td>
        <td width="21%">&nbsp;</td>
       </tr>
       <tr class="tableRow">
        <td>
          <input type="text" name="username" id="username" size="15" maxlength="25" value="<% If IntC(Request.QueryString("code")) <> 2 Then Response.Write(Server.HTMLEncode(decodeString(Request.QueryString("name")))) %>">
          <a href="javascript:winOpener('pop_up_member_search.asp?RP=BUD<% = strQsSID2 %>','memSearch',0,1,580,355)"><img src="<% = strImagePath %>member_search.<% = strForumImageType %>" alt="<% = strTxtFindMember %>" title="<% = strTxtFindMember %>" border="0" align="absmiddle"></a> 
         </td>
        <td> 
         <input type="text" name="description" id="description" size="25" maxlength="30" value="<% If IntC(Request.QueryString("code")) <> 2 Then Response.Write(Server.HTMLEncode(Request.QueryString("desc"))) %>">
        </td>
        <td> 
         <select name="blocked" id="blocked">
          <option value="False" selected><% = strTxtMessageMe %></option>
          <option value="True"><% = strTxtNotMessageMe %></option>
         </select>
        </td>
        <td width="21%" align="right"><input type="hidden" name="formID" id="formID" value="" /> <input type="submit" name="Submit" id="Submit" value="<% = strTxtAddToBuddy %>"></td>
       </tr>
  </table>
</form>
<br />
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
    <tr class="tableLedger">
     <td width="19%"><% = strTxtBuddy %></td>
     <td width="40%"><% = strTxtDescription %></td>
     <td width="35%"><% = strTxtContactStatus %></td><%
If blnActiveUsers Then %>
     <td width="6%"  align="center"><% = strTxtOnLine2 %></td><%
End If %>
     <td width="6%" align="center"><% = strTxtDelete %></td>
    </tr><%
    

    
'Check there are PM messages to display
If isArray(sarryPmBuddy) = false Then

	'If there are no pm messages to display then display the appropriate error message
	Response.Write(vbCrLf & "    <tr class=""tableRow""><td colspan=""5"" align=""center""><br />" & strTxtNoBuddysInList & "<br /><br /></td></tr>")

'Else there the are topic's so write the HTML to display the topic names and a discription
Else 	
	
	'Loop round to read in all the Topics in the database
	Do while intCurrentRecord =< UBound(sarryPmBuddy, 2)
	
	
		'SQL Query Array Look Up table
		'0 = Buddy_ID
		'1 = Address_ID
		'2 = Description
		'3 = Block
		'4 = Username
		'5 = Author_ID

		
		'Get the row number
		intRowColourNumber = intRowColourNumber + 1
	%>
    <tr class="<% If (intRowColourNumber MOD 2 = 0 ) Then Response.Write("evenTableRow") Else Response.Write("oddTableRow") %>"> 
     <td><a href="member_profile.asp?PF=<% = sarryPmBuddy(0,intCurrentRecord) & strQsSID2 %>" rel="nofollow"><% = sarryPmBuddy(4,intCurrentRecord) %></a></td>
     <td><% = sarryPmBuddy(2,intCurrentRecord) %>&nbsp;</td>
     <td><%
     		'Get the contact status
     		If CBool(sarryPmBuddy(3,intCurrentRecord)) = True Then
     			Response.Write(strTxtThisPersonCanNotMessageYou)
     		Else
     			Response.Write(strTxtThisPersonCanMessageYou)
     		End If
     %></td><%
		'If active users is enabled see if any buddies are online
		If blnActiveUsers Then 
			
			'Initilase variable
			blnIsUserOnline = False
			
			'Get the users online status
			For intArrayPass = 1 To UBound(saryActiveUsers, 2)
				If saryActiveUsers(1, intArrayPass) = CLng(sarryPmBuddy(0,intCurrentRecord)) Then blnIsUserOnline = True
			Next
			
			%>
     <td align="center"><% If blnIsUserOnline Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""" & strTxtOnLine2 & """ title=""" & strTxtOnLine2 & """ />") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""" & strTxtOffLine & """ title=""" & strTxtOffLine & """ />") %></td><%
		 	
		End If 
		
		%>
     <td align="center"><a href="pm_delete_buddy.asp?pm_id=<% = sarryPmBuddy(1,intCurrentRecord) %>&XID=<% = strFormID & strQsSID2 %>" OnClick="return confirm('<% = strTxtDeleteBuddyAlert %>')"><img src="<% = strImagePath %>delete.<% = strForumImageType %>" width="15" height="16" alt="<% = strTxtDelete %>" title="<% = strTxtDelete %>" border="0"></a></td>
    </tr><%

		
		'Move to the next record
		intCurrentRecord = intCurrentRecord + 1
	Loop
End If


%>
</table>
<br />
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
   <td><!-- #include file="includes/forum_jump_inc.asp" --></td>
  </tr>
</table>
<div align="center"><br />
 <%
'Clear server objects
Call closeDatabase()

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
<%
'Display a msg letting the user know any add or delete details to the buddy list
Select Case Request.QueryString("ER")
	Case "1"
		Response.Write("<script  language=""JavaScript"">")
		Response.Write("alert('" & Replace(Server.HTMLEncode(Request.QueryString("name")), "'", "\'", 1, -1, 1) & " " & strTxtIsAlreadyInYourBuddyList & ".');")
		Response.Write("</script>")
	Case "2"
		Response.Write("<script  language=""JavaScript"">")
		Response.Write("alert('" & Replace(Server.HTMLEncode(Request.QueryString("name")), "'", "\'", 1, -1, 1) & ", " & strTxtUserCanNotBeFoundInDatabase & ".');")
		Response.Write("</script>")
End Select
%>
<!-- #include file="includes/footer.asp" -->