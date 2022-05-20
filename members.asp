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

'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"

'Dimension variables
Dim strUsername			'Holds the users username
Dim strHomepage			'Holds the users homepage if they have one
Dim strEmail			'Holds the users e-mail address
Dim blnShowEmail		'Boolean set to true if the user wishes there e-mail address to be shown
Dim lngUserID			'Holds the new users ID number
Dim lngNumOfPoints		'Holds the number of posts the user has made
Dim intMemberGroupID		'Holds the users interger group ID
Dim strMemberGroupName		'Holds the umembers group name
Dim intRankStars		'holds the number of rank stars the user holds
Dim dtmRegisteredDate		'Holds the date the usre registered
Dim lngTotalRecordsPages	'Holds the total number of pages
Dim lngTotalRecords		'Holds the total number of forum members
Dim intRecordPositionPageNum	'Holds the page number we are on
Dim dtmLastPostDate		'Holds the date of the users las post
Dim intLinkPageNum		'Holds the page number to link to
Dim strSearchCriteria		'Holds the search critiria
Dim strSortBy			'Holds the way the records are sorted
Dim intGetGroupID		'Holds the group ID
Dim strRankCustomStars		'Holds custom stars for the user group
Dim sarryMembers		'Holds the getrows db call for members
Dim intPageSize			'Holds the number of memebrs shown per page
Dim intStartPosition		'Holds the start poition for records to be shown
Dim intEndPosition		'Holds the end poition for records to be shown
Dim intCurrentRecord		'Holds the current record position
Dim dtmLastActiveDate		'Holds the date this user was last active
Dim strSortDirection		'Holds the sort order
Dim intPageLinkLoopCounter	'Holds the loop counter for the page links
Dim strPassword
Dim strGroupName
Dim intSelGroupID


'Initalise variables
blnSslEnabledPage = True
blnShowEmail = False
intPageSize = 25


'If a login is displayed then switch to HTTPS
If intGroupID = 2 Then blnSslEnabledPage = True



'Redirect if member list is not enabled
If blnAdmin = False AND blnDisplayMemberList = False Then
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
End If



'If this is the first time the page is displayed then the members record position is set to page 1
If isNumeric(Request.QueryString("PN")) = false Then
	intRecordPositionPageNum = 1
	
ElseIf Request.QueryString("PN") < 1 Then
	intRecordPositionPageNum = 1
	
'Else the page has been displayed before so the record postion is set to the Record Position number
Else
	intRecordPositionPageNum = IntC(Request.QueryString("PN"))
End If

'Start position
intStartPosition = ((intRecordPositionPageNum - 1) * intPageSize)	




'Read in goup ID
If isNumeric(Request.QueryString("GID")) Then intGetGroupID = IntC(Request.QueryString("GID")) Else intGetGroupID = 0


'Get the search critiria for the members to display
If NOT Request.QueryString("SF") = "" Then
	strSearchCriteria = Trim(Mid(Request.QueryString("SF"), 1, 15))
End If


'Get rid of milisous code
strSearchCriteria = formatSQLInput(strSearchCriteria)

'Get the sort critiria
Select Case Request.QueryString("SO")
	Case "PT"
		strSortBy = strDbTable & "Author.Points "
	Case "LU"
		strSortBy = strDbTable & "Author.Join_date "
	Case "OU"
		strSortBy = strDbTable & "Author.Join_date "
	Case "GP"
		strSortBy = strDbTable & "Group.Name "
	Case "LA"
		strSortBy = strDbTable & "Author.Last_visit "
	Case Else
		strSortBy = strDbTable & "Author.Username "
End Select

'Sort the direction of db results
If Request.QueryString("OB") = "desc" Then
	strSortDirection = "asc"
	strSortBy = strSortBy & "DESC"
Else
	strSortDirection = "desc"
	strSortBy = strSortBy & "ASC"
End If



'Read in from db
If intGroupID <> 2 Then

		
		
	'If using advanced paging then we need to count the total number of records
	If (strDatabaseType = "SQLServer" AND blnSqlSvrAdvPaging) OR strDatabaseType = "mySQL" Then
		strSQL = "" & _
		"SELECT Count(" & strDbTable & "Author.Author_ID) AS MemberCount " & _
		"FROM " & strDbTable & "Author " & strDBNoLock & ", " & strDbTable & "Group " & strDBNoLock & " " & _
		"WHERE " & strDbTable & "Author.Group_ID = " & strDbTable & "Group.Group_ID "
		
		'If a search criteria
		If NOT strSearchCriteria = "" Then	
			strSQL = strSQL & "AND " & strDbTable & "Author.Username LIKE '" & strSearchCriteria & "%' "
		End If
		
		'If a group id to look for
		If NOT intGetGroupID = 0 Then	
		 	strSQL = strSQL & "AND " & strDbTable & "Author.Group_ID ="  & intGetGroupID & " "
		End If
		
		'If not admin or modertor do not included un-active or banned members
		If blnAdmin = False AND blnModerator = False Then
			strSQL = strSQL & "" & _
			"AND Active = " & strDBTrue & " " & _
			"AND Banned = " & strDBFalse & ";"
		End If
			
	
		'Set error trapping
		On Error Resume Next
			
		'Query the database
		rsCommon.Open strSQL, adoCon
		
		'If an error has occurred write an error to the page
		If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "member_count", "member.asp")
					
		'Disable error trapping
		On Error goto 0
		
		'Read in member count from database
		lngTotalRecords = CLng(rsCommon("MemberCount"))
		
		'Close recordset
		rsCommon.close
	End If
	
	
	
	'Initalise the strSQL variable with an SQL statement to query the database
	'Read in all the topics for this forum and place them in an array
	strSQL = "" & _
	"SELECT "
	
	'If SQL server advanced paging
	If strDatabaseType = "SQLServer" AND blnSqlSvrAdvPaging Then
		strSQL = strSQL & " * " & _
		"FROM (SELECT TOP " & intPageSize * intRecordPositionPageNum  & " "
	End If
	
	strSQL = strSQL & _
	strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.Group_ID, " & strDbTable & "Author.Last_visit, " & strDbTable & "Author.Points, " & strDbTable & "Author.Join_date, " & strDbTable & "Author.Active, " & strDbTable & "Group.Name, " & strDbTable & "Group.Stars, " & strDbTable & "Group.Custom_stars "
	
	
	
	'If SQL Server advanced paging
	If strDatabaseType = "SQLServer"  AND blnSqlSvrAdvPaging Then
		strSQL = strSQL & ", ROW_NUMBER() OVER (ORDER BY " & strSortBy & ") AS RowNum "
	
	End If
	
	strSQL = strSQL & "" & _
	"FROM " & strDbTable & "Author " & strDBNoLock & ", " & strDbTable & "Group " & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Author.Group_ID = " & strDbTable & "Group.Group_ID "
	
	'If a search criteria
	If NOT strSearchCriteria = "" Then	
		strSQL = strSQL & "AND " & strDbTable & "Author.Username LIKE '" & strSearchCriteria & "%' "
	End If
		
	'If a group id to look for
	If NOT intGetGroupID = 0 Then	
		strSQL = strSQL & "AND " & strDbTable & "Author.Group_ID ="  & intGetGroupID & " "
	End If
	
	'If not admin or modertor do not included un-active or banned members
	If blnAdmin = False AND blnModerator = False Then
		strSQL = strSQL & "" & _
		"AND Active = " & strDBTrue & " " & _
		"AND Banned = " & strDBFalse & " "
	End If
	
	'If SQL Server advanced paging
	If strDatabaseType = "SQLServer" AND blnSqlSvrAdvPaging Then
		strSQL = strSQL & ") AS PagingQuery WHERE RowNum BETWEEN " & intStartPosition + 1 & " AND " & intStartPosition + intPageSize & " "
	
	'Else Order by clause here
	Else
		strSQL = strSQL & "ORDER BY " & strSortBy & " "
	End If
	
	'mySQL limit operator
	If strDatabaseType = "mySQL" Then
		strSQL = strSQL & " LIMIT " & intStartPosition & ", " & intPageSize
	End If
	
	strSQL = strSQL & ";"



	'Query the database  
	rsCommon.Open strSQL, adoCon

	'If there are records get em from rs
	If NOT rsCommon.EOF Then
		
		'Read in the row from the db using getrows for better performance
		sarryMembers = rsCommon.GetRows()
		
		
		
		'If advanced paging then workout the end and start position differently
		If (strDatabaseType = "SQLServer" AND blnSqlSvrAdvPaging) OR strDatabaseType = "mySQL" Then
			
			'End Position
			intEndPosition = Ubound(sarryMembers,2) + 1
		
			'Get the start position
			intCurrentRecord = 0
		
		'Else standard slower paging	
		Else
			'Count the number of records
			lngTotalRecords = Ubound(sarryMembers,2) + 1
		
			'Start position
			intStartPosition = ((intRecordPositionPageNum - 1) * intPageSize)
		
			'End Position
			intEndPosition = intStartPosition + intPageSize
		
			'Get the start position
			intCurrentRecord = intStartPosition
		End If

		
		'Count the number of pages for the topics using '\' so that any fraction is omitted 
		lngTotalRecordsPages = lngTotalRecords \ intPageSize
		
		'If there is a remainder or the result is 0 then add 1 to the total num of pages
		If lngTotalRecords Mod intPageSize > 0 OR lngTotalRecordsPages = 0 Then lngTotalRecordsPages = lngTotalRecordsPages + 1
		
	End If
	
	
	'Close the recordset as it is no longer needed
	rsCommon.Close

End If


'Page to link to for mutiple page (with querystrings if required)
strLinkPage = "members.asp?SF=" & Server.URLEncode(Request.QueryString("SF")) & "&GID=" & intGetGroupID & "&"


'If active users is enabled update the active users application array
If blnActiveUsers Then
	'If not logged in display in active users that the user is blocked
	If intGroupID = 2 Then
		'Call active users function
		saryActiveUsers = activeUsers(strTxtViewing & " " & strTxtForumMembers & " [" & strTxtAccessDenied & "]", "", "", 0)
	ElseIf Request.QueryString("SF") = "" Then
		'Call active users function
		saryActiveUsers = activeUsers("", strTxtViewing & " " & strTxtForumMembers, "members.asp?PN=" & intRecordPositionPageNum, 0)
	Else
		'Call active users function
		saryActiveUsers = activeUsers(strTxtViewing & " " & strTxtForumMembers, strTxtSearchingFor & ": &#8216;" & Server.HTMLEncode(Request.QueryString("SF")) & "&#8217;", strLinkPage & "PN=" & intRecordPositionPageNum, 0)
	End If
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & "<a href=""members.asp" & strQsSID1 & """>" & strTxtForumMembers & "</a>"

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strMainForumName & " " & strTxtMembers %><% If lngTotalRecordsPages > 1 Then Response.Write(" - " & strTxtPage & " " & intRecordPositionPageNum) %></title>
<meta name="generator" content="Web Wiz Forums" />
<meta name="robots" content="noindex, follow" />

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


<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtForumMembers %></h1></td>
 </tr>
</table>
<br /><%
    
    
'If the users account is suspended then let them know
If blnActiveMember = false OR blnBanned Then
	
	Response.Write(vbCrLf & "<table class=""errorTable"" cellspacing=""1"" cellpadding=""3"" align=""center"">" & _
	vbCrLf & " <tr>" & _
    	vbCrLf & "  <td><img src=""" & strImagePath & "error.png"" alt=""" & strTxtError & """ /> <strong>" & strTxtError & "</strong></td>" & _
  	vbCrLf & " </tr>" & _
	vbCrLf & " <tr>" & _
	vbCrLf & "  <td>")
	
	'If mem suspended display message
	If blnBanned Then
		Response.Write("<strong>" & strTxtForumMemberSuspended & "</strong>")
	
	'Else account not yet active
	ElseIf blnActiveMember = False Then
		
		Response.Write("<br />" & strTxtForumMembershipNotAct)
		If blnMemberApprove = False Then Response.Write("<br /><br />" & strTxtToActivateYourForumMem)
		
		'If admin activation is enabled let the user know
		If blnMemberApprove Then
			Response.Write("<br />" & strTxtYouAdminNeedsToActivateYourMembership)
		'If email is on then place a re-send activation email link
		ElseIf blnEmailActivation AND blnLoggedInUserEmail Then 
				Response.Write("<br /><br /><a href=""javascript:winOpener('resend_email_activation.asp" & strQsSID1 & "','actMail',1,1,475,300)"">" & strTxtResendActivationEmail & "</a>")
		End If
	End If
	
	
	Response.Write(vbCrLf & "  </td>" & _
	vbCrLf & " </tr>" & _
	vbCrLf & "</table>" & _
	vbCrLf & "<br /><br />")
	
'If the user has not logged in dispaly an error message
ElseIf intGroupID = 2 Then
	
	Response.Write(vbCrLf & "<table class=""errorTable"" cellspacing=""1"" cellpadding=""3"" align=""center"">" & _
	vbCrLf & " <tr>" & _
    	vbCrLf & "  <td><img src=""" & strImagePath & "error.png"" alt=""" & strTxtError & """ /> <strong>" & strTxtError & "</strong></td>" & _
  	vbCrLf & " </tr>" & _
	vbCrLf & " <tr>" & _
	vbCrLf & "  <td>" & strTxtMustBeRegistered  & "</td>" & _
	vbCrLf & " </tr>" & _
	vbCrLf & "</table>")
	%><!--#include file="includes/login_form_inc.asp" --><%

'If the user has logged in then read in the members from the database and dispaly them
Else

	'If there are no memebers to display then show an error message
	If lngTotalRecords <= 0 Then
		
		Response.Write(vbCrLf & "<table class=""errorTable"" cellspacing=""1"" cellpadding=""3"" align=""center"">" & _
		vbCrLf & " <tr>" & _
	    	vbCrLf & "  <td><img src=""" & strImagePath & "error.png"" alt=""" & strTxtError & """ /> <strong>" & strTxtError & "</strong></td>" & _
	  	vbCrLf & " </tr>" & _
		vbCrLf & " <tr>" & _
		vbCrLf & "  <td>" & strTxtSorryYourSearchFoundNoMembers & "</td>" & _
		vbCrLf & " </tr>" & _
		vbCrLf & "</table>" & _
		vbCrLf & "<br />")

	End If


%>
<form name="frmMemberSearch" id="frmMemberSearch" method="get" action="members.asp">
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td colspan="2" align="left"><% = strTxtMemberSearch %></td>
 </tr>
 <tr class="tableRow">
  <td align="left"><% = strTxtMemberSearch %>:
   <input name="SF" id="SF" size="20" maxlength="20" value="<% = Server.HTMLEncode(Request.QueryString("SF")) %>" />
   <input type="hidden" name="SID" id="SID" value="<% = strQsSID %>" />
   &nbsp; <% = strTxtGroup %>:
   <select name="GID">
	<option value="0" selected="selected">-- <% = strTxtSelect & " " & strTxtGroup %>  --</option><%

'Read in the group name from the database
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Group.*, " & strDbTable & "LadderGroup.* " & _
	"FROM " & strDbTable & "Group " & _
	"LEFT JOIN " & strDbTable & "LadderGroup ON " & strDbTable & "Group.Ladder_ID = " & strDbTable & "LadderGroup.Ladder_ID " & _
	"ORDER BY " & strDbTable & "LadderGroup.Ladder_Name ASC, " & strDbTable & "Group.Minimum_posts ASC, " & strDbTable & "Group.Group_ID ASC;"

'Query the database
rsCommon.Open strSQL, adoCon


'Loop through all the categories in the database
Do while NOT rsCommon.EOF

	'Read in the deatils for the category
	strGroupName = rsCommon("Name")
	intSelGroupID = CInt(rsCommon("Group_ID"))

	'Display a link in the link list to the cat
	Response.Write (vbCrLf & "		<option value=""" & intSelGroupID & """")
	If intSelGroupID = intGetGroupID Then Response.Write(" selected=""selected""")
	Response.Write(">" & strGroupName & "</option>")


	'Move to the next record in the recordset
	rsCommon.MoveNext
Loop

'Reset server objects
rsCommon.Close


%>
      </select>
   <input type="submit" name="Submit" id="Submit" value="<% = strTxtSearch %>" />
   <br /><br />
   <a href="members.asp?GID=<% = intGetGroupID %><% = strQsSID1 %>">#</a> 
   <a href="members.asp?SF=A&GID=<% = intGetGroupID %><% = strQsSID2 %>">A</a> 
   <a href="members.asp?SF=B&GID=<% = intGetGroupID %><% = strQsSID2 %>">B</a> 
   <a href="members.asp?SF=C&GID=<% = intGetGroupID %><% = strQsSID2 %>">C</a>
   <a href="members.asp?SF=D&GID=<% = intGetGroupID %><% = strQsSID2 %>">D</a> 
   <a href="members.asp?SF=E&GID=<% = intGetGroupID %><% = strQsSID2 %>">E</a> 
   <a href="members.asp?SF=F&GID=<% = intGetGroupID %><% = strQsSID2 %>">F</a>
   <a href="members.asp?SF=G&GID=<% = intGetGroupID %><% = strQsSID2 %>">G</a> 
   <a href="members.asp?SF=H&GID=<% = intGetGroupID %><% = strQsSID2 %>">H</a> 
   <a href="members.asp?SF=I&GID=<% = intGetGroupID %><% = strQsSID2 %>">I</a>
   <a href="members.asp?SF=J&GID=<% = intGetGroupID %><% = strQsSID2 %>">J</a> 
   <a href="members.asp?SF=K&GID=<% = intGetGroupID %><% = strQsSID2 %>">K</a> 
   <a href="members.asp?SF=L&GID=<% = intGetGroupID %><% = strQsSID2 %>">L</a>
   <a href="members.asp?SF=M&GID=<% = intGetGroupID %><% = strQsSID2 %>">M</a> 
   <a href="members.asp?SF=N&GID=<% = intGetGroupID %><% = strQsSID2 %>">N</a> 
   <a href="members.asp?SF=O&GID=<% = intGetGroupID %><% = strQsSID2 %>">O</a>
   <a href="members.asp?SF=P&GID=<% = intGetGroupID %><% = strQsSID2 %>">P</a> 
   <a href="members.asp?SF=Q&GID=<% = intGetGroupID %><% = strQsSID2 %>">Q</a> 
   <a href="members.asp?SF=R&GID=<% = intGetGroupID %><% = strQsSID2 %>">R</a>
   <a href="members.asp?SF=S&GID=<% = intGetGroupID %><% = strQsSID2 %>">S</a> 
   <a href="members.asp?SF=T&GID=<% = intGetGroupID %><% = strQsSID2 %>">T</a> 
   <a href="members.asp?SF=U&GID=<% = intGetGroupID %><% = strQsSID2 %>">U</a>
   <a href="members.asp?SF=V&GID=<% = intGetGroupID %><% = strQsSID2 %>">V</a> 
   <a href="members.asp?SF=W&GID=<% = intGetGroupID %><% = strQsSID2 %>">W</a> 
   <a href="members.asp?SF=X&GID=<% = intGetGroupID %><% = strQsSID2 %>">X</a>
   <a href="members.asp?SF=Y&GID=<% = intGetGroupID %><% = strQsSID2 %>">Y</a> 
   <a href="members.asp?SF=Z&GID=<% = intGetGroupID %><% = strQsSID2 %>">Z</a>
  </td>
 </tr>
</table>
</form>
<br />
<table class="basicTable" cellspacing="0" cellpadding="10" align="center">
 <tr>
  <td align="left"><%

	'Display some text on search
	If lngTotalRecords > 0 Then
		
		Response.Write(strTxtSearchResults & " ")
		
		'If this is a keyword search display keywrds
		If strSearchCriteria <> "" Then 
			Response.Write(strTxtFor & " '" & Server.HTMLEncode(strSearchCriteria) & "' ")
		End If
		
		Response.Write(strTxtHasFound & " " & FormatNumber(lngTotalRecords, 0) & " " & strTxtResultsIn & " " & FormatNumber(Timer() - dblStartTime, 4) & " " & strTxtSecounds & ".")
	End If
%></td>
  <td align="right" nowrap>
   <!-- #include file="includes/page_link_inc.asp" -->
  </td>
 </tr>
</table>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td width="20%"><a href="members.asp?SF=<% = Server.URLEncode(Request.QueryString("SF")) %>&GID=<% = intGetGroupID %>&SO=UN<% = strQsSID2 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtUsername %></a><% If Request.QueryString("SO") = "UN" OR Request.QueryString("SO") = "" Then Response.Write(" <a href=""members.asp?SF=" & Server.URLEncode(Request.QueryString("SF")) & "&GID=" & intGetGroupID & "&SO=UN&OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %></td>
  <td width="20%"><a href="members.asp?SF=<% = Server.URLEncode(Request.QueryString("SF")) %>&GID=<% = intGetGroupID %>&SO=GP<% = strQsSID2 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtType %></a><% If Request.QueryString("SO") = "GP" Then Response.Write(" <a href=""members.asp?SF=" & Server.URLEncode(Request.QueryString("SF")) & "&GID=" & intGetGroupID & "&SO=GP&OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %></td>
  <td width="20%"><a href="members.asp?SF=<% = Server.URLEncode(Request.QueryString("SF")) %>&GID=<% = intGetGroupID %>&SO=LU<% = strQsSID2 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtRegistered %></a><% If Request.QueryString("SO") = "LU" Then Response.Write(" <a href=""members.asp?SF=" & Server.URLEncode(Request.QueryString("SF")) & "&GID=" & intGetGroupID & "&SO=LU&OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %></td>
  <td width="6%" align="center"><a href="members.asp?SF=<% = Server.URLEncode(Request.QueryString("SF")) %>&GID=<% = intGetGroupID %>&SO=PT<% = strQsSID2 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtPoints %></a><% If Request.QueryString("SO") = "PT" Then Response.Write(" <a href=""members.asp?SF=" & Server.URLEncode(Request.QueryString("SF")) & "&GID=" & intGetGroupID & "&SO=PT&OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %></td>
  <td width="20%"><a href="members.asp?SF=<% = Server.URLEncode(Request.QueryString("SF")) %>&GID=<% = intGetGroupID %>&SO=LA<% = strQsSID2 %>" title="<% = strTxtReverseSortOrder %>"><% = strTxtLastActive %></a><% If Request.QueryString("SO") = "LA" Then Response.Write(" <a href=""members.asp?SF=" & Server.URLEncode(Request.QueryString("SF")) & "&GID=" & intGetGroupID & "&SO=LA&OB=" & strSortDirection & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ title=""" & strTxtReverseSortOrder & """ alt=""" & strTxtReverseSortOrder & """ /></a>") %></td><% 
          
          If blnPrivateMessages = True Then 
          		%>
  <td width="6%"  align="center" nowrap><% = strTxtAddBuddy %></td><% 
          
          End If 
          
          %>
  <td width="6%" align="center"><% = strTxtSearch %></td>
 </tr><%
         
         
	'If there are no search results display an error msg
	If lngTotalRecords <= 0 Then
		
		'If there are no search results to display then display the appropriate error message
		Response.Write vbCrLf & "  <tr class=""tableRow""><td colspan=""7"" align=""center""><br />" & strTxtSorryYourSearchFoundNoMembers & "<br /><br /></td></tr>"
	
	
	
	
	'Disply any search results in the forum
	Else


		'Do....While Loop to loop through the recorset to display the forum members
		Do While intCurrentRecord < intEndPosition

			'If there are no member's records left to display then exit loop
			If intCurrentRecord >= lngTotalRecords Then Exit Do
			
			'Initialise varibles
			dtmLastPostDate = ""

			'Read in the profile from the recordset
			lngUserID = CLng(sarryMembers(0,intCurrentRecord))
			strUsername = sarryMembers(1,intCurrentRecord)
			If isDate(sarryMembers(3,intCurrentRecord)) Then dtmLastActiveDate = CDate(sarryMembers(3,intCurrentRecord)) Else dtmLastActiveDate = "2000-01-01 00:00:00"
			If isNumeric(sarryMembers(4,intCurrentRecord)) Then lngNumOfPoints = CLng(sarryMembers(4,intCurrentRecord)) Else lngNumOfPoints = 0
			dtmRegisteredDate = CDate(sarryMembers(5,intCurrentRecord))
			intMemberGroupID = CInt(sarryMembers(2,intCurrentRecord))
			strMemberGroupName = sarryMembers(7,intCurrentRecord)
			intRankStars = CInt(sarryMembers(8,intCurrentRecord))
			strRankCustomStars = sarryMembers(9,intCurrentRecord)
			
			

			'If the users account is not active make there account level guest
			If CBool(sarryMembers(6,intCurrentRecord)) = False Then intMemberGroupID = 0

			'Write the HTML of the Topic descriptions as hyperlinks to the Topic details and message
			%>
 <tr class="<% If (intCurrentRecord MOD 2 = 0 ) Then Response.Write("evenTableRow") Else Response.Write("oddTableRow") %>">
  <td><a href="member_profile.asp?PF=<% = lngUserID & strQsSID2 %>" rel="nofollow"><% = strUsername %></a></td>
  <td class="smText"><% = strMemberGroupName %><br /><img src="<% If strRankCustomStars <> "" Then Response.Write(strRankCustomStars) Else Response.Write(strImagePath & intRankStars & "_star_rating.png") %>" alt="<% = strMemberGroupName %>" title="<% = strMemberGroupName %>" /></td>
  <td class="smText"><% = DateFormat(dtmRegisteredDate) %></td>
  <td align="center"><% = lngNumOfPoints %></td>
  <td class="smText"><% = DateFormat(dtmLastActiveDate) %></td><% 
          	If blnPrivateMessages = True Then %>
  <td align="center"><a href="pm_buddy_list.asp?name=<% = Server.URLEncode(strUsername) %><% = strQsSID2 %>"><img src="<% = strImagePath %>add_buddy.<% = strForumImageType %>" border="0" alt="<% = strTxtAddToBuddyList %>" title="<% = strTxtAddToBuddyList %>" /></a></td><% 
          	End If %>
  <td align="center"><a href="search_form.asp?USR=<% = Server.URLEncode(strUsername) %><% = strQsSID2 %>"><img src="<% = strImagePath %>profile_search.<% = strForumImageType %>" border="0" alt="<% = strTxtSearchForPosts %>&nbsp;<% = strUsername %>" title="<% = strTxtSearchForPosts %>&nbsp;<% = strUsername %>" /></a></td>
 </tr><%
			
			'Move to the next record
			intCurrentRecord = intCurrentRecord + 1
					
		'Loop back round
		Loop
	End If	
			%>
</table>
<%
	
End If


%>
<table class="basicTable" cellspacing="0" cellpadding="4" align="center">
 <tr>
  <td align="right" nowrap>
   <!-- #include file="includes/page_link_inc.asp" -->
  </td>
 </tr>
</table><%


'Reset Server Objects
Call closeDatabase()

%>    
      <div align="center"><br />
<%
 
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