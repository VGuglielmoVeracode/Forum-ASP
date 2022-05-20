<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
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
Dim strHomepage			'Holds the users homepage if they have one
Dim strEmail			'Holds the users e-mail address
Dim blnShowEmail		'Boolean set to true if the user wishes there e-mail address to be shown
Dim lngUserID			'Holds the new users ID number
Dim lngNumOfPosts		'Holds the number of posts the user has made
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
Dim strUsername			'Holds the users username
Dim strSearchBy
Dim strAuthorEmail
Dim strLoginIP
Dim strGroupName
Dim intSelGroupID
Dim strLocation


'Initalise variables
blnShowEmail = False
intGetGroupID = IntC(Request.QueryString("GID"))
intPageSize = 25


'If this is the first time the page is displayed then the members record position is set to page 1
If Request.QueryString("PN") = "" Then
	intRecordPositionPageNum = 1

'Else the page has been displayed before so the members page record postion is set to the Record Position number
Else
	intRecordPositionPageNum = IntC(Request.QueryString("PN"))
End If



'Get the what we are searching in email or username
If NOT Request.QueryString("SB") = "" Then
	strSearchBy = Trim(Mid(Request.QueryString("SB"), 1, 15))
End If

'Get the search critiria for the members to display
If NOT Request.QueryString("SF") = "" Then
	strSearchCriteria = Trim(Mid(Request.QueryString("SF"), 1, 40))
End If

'Get rid of milisous code
strSearchCriteria = formatSQLInput(strSearchCriteria)

'Get the sort critiria
Select Case Request.QueryString("SO")
	Case "PT"
		strSortBy = strDbTable & "Author.No_of_posts "
	Case "LU"
		strSortBy = strDbTable & "Author.Join_date "
	Case "OU"
		strSortBy = strDbTable & "Author.Join_date "
	Case "GP"
		strSortBy = strDbTable & "Group.Name "
	Case "LA"
		strSortBy = strDbTable & "Author.Last_visit "
	Case "EM"
		strSortBy = strDbTable & "Author.Author_email "
	Case "IP"
		strSortBy = strDbTable & "Author.Login_IP "
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

	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.Group_ID, " & strDbTable & "Author.Last_visit, " & strDbTable & "Author.No_of_posts, " & strDbTable & "Author.Join_date, " & strDbTable & "Author.Active, " & strDbTable & "Group.Name, " & strDbTable & "Group.Stars, " & strDbTable & "Group.Custom_stars, " & strDbTable & "Author.Author_email, " & strDbTable & "Author.Login_IP, " & strDbTable & "Author.Location " & _
	"FROM " & strDbTable & "Author " & strDBNoLock & ", " & strDbTable & "Group " & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Author.Group_ID = " & strDbTable & "Group.Group_ID "
		
	
	'If a search criteria
	If NOT strSearchCriteria = "" Then	
		If strSearchBy = "email" Then
			strSQL = strSQL & "AND " & strDbTable & "Author.Author_email LIKE '%" & strSearchCriteria & "%' "
		ElseIf strSearchBy = "IP" Then
			strSQL = strSQL & "AND " & strDbTable & "Author.Login_IP LIKE '%" & strSearchCriteria & "%' "
		ElseIf strSearchBy = "Location" Then
			strSQL = strSQL & "AND " & strDbTable & "Author.Location LIKE '%" & strSearchCriteria & "%' "
		Else
			strSQL = strSQL & "AND " & strDbTable & "Author.Username LIKE '" & strSearchCriteria & "%' "
		End If
	End If
	
	'If a group id to look for
	If NOT intGetGroupID = 0 Then	
	 	strSQL = strSQL & "AND " & strDbTable & "Author.Group_ID ="  & intGetGroupID & " "
	End If
		
	strSQL = strSQL & "ORDER BY " & strSortBy & ";"
	

	'Query the database  
	rsCommon.Open strSQL, adoCon
	
	'If there are records get em from rs
	If NOT rsCommon.EOF Then
		
		'Read in the row from the db using getrows for better performance
		sarryMembers = rsCommon.GetRows()
		
		
		'Count the number of records
		lngTotalRecords = Ubound(sarryMembers,2) + 1
		
		'Count the number of pages for the topics using '\' so that any fraction is omitted 
		lngTotalRecordsPages = lngTotalRecords \ intPageSize
		
		'If there is a remainder or the result is 0 then add 1 to the total num of pages
		If lngTotalRecords Mod intPageSize > 0 OR lngTotalRecordsPages = 0 Then lngTotalRecordsPages = lngTotalRecordsPages + 1
		
		
		'Start position
		intStartPosition = ((intRecordPositionPageNum - 1) * intPageSize)	
		
		'End Position
		intEndPosition = intStartPosition + intPageSize
		
		'Get the start position
		intCurrentRecord = intStartPosition
	End If
	
	
	'Close the recordset as it is no longer needed
	rsCommon.Close

End If

'Page to link to for mutiple page (with querystrings if required)
strLinkPage = "admin_select_members.asp?SF=" & Server.URLEncode(Request.QueryString("SF")) & "&GID=" & intGetGroupID & "&SB=" & strSearchBy & "&"


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Forum Member Adminstration</title>
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

<script  language="JavaScript">
function CheckForm () {

	//Check for a somthing to search for
	if (document.getElementById('frmMemberSearch').SF.value==""){

		msg = "_______________________________________________________________\n\n";
		msg += "The form has not been submitted because there are problem(s) with the form.\n";
		msg += "Please correct the problem(s) and re-submit the form.\n";
		msg += "_______________________________________________________________\n\n";
		msg += "The following field(s) need to be corrected: -\n";

		alert(msg + "\nMember Search\t- Enter a Members Username to search for\n\n");
		document.getElementById('frmMemberSearch').SF.focus();
		return false;
	}

	return true;
}
</script>
<script language="javascript" src="includes/default_javascript_v9.js" type="text/javascript"></script>

<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center"> 
 <h1>Forum Member Administration</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br />
  Click on the members name to administer their forum membership account, <br />
  from where you can, change their details, member group, reset password, suspend, delete, etc. from the Forum.<br />
 <br />
 <form name="frmMemberSearch" method="get" action="admin_select_members.asp" onSubmit="return CheckForm();">
   <table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
     <tr class="tableLedger">
       <td colspan="2" align="left">Member Search </td>
     </tr>
     <tr class="tableRow">
       <td align="left">Member Search :
           <input name="SF" size="30" maxlength="40" value="<% = Server.HTMLEncode(Request.QueryString("SF")) %>" />
           <input type="hidden" name="SID" id="SID" value="<% = strQsSID %>" />
         &nbsp; &nbsp;
           Forum Group : <select name="GID">
		<option value="0" selected="selected">-</option>
            	<%

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
           
           <input type="submit" name="Submit" value="Search" />
           <br />
           Search By: 
           <input name="SB" type="radio" value="user"<% If strSearchBy = "user" OR strSearchBy = "" Then Response.Write(" checked=""checked""") %> />
        Username&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <input name="SB" type="radio" value="email"<% If strSearchBy = "email" Then Response.Write(" checked=""checked""") %> />
        Email Address&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
        <input name="SB" type="radio" value="IP"<% If strSearchBy = "IP" Then Response.Write(" checked=""checked""") %> /> IP Address      
         &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
        <input name="SB" type="radio" value="Location"<% If strSearchBy = "Location" Then Response.Write(" checked=""checked""") %> /> Location       <br />

          
      <br />   <br />
          <a href="admin_select_members.asp?GID=<% = intGetGroupID %><% = strQsSID1 %>">#</a> 
   <a href="admin_select_members.asp?SF=A&GID=<% = intGetGroupID %><% = strQsSID2 %>">A</a> 
   <a href="admin_select_members.asp?SF=B&GID=<% = intGetGroupID %><% = strQsSID2 %>">B</a> 
   <a href="admin_select_members.asp?SF=C&GID=<% = intGetGroupID %><% = strQsSID2 %>">C</a>
   <a href="admin_select_members.asp?SF=D&GID=<% = intGetGroupID %><% = strQsSID2 %>">D</a> 
   <a href="admin_select_members.asp?SF=E&GID=<% = intGetGroupID %><% = strQsSID2 %>">E</a> 
   <a href="admin_select_members.asp?SF=F&GID=<% = intGetGroupID %><% = strQsSID2 %>">F</a>
   <a href="admin_select_members.asp?SF=G&GID=<% = intGetGroupID %><% = strQsSID2 %>">G</a> 
   <a href="admin_select_members.asp?SF=H&GID=<% = intGetGroupID %><% = strQsSID2 %>">H</a> 
   <a href="admin_select_members.asp?SF=I&GID=<% = intGetGroupID %><% = strQsSID2 %>">I</a>
   <a href="admin_select_members.asp?SF=J&GID=<% = intGetGroupID %><% = strQsSID2 %>">J</a> 
   <a href="admin_select_members.asp?SF=K&GID=<% = intGetGroupID %><% = strQsSID2 %>">K</a> 
   <a href="admin_select_members.asp?SF=L&GID=<% = intGetGroupID %><% = strQsSID2 %>">L</a>
   <a href="admin_select_members.asp?SF=M&GID=<% = intGetGroupID %><% = strQsSID2 %>">M</a> 
   <a href="admin_select_members.asp?SF=N&GID=<% = intGetGroupID %><% = strQsSID2 %>">N</a> 
   <a href="admin_select_members.asp?SF=O&GID=<% = intGetGroupID %><% = strQsSID2 %>">O</a>
   <a href="admin_select_members.asp?SF=P&GID=<% = intGetGroupID %><% = strQsSID2 %>">P</a> 
   <a href="admin_select_members.asp?SF=Q&GID=<% = intGetGroupID %><% = strQsSID2 %>">Q</a> 
   <a href="admin_select_members.asp?SF=R&GID=<% = intGetGroupID %><% = strQsSID2 %>">R</a>
   <a href="admin_select_members.asp?SF=S&GID=<% = intGetGroupID %><% = strQsSID2 %>">S</a> 
   <a href="admin_select_members.asp?SF=T&GID=<% = intGetGroupID %><% = strQsSID2 %>">T</a> 
   <a href="admin_select_members.asp?SF=U&GID=<% = intGetGroupID %><% = strQsSID2 %>">U</a>
   <a href="admin_select_members.asp?SF=V&GID=<% = intGetGroupID %><% = strQsSID2 %>">V</a> 
   <a href="admin_select_members.asp?SF=W&GID=<% = intGetGroupID %><% = strQsSID2 %>">W</a> 
   <a href="admin_select_members.asp?SF=X&GID=<% = intGetGroupID %><% = strQsSID2 %>">X</a>
   <a href="admin_select_members.asp?SF=Y&GID=<% = intGetGroupID %><% = strQsSID2 %>">Y</a> 
   <a href="admin_select_members.asp?SF=Z&GID=<% = intGetGroupID %><% = strQsSID2 %>">Z</a>
  
  </td>
     
   
     </tr>
   </table>
 </form>
</div>
<br />
 <table class="basicTable" cellspacing="0" cellpadding="4" align="center">
     <tr>
       <td align="right" nowrap="nowrap"><!-- #include file="includes/page_link_inc.asp" -->
       </td>
     </tr>
   </table>
   <table cellspacing="1" cellpadding="3" class="tableBorder" align="center">
     <tr class="tableLedger">
       <td width="80"><a href="admin_select_members.asp?SF=<% = Server.URLEncode(Request.QueryString("SF")) %>&amp;GID=<% = intGetGroupID %>&amp;SO=UN&SB=<% = strSearchBy & strQsSID2 %>">Username</a> <% If Request.QueryString("SO") = "UN" OR Request.QueryString("SO") = "" Then Response.Write(" <a href=""admin_select_members.asp?SF=" & Server.URLEncode(Request.QueryString("SF")) & "&GID=" & intGetGroupID & "&SO=UN&OB=" & strSortDirection & "&SB=" & strSearchBy & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ alt=""Reverse Sort Order"" /></a>") %></td>
       <td width="110"><a href="admin_select_members.asp?SF=<% = Server.URLEncode(Request.QueryString("SF")) %>&amp;GID=<% = intGetGroupID %>&amp;SO=EM&SB=<% = strSearchBy & strQsSID2 %>">Email Address</a> <% If Request.QueryString("SO") = "EM" OR Request.QueryString("SO") = "" Then Response.Write("<a href=""admin_select_members.asp?SF=" & Server.URLEncode(Request.QueryString("SF")) & "&GID=" & intGetGroupID & "&SO=EM&OB=" & strSortDirection & "&SB=" & strSearchBy & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ alt=""Reverse Sort Order"" /></a>") %></td>
       <td width="80"><a href="admin_select_members.asp?SF=<% = Server.URLEncode(Request.QueryString("SF")) %>&amp;GID=<% = intGetGroupID %>&amp;SO=IP&SB=<% = strSearchBy & strQsSID2 %>">Login IP</a> <% If Request.QueryString("SO") = "IP" Then Response.Write(" <a href=""admin_select_members.asp?SF=" & Server.URLEncode(Request.QueryString("SF")) & "&GID=" & intGetGroupID & "&SO=IP&OB=" & strSortDirection & "&SB=" & strSearchBy & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ alt=""Reverse Sort Order"" /></a>") %></td>
       <td width="60"><a href="admin_select_members.asp?SF=<% = Server.URLEncode(Request.QueryString("SF")) %>&amp;GID=<% = intGetGroupID %>&amp;SO=GP&SB=<% = strSearchBy & strQsSID2 %>">Group</a> <% If Request.QueryString("SO") = "GP" Then Response.Write(" <a href=""admin_select_members.asp?SF=" & Server.URLEncode(Request.QueryString("SF")) & "&GID=" & intGetGroupID & "&SO=GP&OB=" & strSortDirection & "&SB=" & strSearchBy & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ alt=""Reverse Sort Order"" /></a>") %></td>
       <td width="90"><a href="admin_select_members.asp?SF=<% = Server.URLEncode(Request.QueryString("SF")) %>&amp;GID=<% = intGetGroupID %>&amp;SO=LU&SB=<% = strSearchBy & strQsSID2 %>">Registered</a> <% If Request.QueryString("SO") = "LU" Then Response.Write(" <a href=""admin_select_members.asp?SF=" & Server.URLEncode(Request.QueryString("SF")) & "&GID=" & intGetGroupID & "&SO=LU&OB=" & strSortDirection & "&SB=" & strSearchBy & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ alt=""Reverse Sort Order"" /></a>") %></td>
       <td width="64" align="center"><a href="admin_select_members.asp?SF=<% = Server.URLEncode(Request.QueryString("SF")) %>&amp;GID=<% = intGetGroupID %>&amp;SO=PT&SB=<% = strSearchBy & strQsSID2 %>">Posts</a> <% If Request.QueryString("SO") = "PT" Then Response.Write(" <a href=""admin_select_members.asp?SF=" & Server.URLEncode(Request.QueryString("SF")) & "&GID=" & intGetGroupID & "&SO=PT&OB=" & strSortDirection & "&SB=" & strSearchBy & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ alt=""Reverse Sort Order"" /></a>") %></td>
        <td width="90">Location</td>
       <td width="90"><a href="admin_select_members.asp?SF=<% = Server.URLEncode(Request.QueryString("SF")) %>&amp;GID=<% = intGetGroupID %>&amp;SO=LA&SB=<% = strSearchBy & strQsSID2 %>">Last Active</a> <% If Request.QueryString("SO") = "LA" Then Response.Write(" <a href=""admin_select_members.asp?SF=" & Server.URLEncode(Request.QueryString("SF")) & "&GID=" & intGetGroupID & "&SO=LA&OB=" & strSortDirection & "&SB=" & strSearchBy & strQsSID2 & """><img src=""" & strImagePath & strSortDirection & "." & strForumImageType & """ alt=""Reverse Sort Order"" /></a>") %></td>
       <td width="10" align="center">Delete&nbsp;Member</td>
     </tr><%
         
         
	'If there are no search results display an error msg
	If lngTotalRecords <= 0 Then
		
		'If there are no search results to display then display the appropriate error message
		Response.Write vbCrLf & "    <tr class=""tableRow""><td colspan=""7"" align=""center""><br />Your search has found no results<br /><br /></td></tr>"
	
	
	
	
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
			If isDate(sarryMembers(3,intCurrentRecord)) Then dtmLastActiveDate = CDate(sarryMembers(3,intCurrentRecord)) Else dtmLastActiveDate = Now()
			lngNumOfPosts = CLng(sarryMembers(4,intCurrentRecord))
			dtmRegisteredDate = CDate(sarryMembers(5,intCurrentRecord))
			intMemberGroupID = CInt(sarryMembers(2,intCurrentRecord))
			strMemberGroupName = sarryMembers(7,intCurrentRecord)
			intRankStars = CInt(sarryMembers(8,intCurrentRecord))
			strRankCustomStars = sarryMembers(9,intCurrentRecord)
			strAuthorEmail = sarryMembers(10,intCurrentRecord)
			strLoginIP = sarryMembers(11,intCurrentRecord)
			strLocation = sarryMembers(12,intCurrentRecord)
			
			

			'If the users account is not active make there account level guest
			If CBool(sarryMembers(6,intCurrentRecord)) = False Then intMemberGroupID = 0

			'Write the HTML of the Topic descriptions as hyperlinks to the Topic details and message
			%>
     <tr class="<% If (intCurrentRecord MOD 2 = 0 ) Then Response.Write("evenTableRow") Else Response.Write("oddTableRow") %>">
       <td><a href="admin_register.asp?PF=<% = lngUserID & strQsSID2 %>"><% = strUsername %></a></td>
       <td><% = strAuthorEmail %></td>
       <td><% If strLoginIP <> "" Then Response.Write("<a href=""admin_ip_blocking.asp?IP=" & strLoginIP & strQsSID2 & """>" & strLoginIP & "</a> <a href=""https://network-tools.webwiz.net/ip-information.htm?ip=" & Server.URLEncode(strLoginIP) & """ target=""_blank""><img src=""" & strImagePath & "new_window.png"" alt=""IP Information"" title=""IP Information"" /></a>") %></td>
       <td class="smText"><% = strMemberGroupName %><br /><img src="<% If strRankCustomStars <> "" Then Response.Write(strRankCustomStars) Else Response.Write(strImagePath & intRankStars & "_star_rating.png") %>" alt="<% = strMemberGroupName %>" /></td>
       <td class="smText"><% = DateFormat(dtmRegisteredDate) %></td>
       <td  align="center"><% = lngNumOfPosts %></td>
       <td class="smText"><% = strLocation %></td>
       <td class="smText"><% = DateFormat(dtmLastActiveDate) %></td>
       <td align="center"><% If lngUserID > 2 Then %><a href="admin_delete_member.asp?PF=<% = lngUserID & "&XID=" & getSessionItem("KEY") & strQsSID2 %>"><img src="<% = strImagePath %>delete.png" border="0" title="Delete Member" /></a><% End If %></td>
      </tr><%
			
			'Move to the next record
			intCurrentRecord = intCurrentRecord + 1
					
		'Loop back round
		Loop
	End If	
			%>
   </table>
   <table class="basicTable" cellspacing="0" cellpadding="4" align="center">
     <tr>
       <td align="right" nowrap="nowrap"><!-- #include file="includes/page_link_inc.asp" -->
       </td>
     </tr>
   </table>
   <%


'Reset Server Objects
Call closeDatabase()

%>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->