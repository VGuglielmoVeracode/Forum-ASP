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





'If active users is off redirect back to the homepage
If blnActiveUsers = False Then
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("default.asp" & strQsSID1)
End If




'Make sure this page is not cached
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "No-Store"



'Dimension variables
Dim lngUserID			'Holds the active users ID
Dim strUsername			'Holds the active users username
Dim strForumName 		'Holds the forum name
Dim intGuestNumber		'Holds the Guest Number
Dim dtmLoggedIn			'Holds the date/time the user logged in
Dim dtmLastActive		'Holds the date/time the user was last active
Dim intActiveUsers		'Holds the number of active users
Dim intActiveGuests		'Holds the number of active guests
Dim intActiveMembers		'Holds the number of logged in active members
Dim intActiveSearchRobots
Dim intForumColourNumber	'Holds the number to calculate the table row colour	
Dim intArrayPass		'Loop counter
Dim strOSbrowser		'Holds the OS and browser of the user
Dim strLocation			'Holds the users location
Dim strURL			'Holds the URL to the users location
Dim blnHideActiveUser		'Holds if the user wants to be hidden
Dim intAnonymousMembers		'Holds the number of intAnonymous members online
Dim strActiveUserIP
Dim strLastActive

'Initilise variables
blnSslEnabledPage = True
intActiveMembers = 0
intActiveGuests = 0
intActiveUsers = 0
intGuestNumber = 0
intForumColourNumber = 0
intAnonymousMembers = 0
intActiveSearchRobots = 0


'Call active users function
saryActiveUsers = activeUsers(strTxtActiveUsers, "", "", 0)


'Sort the active users array
Call SortActiveUsersList(saryActiveUsers)


'If the user has logged in then the Logged In User ID number will be more than 0
If intGroupID <> 2 Then


	'See if the user is a in a moderator group for any forum
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
End If


'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtActiveForumUsers


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strTxtActiveForumUsers %></title>
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
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
 
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />  	 
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% = strTxtActiveForumUsers %></h1></td>
 </tr>
</table>
<br /><%




'Get the number of active users
'Get the active users online
For intArrayPass = 1 To UBound(saryActiveUsers, 2)

	'If this is a search robot then increment serach robot
	If Instr(saryActiveUsers(5, intArrayPass), "Search Robot") Then
		
		intActiveSearchRobots = intActiveSearchRobots + 1
		
	'If this is a guest user then increment the number of active guests veriable	
	ElseIf saryActiveUsers(1, intArrayPass) = 2 Then 
			
		intActiveGuests = intActiveGuests + 1
		
	'Else if the user is Anonymous increment the Anonymous count
	ElseIf CBool(saryActiveUsers(8, intArrayPass)) Then	
			
		intAnonymousMembers = intAnonymousMembers + 1
	End If	
Next 



'Calculate the number of members online and total people online
intActiveUsers = UBound(saryActiveUsers, 2)

'Calculate the members online by using the total - Guests - Anonymous Members
intActiveMembers = intActiveUsers - intActiveGuests - intAnonymousMembers - intActiveSearchRobots

%>    
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left">
   <div style="float:left;"><% = (strTxtInTotalThereAre & " " & intActiveUsers & " " & strTxtActiveUsers & " " & strTxtOnLine  & ", " & intActiveMembers & " " & strTxtMembers & ", " & intAnonymousMembers & " " & strTxtAnonymousMembers & ", " & intActiveGuests & " " & strTxtGuests & ", " & intActiveSearchRobots & " " & strTxtSearchEngines) %></div>
   <div style="float:right;"><a href="active_users.asp<% = strQsSID1 %>"><img src="<% = strImagePath %>refresh.<% = strForumImageType %>" alt="<% = strTxtRefreshPage %>" title="<% = strTxtRefreshPage %>" /></a>&nbsp;</div>
  </td>
 </tr>
</table>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td width="17%"><% = strTxtUsername %></td>
  <td width="14%"><% = strTxtLastActive %></td>
  <td width="10%"><% = strTxtActive %></td>
  <td width="20%" align="left"><% = strTxtOS & "/" & strTxtBrowser %></td><%

'If this is an admin or moderator then display the IP address of the user
If blnAdmin OR (blnModerator AND blnModViewIpAddresses) Then Response.Write(vbCrLf & "  <td width=""10%"" align=""left"">" & strTxtIP & "</td>")
%>
  <td width="35%" align="left"><% = strTxtLocation %></td>
 </tr><%  


        		
'display the active users
For intArrayPass = 1 To UBound(saryActiveUsers, 2)

	intForumColourNumber = intForumColourNumber + 1
	
	'Array dimension lookup table
	' 0 = IP
	' 1 = Author ID
	' 2 = Username
	' 3 = Login Time
	' 4 = Last Active Time
	' 5 = OS/Browser
	' 6 = Location Page Name
	' 7 = URL
	' 8 = Hids user details

	'Read in the details from the rs
	strActiveUserIP = saryActiveUsers(0, intArrayPass)
	lngUserID = saryActiveUsers(1, intArrayPass)
	strUsername = saryActiveUsers(2, intArrayPass)
	dtmLoggedIn = saryActiveUsers(3, intArrayPass)
	dtmLastActive = saryActiveUsers(4, intArrayPass)
	strOSbrowser = saryActiveUsers(5, intArrayPass)
	strLocation = saryActiveUsers(6, intArrayPass)
	strURL = saryActiveUsers(7, intArrayPass)
	blnHideActiveUser = BoolC(saryActiveUsers(8, intArrayPass))
	intForumID = IntC(saryActiveUsers(9, intArrayPass))

	'Check the permissions to see if the user has permission to see the topic subject
	If intForumID > 0 Then
		'Check permissions
		Call forumPermissions(intForumID, intGroupID)

		'If the user doesn't have read permissions then remove the Location URL (which includes topic subjects)
		If blnRead = False Then strURL = ""
	End If
	
	'Setup the last active date
	strLastActive = DateFormat(dtmLastActive)
	strLastActive = Replace(strLastActive, "<strong>", "")
	strLastActive = Replace(strLastActive, "</strong>", "")
	
	'Replace Search Robot with Search Engine
	strOSbrowser = Replace(strOSbrowser, "Search Robot", strTxtSearchEngine)
	
			
	'Write the HTML of the Topic descriptions as hyperlinks to the Topic details and message
%>
 <tr class="<% If (intForumColourNumber MOD 2 = 0 ) Then Response.Write("evenTableRow") Else Response.Write("oddTableRow") %>"> 
  <td><% 
          
         'If the user is a Guest then display them as a Guest
         If lngUserID = 2 Then
         
         	'Add 1 to the Guest number
         	intGuestNumber = intGuestNumber + 1
         	
         	'Display the User as Guest
         	Response.Write(strTxtGuest & " "& intGuestNumber)
         
         'If the user wants to hide there ID then do so (unless this is an admin or moderator)
         ElseIf blnHideActiveUser AND blnAdmin = False AND blnModerator = False Then
         	
         	'Display the user as an annoy
         	Response.Write(strTxtAnnoymous)
         
         'Else display the users name
         Else 
         
         %><a href="member_profile.asp?PF=<% = lngUserID %><% = strQsSID2 %>" rel="nofollow"><% = strUsername %></a><% 
        
        End If
        
        %></td>
  <td nowrap><% Response.Write(strLastActive & " " & strTxtAt & "&nbsp;" & TimeFormat(dtmLastActive))  %></td>
  <td><% = DateDiff("n", dtmLoggedIn, dtmLastActive) %>&nbsp;<% = strTxtMinutes %></td>
  <td nowrap><% = strOSbrowser %></td><%
	
	'If admin or moderator display the IP address of the user
	If blnAdmin OR (blnModerator AND blnModViewIpAddresses) Then Response.Write(vbCrLf & "  <td nowrap><a href=""javascript:winOpener('pop_up_IP_blocking.asp?IP=" & strActiveUserIP & strQsSID2 & "','ip',1,1,500,475)"">" & strActiveUserIP & "</a> <a href=""https://network-tools.webwiz.net/ip-information.htm?ip=" & Server.URLEncode(strActiveUserIP) & """ target=""_blank""><img src=""" & strImagePath & "new_window.png"" alt=""" & strTxtIP & " " & strTxtInformation & """ title=""" & strTxtIP & " " & strTxtInformation & """ /></td>")

%>
  <td nowrap><% = strLocation %><% If strLocation <> "" AND strURL <> "" Then Response.Write("<br />") %><% = strURL %></td>
 </tr><%
		
	   		
Next
	
'Clean up
Call closeDatabase()

%>
</table>
<div align="center">
    <br /><span class="smText"><% = strTxtDataBasedOnActiveUsersInTheLastXMinutes %></span><br /><br /><% 
    
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