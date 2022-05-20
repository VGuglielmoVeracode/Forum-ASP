<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
<!--#include file="language_files/pm_language_file_inc.asp" -->
<!--#include file="includes/ISO_country_list_inc.asp" -->
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

Dim lngProfileNum		'Holds the profile number of the user we are getting the profile for
Dim strUsername			'Holds the users username
Dim intUsersGroupID		'Holds the users group ID
Dim strEmail			'Holds the new users e-mail address
Dim blnShowEmail		'Boolean set to true if the user wishes there e-mail address to be shown
Dim strLocation			'Holds the new users location
Dim strHomepage			'Holds the new users homepage if they have one
Dim strAvatar			'Holds the avatar image
Dim strICQNum			'Holds the users ICQ Number
Dim strAIMAddress		'Holds the users AIM address
Dim strMSNAddress		'Holds the users MSN address
Dim strYahooAddress		'Holds the users Yahoo Address
Dim strOccupation		'Holds the users Occupation
Dim strInterests		'Holds the users Interests
Dim dtmJoined			'Holds the joined date
Dim lngNumOfPosts		'Holds the number of posts the user has made
Dim lngNumOfPoints		'Holds the number of points the user has 
Dim dtmDateOfBirth		'Holds the users Date Of Birth
Dim dtmLastVisit		'Holds the date the user last came to the forum
Dim strGroupName		'Holds the group name
Dim intRankStars 		'Holds the rank stars
Dim strRankCustomStars		'Holds the custom stars image if there is one
Dim blnProfileReturned		'Boolean set to false if the user's profile is not found in the database
Dim blnGuestUser		'Set to True if the user is a guest or not logged in
Dim blnActive			'Set to true of the users account is active
Dim strRealName			'Holds the persons real name
Dim strMemberTitle		'Holds the members title
Dim blnIsUserOnline		'Set to true if the user is online
Dim strPassword			'Holds the password
Dim strSignature		'Holds the signature
Dim strSkypeName		'Holds the users Skype Name
Dim intArrayPass		'Holds the array loop
Dim intAge			'Holds the age of the user
Dim strAdminNotes		'Holds the admin notes on the user
Dim blnAccSuspended		'Holds if the user account is suspended
Dim strOnlineLocation		'Holds the users location in the forum
Dim strOnlineURL		'Holds the users online location URL
Dim blnNewsletter		'set to true if user is signed up to newsletter
Dim strGender			'Holds the users gender
Dim strLadderName		'Ladder group name
Dim strLastLoginIP		'Holds the login/registration IP for user
Dim intOnlineForumID		'Holds the forum id for active user
Dim strCustItem1		'Custom item 1
Dim strCustItem2		'Custom item 2
Dim strCustItem3		'Custom item 3
Dim lngNumOfAnwsers		'Number of asnwsers
Dim lngNumOfThanked		'Number of thanks
Dim strFacebookUsername		'Holds the facebook username
Dim strTwitterUsername		'Holds the twitter username
Dim strLinkedInUsername		'Holds the linkedin username
Dim blnViewProfile 		'set to true if the user has permission to view the profile
Dim intIsoLoop
Dim strCountryCode






'Initalise variables
blnSslEnabledPage = True
blnProfileReturned = True
blnViewProfile = False
blnGuestUser = False
blnShowEmail = False
blnModerator = False
blnIsUserOnline = False
lngNumOfPosts = 0
lngNumOfPoints = 0
lngNumOfAnwsers = 0
lngNumOfThanked = 0




'If the user is using a banned IP address then don't let the view a profile
If bannedIP()  Then blnBanned = True

'Read in the profile number to get the details on
lngProfileNum = LngC(Request.QueryString("PF"))


'If who can view profiles is not set then set it to members
If strMemberProfileView = "" Then strMemberProfileView = "members"


'If the user has logged in then the Logged In User ID number will be more than 0
If strMemberProfileView = "everyone" OR intGroupID <> 2 Then


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


	'Read the members profile
	If (strMemberProfileView = "everyone") OR (strMemberProfileView = "members") OR (strMemberProfileView = "admins-moderators" AND (blnAdmin OR blnModerator)) OR (strMemberProfileView = "admins" AND blnAdmin) Then
	
		'Set the veriable that the user is allowed to view the profile to true
		blnViewProfile = True
	
		'Initalise the strSQL variable with an SQL statement to query the database
		strSQL = "" & _
		"SELECT " & strDbTable & "Author.*, " & strDbTable & "Group.Name, " & strDbTable & "Group.Stars, " & strDbTable & "Group.Custom_stars, " & strDbTable & "LadderGroup.Ladder_Name, " & strDbTable & "Group.Signatures " & _
		"FROM (" & strDbTable & "Author INNER JOIN " & strDbTable & "Group ON " & strDbTable & "Author.Group_ID = " & strDbTable & "Group.Group_ID) " & _
			"LEFT JOIN " & strDbTable & "LadderGroup ON " & strDbTable & "Group.Ladder_ID = " & strDbTable & "LadderGroup.Ladder_ID " & _
		"WHERE " & strDbTable & "Author.Author_ID = " & lngProfileNum
	
		'Query the database
		rsCommon.Open strSQL, adoCon
	
		'Read in the details if a profile is returned
		If NOT rsCommon.EOF Then
			
			'Read in the new user's profile from the recordset
			strUsername = rsCommon("Username")
			strRealName = rsCommon("Real_name")
			strCustItem1 = rsCommon("Custom1")
			strCustItem2 = rsCommon("Custom2")
			strCustItem3 = rsCommon("Custom3")
			intUsersGroupID = CInt(rsCommon("Group_ID"))
			strEmail = rsCommon("Author_email")
			strGender = rsCommon("Gender")
			blnShowEmail = CBool(rsCommon("Show_email"))
			strHomepage = rsCommon("Homepage")
			strLocation = rsCommon("Location")
			strAvatar = rsCommon("Avatar")
			strMemberTitle = rsCommon("Avatar_title")
			strFacebookUsername = rsCommon("Facebook")
			strTwitterUsername = rsCommon("Twitter")
			strLinkedInUsername = rsCommon("LinkedIn")
			strICQNum = rsCommon("ICQ")
			strAIMAddress = rsCommon("AIM")
			strMSNAddress = rsCommon("MSN")
			strYahooAddress = rsCommon("Yahoo")
			strOccupation = rsCommon("Occupation")
			strInterests = rsCommon("Interests")
			If isDate(rsCommon("DOB")) Then dtmDateOfBirth = CDate(rsCommon("DOB"))
			dtmJoined = CDate(rsCommon("Join_date"))
			lngNumOfPosts = CLng(rsCommon("No_of_posts"))
			If isNull(rsCommon("Points")) Then lngNumOfPoints = 0 Else lngNumOfPoints = CLng(rsCommon("Points"))
			If isNull(rsCommon("Answered")) Then lngNumOfAnwsers = 0 Else lngNumOfAnwsers = CLng(rsCommon("Answered"))
			If isNull(rsCommon("Thanked")) Then lngNumOfThanked = 0 Else lngNumOfThanked = CLng(rsCommon("Thanked"))
			dtmLastVisit = rsCommon("Last_visit")
			strGroupName = rsCommon("Name")
			intRankStars = CInt(rsCommon("Stars"))
			strRankCustomStars = rsCommon("Custom_stars")
			blnActive = CBool(rsCommon("Active"))
			strSignature = rsCommon("Signature")
			strSkypeName = rsCommon("Skype")
			strAdminNotes = rsCommon("Info")
			blnAccSuspended = CBool(rsCommon("Banned"))
			If isNull(rsCommon("Newsletter")) = False Then blnNewsletter = CBool(rsCommon("Newsletter")) Else blnNewsletter = False
			strLadderName = rsCommon("Ladder_Name")
			If isNull(rsCommon("Login_IP")) Then  strLastLoginIP = "Unknown" Else strLastLoginIP = rsCommon("Login_IP")
			'If signatures are not allowed for this group update the global blnSignatures to be fales for this page so the signature is not displayed
			If CBool(rsCommon("Signatures")) = False Then blnSignatures = False
	
		'Else no profile is returned so set an error variable
		Else
			blnProfileReturned = False
	
		End If
	
		'Reset Server Objects
		rsCommon.Close
		
		'Clean up email link
		If strEmail <> "" Then
			strEmail = formatInput(strEmail)
		End If
		
		
		'If active user is enabled then get the users location
		If blnActiveUsers Then
			
			'Call active users function
			saryActiveUsers = activeUsers(strTxtViewing & " " & strTxtProfile, "&#8216;" & strUsername & "&#8217; " & strTxtProfile, "member_profile.asp?PF=" & lngProfileNum, 0)
			
			'Get the users online status
			For intArrayPass = 1 To UBound(saryActiveUsers, 2)
				If saryActiveUsers(1, intArrayPass) = lngProfileNum Then 
					blnIsUserOnline = True
					strOnlineLocation = saryActiveUsers(6, intArrayPass)
					strOnlineURL = saryActiveUsers(7, intArrayPass)
					intOnlineForumID = IntC(saryActiveUsers(9, intArrayPass))
				End If
			Next
			
			'Check the permissions to see if the user has permission to see the topic subject for active users
			If intOnlineForumID > 0 Then
				'Check permissions
				Call forumPermissions(intOnlineForumID, intGroupID)
		
				'If the user doesn't have read permissions then remove the Location URL (which includes topic subjects)
				If blnRead = False Then strOnlineURL = ""
			End If
		End If
	End If




'Else the user is not logged in
Else
	'Set the Guest User boolean to true as the user must be a guest
	blnGuestUser = True
End If



'Set the active user location
If blnGuestUser OR blnViewProfile = False Then
	'If active users is enabled update the active users application array
	If blnActiveUsers Then
		'Call active users function
		saryActiveUsers = activeUsers(strTxtViewing & " " & strTxtProfile & " [" & strTxtAccessDenied & "]", "", "", 0)
	End If
End If




'If no avatar then use generic
If strAvatar = "" OR blnAvatar = false Then strAvatar = "avatars/blank_avatar.jpg"
	
'If SSL is enabled, but the avatar is not HTTPS then remove avatar
If strSslEnabled = "Enabled" AND InStr(strAvatar, "http://") Then strAvatar = "avatars/blank_avatar.jpg"
 

'Set bread crumb trail
strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtProfile

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strTxtProfile %></title>
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
  <td align="left"><h1><% = strTxtProfile %></h1></td>
 </tr>
</table><%


'If there is a problem then display error message
If blnProfileReturned = False OR blnGuestUser OR blnActiveMember = False OR blnBanned OR blnViewProfile = False Then

%><table class="errorTable" cellspacing="1" cellpadding="3" align="center">
 <tr>
  <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong><%
  	
	'If no profile can be found then display the appropriate message
	If blnProfileReturned = False Then
	
		Response.Write ("<br /><br />" & strTxtNoUserProfileFound)
	
	'If the user is a guest then tell them they must register or login before they can view other users profiles
	ElseIf blnGuestUser OR blnActiveMember = False OR blnBanned Then
	
		Response.Write ("<br /><br />" & strTxtRegisteredToViewProfile)
		
		'If mem suspended display message
		If blnBanned Then
			Response.Write("<br /><br />" & strTxtForumMemberSuspended)
		
		'Else account not yet active
		ElseIf blnActiveMember = false Then
			
			Response.Write("<br /><br />" & strTxtForumMembershipNotAct)
			If blnMemberApprove = False Then Response.Write("<br /><br />" & strTxtToActivateYourForumMem)
		
			'If admin activation is enabled let the user know
			If blnMemberApprove Then
				Response.Write("<br /><br />" & strTxtYouAdminNeedsToActivateYourMembership)
			'If email is on then place a re-send activation email link
			ElseIf blnEmailActivation AND blnLoggedInUserEmail Then 
				Response.Write("<br /><br /><a href=""javascript:winOpener('resend_email_activation.asp" & strQsSID1 & "','actMail',1,1,475,300)"">" & strTxtResendActivationEmail & "</a>")
			End If
		End If
	
	'If user does not have permission to view profile
	ElseIf blnViewProfile = False Then
		Response.Write ("<br /><br />" & strTxtYouDoNotHavePermissionToViewPorfiles)
	
	End If
	
	Response.Write("<br /><br /><a href=""javascript:history.back(1)"">" & strTxtReturnToDiscussionForum & "</a>")
%> 
  </td>
 </tr>
</table><%



	'If guest user let the user login
	If blnGuestUser Then
	
		%><!--#include file="includes/login_form_inc.asp" --><%
		
	End If



'Else display user profile
Else



	'If the user is an admin or a moderator give them the chance to edit the profile unless it's the main admin account of the guest account
	If blnAdmin OR (blnModerator AND blnModeratorProfileEdit) Then

%>
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="right"><a href="spam_cleaner.asp?PF=<% = lngProfileNum & strQsSID2 %>"><% = strTxtSpamCleaner %></a> | <a href="member_control_panel.asp?PF=<% = lngProfileNum %>&M=A<% = strQsSID2 %>"><% = strTxtEditMembersSettings %></a></td>
 </tr>
</table><%

	'Else no admin options so just have a line break
	Else
		Response.Write("<br />")
	End If

%>
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center">
 <tr class="tableLedger">
  <td width="50%"><% = strTxtProfile & " - " & strUsername %></td>
  <td width="50%"><% = strTxtActiveStats %></td>
 </tr>
 <tr class="tableRow">
  <td valign="top">
   <table width="100%" border="0" cellspacing="1" cellpadding="2">
    <tr>
     <td width="30%"><img src="<% = strAvatar %>" id="avatar" alt="<% = strTxtAvatar %>" align="left" onError="document.getElementById('avatar').src='avatars/blank_avatar.jpg'"></td>
     <td width="80%">&nbsp;</td>
    </tr>
    <tr>
     <td><% = strTxtUsername %>:</td>
     <td><% = strUsername %></td>
    </tr><%

	'If there is a member title display it
	If strMemberTitle <> "" Then
	
	%>
    <tr>
     <td><% = strTxtMemberTitle %>:</td>
     <td><% = strMemberTitle %></td>
    </tr><%

	End If

%>
    <tr>
     <td><% = strTxtGroup %>:</td>
     <td><% = strGroupName %> <img src="<% 
        If strRankCustomStars <> "" Then Response.Write(strRankCustomStars) Else Response.Write(strImagePath & intRankStars & "_star_rating.png") 
	Response.Write(""" alt=""" & strGroupName & """ title=""" & strGroupName & """>") %></td>
    </tr>
    <tr>
     <td><% = strTxtLadderGroup %>:</td>
     <td><% If strLadderName = "" OR isNull(strLadderName) Then Response.Write(strTxtNone) Else Response.Write(strLadderName) %>
      </td>
    </tr>
   </table>
  </td>
  <td valign="top">
   <table width="100%" border="0" cellspacing="1" cellpadding="2">
    <tr>
     <td width="30%"><% = strTxtAccountStatus %>:</td>
     <td width="80%"><% 
     	'Display account status
     	If blnAccSuspended Then
     		Response.Write(strTxtSuspended)
     		'If suspended allow their account to be submitted as a spammer account
     		'If (blnAdmin OR (blnModerator AND blnModViewIpAddresses)) AND blnStopForumSpam AND strStopForumSpamApiKey <> "" Then Response.Write("  - <a href=""http://www.stopforumspam.com/add?username=" & strUsername & "&email=" & strEmail & "&ip_addr=" & strLastLoginIP & "&api_key=" & strStopForumSpamApiKey & """ target=""_blank"">" & strSubmitAsSpammer & "</a>")
     		
     	ElseIf blnActive Then 
     	 	Response.Write(strTxtActive) 
     	Else 
     	 	Response.Write(strTxtNotActive)
     	End If 	
     	 	%></td>
    </tr>
    <tr>
     <td><% = strTxtJoined %>:</td>
     <td><% = DateFormat(dtmJoined) & " " & strTxtAt & " " & TimeFormat(dtmJoined) %></td>
    </tr>
    <tr>
     <td><% = strTxtLastVisit %>:</td>
     <td><% 
 
	'last Login date/time   	
	If isDate(dtmLastVisit) Then Response.Write(DateFormat(dtmLastVisit) & " " & strTxtAt & " " & TimeFormat(dtmLastVisit)) 
	
	'Last login IP
	If (blnAdmin OR (blnModerator AND blnModViewIpAddresses)) AND (NOT strLastLoginIP = "Unknown") Then 
		Response.Write(" - " & strTxtIP & ": <a href=""javascript:winOpener('pop_up_IP_blocking.asp?IP=" & strLastLoginIP & strQsSID2 & "','ip',1,1,500,475)"">" & strLastLoginIP & "</a> ")
		
		'Read in country code
		strCountryCode = IpCountryLookup(strLastLoginIP, strInstallID)
		
		'If we have a country code display it
		If NOT strCountryCode = "-" Then
		
			'Loop through ISO country array to display all the coutries in a drop down
			For intIsoLoop = 1 to UBound(saryISOCountryCode,2)
				
				If saryISOCountryCode(0,intIsoLoop) = strCountryCode Then Response.Write("(" & saryISOCountryCode(1,intIsoLoop) & ") ")
	
			Next
		End If
		 
		Response.Write("<a href=""https://network-tools.webwiz.net/ip-information.htm?ip=" & Server.URLEncode(strLastLoginIP) & """ target=""_blank""><img src=""" & strImagePath & "new_window.png"" alt=""" & strTxtIP & " " & strTxtInformation & """ title=""" & strTxtIP & " " & strTxtInformation & """ /></a>") 	
	
	End If
%></td>
    </tr><%

	'If Web Wiz NewsPad integration is enabled show if teh user has subscribed
	If blnWebWizNewsPad Then
%>
    <tr>
     <td><% = strTxtNewsletterSubscription %>:</td>
     <td><% If blnNewsletter Then Response.Write("<img src=""" & strImagePath & "yes.png"" alt=""" & strTxtYes & """ title=""" & strTxtYes & """ />") Else Response.Write("<img src=""" & strImagePath & "no.png"" alt=""" & strTxtNo & """ title=""" & strTxtNo & """ />") %></td>
    </tr><%
	End If

%>
    <tr>
     <td><% = strTxtPoints %>:</td>
     <td><% = lngNumOfPoints %></td>
    </tr>
    <tr>
     <td><% = strTxtPosts %>:</td>
     <td><% = lngNumOfPosts %> <% If lngNumOfPosts > 0 AND DateDiff("d", dtmJoined, Now()) > 0 Then Response.Write(" [" & FormatNumber(lngNumOfPosts / DateDiff("d", dtmJoined, Now()), 2) & " " & strTxtPostsPerDay) & "]" %></td>
    </tr>
    <tr>
     <td><% = strTxtFindPosts %>:</td>
     <td><img src="<% = strImagePath %>profile_search.png" border="0" alt="<% = strTxtSearchForPosts %>&nbsp;<% = strUsername %>" title="<% = strTxtSearchForPosts %>&nbsp;<% = strUsername %>" /> <a href="search_form.asp?USR=<% = Server.URLEncode(strUsername) & strQsSID2 %>" title="<% = strTxtSearchForPosts %>&nbsp;<% = strUsername %>"><% = strTxtFindMembersPosts %></a></td>
    </tr><%

	'If Answer posts are on display the number of answers from the member
	If NOT strAnswerPosts = "Off" Then
	
%>
    <tr>
     <td><% = strAnswerPostsWording %>:</td>
     <td><% = lngNumOfAnwsers %></td>
    </tr><%
    
	End If

	'If thanking members is enabled display the number of times the member has been thanked
	If blnPostThanks Then
%>
    <tr>
     <td><% = strTxtThanked %>:</td>
     <td><% = lngNumOfThanked %></td>
    </tr><%

	End If
        
        'If active users are enabled display if they are online or not
        If blnActiveUsers Then
        	%>
    <tr>
     <td><% = strTxtStatus %>:</td>
     <td><% If blnIsUserOnline Then Response.Write("<img src=""" & strImagePath & "online.png"" alt=""" & strTxtOnLine2 & """ title=""" & strTxtOnLine2 & """ /> " & strTxtOnLine2) Else Response.Write("<img src=""" & strImagePath & "offline.png"" alt=""" & strTxtOffLine & """ title=""" & strTxtOffLine & """ /> " & strTxtOffLine)%></td>
    </tr><%
    		'If the user is online display their location in the forum
    		If blnIsUserOnline Then
    			%>
    <tr>
     <td valign="top"><% = strTxtOnLine2 & " " & strTxtLocation %>:</td>
     <td><% = strOnlineLocation %><% If strOnlineLocation <> "" AND strOnlineURL <> "" Then Response.Write("<br />") %><% = strOnlineURL %></td>
    </tr><%
    		
    		End If

	End If

%>
   </table>
  </td>
 </tr>
 <tr class="tableLedger">
  <td><% = strTxtInformation %></td>
  <td><% = strTxtCommunicate %></td>
 </tr>
 <tr class="tableRow">
  <td valign="top">
   <table width="100%" border="0" cellspacing="1" cellpadding="2"><%
   	
  	'If custom field 1 is required
	If strCustRegItemName1 <> "" AND (blnViewCustRegItemName1 OR (blnAdmin OR  blnModerator)) Then 	
%>
    <tr>
     <td width="30%"><% = strCustRegItemName1 %>:</td>
     <td width="80%"><% If strCustItem1 <> "" Then Response.Write(strCustItem1) Else Response.Write(strTxtNotGiven) %></td>
    </tr><%
	End If

	'If custom field 2 is required
	If strCustRegItemName2 <> "" AND (blnViewCustRegItemName2 OR (blnAdmin OR  blnModerator)) Then 	
%>
    <tr>
     <td width="30%"><% = strCustRegItemName2 %>:</td>
     <td width="80%"><% If strCustItem2 <> "" Then Response.Write(strCustItem2) Else Response.Write(strTxtNotGiven) %></td>
    </tr><%
	End If

	'If custom field 3 is required
	If strCustRegItemName3 <> "" AND (blnViewCustRegItemName3 OR (blnAdmin OR  blnModerator)) Then 	
%>
    <tr>
     <td width="30%"><% = strCustRegItemName3 %>:</td>
     <td width="80%"><% If strCustItem3 <> "" Then Response.Write(strCustItem3) Else Response.Write(strTxtNotGiven) %></td>
    </tr><%
	End If

%>
    <tr>
     <td width="30%"><% = strTxtRealName %>:</td>
     <td width="80%"><% If strRealName <> "" Then Response.Write(strRealName) Else Response.Write(strTxtNotGiven) %></td>
    </tr>
    <tr>
     <td width="30%"><% = strTxtGender %>:</td>
     <td width="80%"><% If strGender <> "" Then Response.Write(strGender) Else Response.Write(strTxtNotGiven) %></td>
    </tr>
    <tr>
     <td><% = strTxtDateOfBirth %>:</td>
     <td><% 
         
         'If there is a Date of Birth display it
         If isDate(dtmDateOfBirth) Then 
         	
         	'Calculate the age (use months / 12 as counting years is not accurate) (use FIX to get the whole number)
		intAge = Fix(DateDiff("m", dtmDateOfBirth, now())/12)
         	
         	'Display the persons Date of Birth
         	Response.Write(stdDateFormat(dtmDateOfBirth, False)) 
         	
         Else 	
         	'Display that a Date of Birth was not given
         	Response.Write(strTxtNotGiven) 
         	
        End If
        
        %></td>
    </tr>
    <tr>
     <td><% = strTxtAge %>:</td>
     <td><% If intAge > 0 Then Response.Write(intAge) Else Response.Write(strTxtUnknown) %></td>
    </tr>
    <tr>
     <td><% = strTxtLocation %>:</td>
     <td><% If strLocation = "" Or isNull(strLocation) Then Response.Write(strTxtNotGiven) Else Response.Write(strLocation) %></td>
    </tr><%
    
    	'If homepages are enabled
    	If blnHomePage Then
%>
    <tr>
     <td><% = strTxtHomepage %>:</td>
     <td><img src="<% = strImagePath %>profile_homepage.png" border="0" alt="<% = strTxtHomepage %>" title="<% = strTxtVisitMembersHomepage %>" /> <% If strHomepage = "" OR IsNull(strHomepage) OR blnAccSuspended Then Response.Write(strTxtNotGiven) Else Response.Write("<a href=""" & formatInput(strHomepage) & """ target=""_blank"" title=""" & formatInput(strHomepage) & """>" & strTxtHomepage & "</a>") %></td>
    </tr><%
    
	End If

%>
    <tr>
     <td><% = strTxtOccupation %>:</td>
     <td><% If strOccupation = "" OR IsNull(strOccupation) Then Response.Write(strTxtNotGiven) Else Response.Write(strOccupation) %></td>
    </tr>
    <tr>
     <td><% = strTxtInterests %>:</td>
     <td><% If strInterests = "" OR IsNull(strInterests) Then Response.Write(strTxtNotGiven) Else Response.Write(strInterests) %></td>
    </tr>
   </table>
  </td>
  <td valign="top">
   <table width="100%" border="0" cellspacing="1" cellpadding="2"><%

    	'If the private messager is on show PM link
    	If blnPrivateMessages Then

%>
    <tr>
     <td><% = strTxtPrivateMessage %>:</td>
     <td><img src="<% = strImagePath %>profile_pm.png" border="0" alt="<% = strTxtSendPrivateMessage %>" title="<% = strTxtSendPrivateMessage %>" /> <% Response.Write("<a href=""pm_new_message_form.asp?name=" & Server.URLEncode(Replace(strUsername, "'", "\'",  1, -1, 1)) & strQsSID2 & """>" & strTxtSendPrivateMessage & "</a>") %></td>
    </tr>
    <tr>
     <td><% = strTxtBuddyList %>:</td>
     <td><img src="<% = strImagePath %>add_buddy.png" border="0" alt="<% = strTxtAddToBuddyList %>" title="<% = strTxtAddToBuddyList %>" /> <a href="pm_buddy_list.asp?name=<% = Server.URLEncode(Replace(strUsername, "'", "\'",  1, -1, 1)) & strQsSID2 %>"><% = strTxtAddToBuddyList %></a></td>
    </tr><%

	End If

%>
    <tr>
     <td width="30%"><% = strTxtEmailAddress %>:</td>
     <td width="80%"><img src="<% = strImagePath %>profile_email.png" border="0" alt="<% = strTxtSendEmail %>" title="<% = strTxtSendEmail %>" /> <%

         'If member account is suspend don't show email
        If blnAccSuspended AND blnAdmin = False AND strEmail <> "" Then
        	Response.Write(strTxtPrivate)
        
        'If the user has choosen not to display there e-mail then this field will show private
	ElseIf blnShowEmail = False AND blnAdmin = False AND strEmail <> "" Then
        	Response.Write(strTxtPrivate)

        'If no email address given then display not given
        ElseIf strEmail = "" OR isNull(strEmail) Then
            	Response.Write(strTxtNotGiven)

        'If email address is shown and the email messenger of the forum is enabled show link button
        ElseIf blnEmailMessenger Then

        	Response.Write("<a href=""email_messenger.asp?SEID=" & lngProfileNum & strQsSID2 & """>" & strTxtSendEmail & "</a>")

        'Else the user allows there email address to be shown so show there email address
        Else
            	Response.Write("<a href=""mailto:" & strEmail & """>" & strEmail & "</a>")
        End If
        
        'If admin or moderator display the members email address
        If blnAdmin OR (blnModerator AND blnModeratorProfileEdit) Then Response.Write(" (" & strEmail & ")")


    %></td>
    
    <tr>
     <td><% = strTxtFacebook %>:</td>
     <td><img src="<% = strImagePath %>profile_facebook.gif" border="0" alt="<% = strTxtFacebook %>" title="<% = strTxtFacebook %>" /> <% If strFacebookUsername <> "" Then Response.Write("<a href=""http://www.facebook.com/" & formatInput(strFacebookUsername) & """ target=""_blank"">" & strTxtFacebook &"</a>") Else Response.Write(strTxtNotGiven) %></td>
    </tr>
    <tr>
     <td><% = strTxtTwitter %>:</td>
     <td><img src="<% = strImagePath %>profile_twitter.gif" border="0" alt="<% = strTxtTwitter %>" title="<% = strTxtTwitter %>" /> <% If strTwitterUsername <> "" Then Response.Write("<a href=""https://twitter.com/#!/" & formatInput(strTwitterUsername) & """ target=""_blank"">" & strTwitterUsername &"</a>") Else Response.Write(strTxtNotGiven) %></td>
    </tr>
    <tr>
     <td><% = strTxtLinkedIn %>:</td>
     <td><img src="<% = strImagePath %>profile_linkedin.gif" border="0" alt="<% = strTxtLinkedIn %>" title="<% = strTxtLinkedIn %>" /> <% If strLinkedInUsername <> "" Then Response.Write("<a href=""http://www.linkedin.com/in/" & formatInput(strLinkedInUsername) & """ target=""_blank"">" & strTxtLinkedIn &"</a>") Else Response.Write(strTxtNotGiven) %></td>
    </tr>
    
    <tr>
     <td><% = strTxtMSNMessenger %>:</td>
     <td><img src="<% = strImagePath %>profile_msn.png" border="0" alt="<% = strTxtMSNMessenger %>" title="<% = strTxtMSNMessenger %>" /> <% If strMSNAddress <> "" Then Response.Write(formatInput(strMSNAddress)) Else Response.Write(strTxtNotGiven) %></td>
    </tr>
    <tr>
     <td><% = strTxtSkypeName %>:</td>
     <td><img src="<% = strImagePath %>profile_skype.png" border="0" alt="<% = strTxtSkypeName %>" title="<% = strTxtSkypeName %>" /> <% If strSkypeName <> "" Then Response.Write("<a href=""skype:" & formatInput(strSkypeName) & "?call"">" & strTxtSkypeName & "</a>") Else Response.Write(strTxtNotGiven) %></td>
    </tr>
    <tr>
     <td><% = strTxtYahooMessenger %>:</td>
     <td><img src="<% = strImagePath %>profile_yim.png" border="0" alt="<% = strTxtYahooMessenger %>" title="<% = strTxtYahooMessenger %>" /> <% If strYahooAddress <> "" Then Response.Write("<a href=""http://edit.yahoo.com/config/send_webmesg?.target=" & formatInput(strYahooAddress) & "&amp;.src=pg"">" & strTxtYahooMessenger &"</a>") Else Response.Write(strTxtNotGiven) %></td>
    </tr>
    <tr>
     <td><% = strTxtAIMAddress %>:</td>
     <td><img src="<% = strImagePath %>profile_aol.png" border="0" alt="<% = strTxtAIMAddress %>" title="<% = strTxtAIMAddress %>" /> <% If strAIMAddress <> "" Then Response.Write("<a href=""aim:goim?screenname=" & formatInput(strAIMAddress) & "&message=Hello+Are+you+there?"">" & strTxtAIMAddress & "</a>") Else Response.Write(strTxtNotGiven) %></td>
    </tr>
    <tr>
     <td><% = strTxtICQNumber %>:</td>
     <td><img src="<% = strImagePath %>profile_icq.png" border="0" alt="<% = strTxtICQNumber %>" title="<% = strTxtICQNumber %>" /> <% If strICQNum <> "" Then Response.Write("<a href=""http://wwp.icq.com/scripts/search.dll?to=" & formatInput(strICQNum) & """>" & strTxtICQNumber & "</a>") Else Response.Write(strTxtNotGiven) %></td>
    </tr>
   
   </table>
  </td>
 </tr><%
  
  	'If there is a signature display it
  	If strAdminNotes <> "" AND (blnAdmin OR blnModerator) Then
  		
  		%>
 <tr class="tableLedger">
  <td colspan="2"><% = strTxtAdminNotes %></td>
 </tr>
 <tr class="tableRow">
  <td colspan="2"><% 
  
  		'Put in line breaks
  		strAdminNotes = Replace(strAdminNotes, vbCrLf, "<br />", 1, -1, 1)
  
  		'Display admin notes
  		Response.Write(strAdminNotes) 
 %></td>
 </tr><%
  
	End If

  
  	'If there are notes on the user display them
  	If blnSignatures AND (strSignature <> "" AND blnAccSuspended = False) Then
  		
  		%>
 <tr class="tableLedger">
  <td colspan="2"><% = strTxtSignature %></td>
 </tr>
 <tr class="tableRow">
  <td colspan="2"><% Response.Write(formatSignature(strSignature)) %></td>
 </tr><%
  
	End If

%>
</table><%


End If


'Clean up (done down here are session data may need to be saved)
Call closeDatabase()
%>
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
%></div>
<!-- #include file="includes/footer.asp" -->