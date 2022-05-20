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
If intGroupID <> 2 Then


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



	'Read the members profile if it is there profile or if admin/moderator
	If (lngProfileNum = lngLoggedInUserID) OR (strMemberProfileView = "members" AND (blnAdmin OR blnModerator)) OR (strMemberProfileView = "admins-moderators" AND (blnAdmin OR blnModerator)) OR (strMemberProfileView = "admins" AND blnAdmin) Then
	
	
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
	
	
	
		End If
	
		'Reset Server Objects
		rsCommon.Close
		
		'Clean up email link
		If strEmail <> "" Then
			strEmail = formatInput(strEmail)
		End If
		
	End If
End If





'If no avatar then use generic
If strAvatar = "" OR blnAvatar = false Then strAvatar = "avatars/blank_avatar.jpg"
	
'If SSL is enabled, but the avatar is not HTTPS then remove avatar
If strSslEnabled = "Enabled" AND InStr(strAvatar, "http://") Then strAvatar = "avatars/blank_avatar.jpg"



'Kick user if nothing to display
If strUsername = "" Then
	
	Response.End
	
	'Clean up (done down here are session data may need to be saved)
	Call closeDatabase()
End If

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strTxtMemberData %></title>
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
<style>
body {
    font: normal 14px Verdana, Arial, sans-serif;
}
	
</style>
</head>
<body>
<h1><% = strTxtMemberData & " - " & strUsername %></h1>
<h2><% = strTxtProfile %></h2>	
<br />	<img src="<% = strAvatar %>" id="avatar" alt="<strong><% = strTxtAvatar %>" align="left" onError="document.getElementById('avatar').src='avatars/blank_avatar.jpg'">
<br /><br /><br /><br />
<br />	<strong><% = strTxtUsername %>:</strong> <% = strUsername %>
<br />	<strong><% = strTxtMemberTitle %>:</strong> <% = strMemberTitle %>
<br />	<strong><% = strTxtGroup %>:</strong> <% = strGroupName %> <img src="<% 
        If strRankCustomStars <> "" Then Response.Write(strRankCustomStars) Else Response.Write(strImagePath & intRankStars & "_star_rating.png") 
	Response.Write(""" alt=""" & strGroupName & """ title=""" & strGroupName & """>") %>
<br />	<strong><% = strTxtLadderGroup %>:</strong> <% If strLadderName = "" OR isNull(strLadderName) Then Response.Write(strTxtNone) Else Response.Write(strLadderName) %>
<br />	<strong><% = strTxtAccountStatus %>:</strong> <% 
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
     	 	%>
<br />	<strong><% = strTxtJoined %>:</strong> <% = DateFormat(dtmJoined) & " " & strTxtAt & " " & TimeFormat(dtmJoined) %>
<br />	<strong><% = strTxtLastVisit %>:</strong> <% 
 
	'last Login date/time   	
	If isDate(dtmLastVisit) Then Response.Write(DateFormat(dtmLastVisit) & " " & strTxtAt & " " & TimeFormat(dtmLastVisit)) 
%>
<br />	<strong><% = strTxtLastVisit & " " & strTxtIPAddress %>:</strong>
<%	
	'Last login IP
	Response.Write(strLastLoginIP)
		
	'Read in country code
	strCountryCode = IpCountryLookup(strLastLoginIP, strInstallID)
		
	'If we have a country code display it
	If NOT strCountryCode = "-" Then
		
		'Loop through ISO country array to display all the coutries in a drop down
		For intIsoLoop = 1 to UBound(saryISOCountryCode,2)
				
			If saryISOCountryCode(0,intIsoLoop) = strCountryCode Then Response.Write("(" & saryISOCountryCode(1,intIsoLoop) & ") ")
	
		Next
	End If
	
%>

<br />	<%

	'If Web Wiz NewsPad integration is enabled show if teh user has subscribed
	If blnWebWizNewsPad Then
%>
    <strong><% = strTxtNewsletterSubscription %>:</strong> 
     <% If blnNewsletter Then Response.Write("No") Else Response.Write("Yes") %>
    <%
	End If

%>
<br />	<strong><% = strTxtPoints %>:</strong> <% = lngNumOfPoints %>
<br />	<strong><% = strTxtPosts %>:</strong> <% = lngNumOfPosts %> <% If lngNumOfPosts > 0 AND DateDiff("d", dtmJoined, Now()) > 0 Then Response.Write(" [" & FormatNumber(lngNumOfPosts / DateDiff("d", dtmJoined, Now()), 2) & " " & strTxtPostsPerDay) & "]" %>

    <%

	'If Answer posts are on display the number of answers from the member
	If NOT strAnswerPosts = "Off" Then
	
%>
<br /> <strong><% = strAnswerPostsWording %>:</strong> <% = lngNumOfAnwsers %> 
    <%
    
	End If

	'If thanking members is enabled display the number of times the member has been thanked
	If blnPostThanks Then
%>
<br /> <strong><% = strTxtThanked %>:</strong> <% = lngNumOfThanked %> 
    <%

	End If
       
   	
  	'If custom field 1 is required
	If strCustRegItemName1 <> "" AND (blnViewCustRegItemName1 OR (blnAdmin OR  blnModerator)) Then 	
%>
<br /> <strong><% = strCustRegItemName1 %>:</strong> <% If strCustItem1 <> "" Then Response.Write(strCustItem1)  %> 
    <%
	End If

	'If custom field 2 is required
	If strCustRegItemName2 <> "" AND (blnViewCustRegItemName2 OR (blnAdmin OR  blnModerator)) Then 	
%>
<br /> <strong><% = strCustRegItemName2 %>:</strong> <% If strCustItem2 <> "" Then Response.Write(strCustItem2)  %> 
    <%
	End If

	'If custom field 3 is required
	If strCustRegItemName3 <> "" AND (blnViewCustRegItemName3 OR (blnAdmin OR  blnModerator)) Then 	
%>
<br /> <strong><% = strCustRegItemName3 %>:</strong> <% If strCustItem3 <> "" Then Response.Write(strCustItem3)  %> 
    <%
	End If

%>
<br /> <strong><% = strTxtRealName %>:</strong> <% If strRealName <> "" Then Response.Write(strRealName)  %> 
    
<br /> <strong><% = strTxtGender %>:</strong> <% If strGender <> "" Then Response.Write(strGender)  %> 
    
<br /><strong><% = strTxtDateOfBirth %>:</strong> 
     <% 
         
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
        
        %> 
    
<br /><strong><% = strTxtAge %>:</strong> 
     <% If intAge > 0 Then Response.Write(intAge) Else Response.Write(strTxtUnknown) %> 
    
<br /><strong><% = strTxtLocation %>:</strong> 
     <% If strLocation = "" Or isNull(strLocation) Then Response.Write(strTxtNotGiven) Else Response.Write(strLocation) %> 
    <%
    
    	'If homepages are enabled
    	If blnHomePage Then
%>
<br /><strong><% = strTxtHomepage %>:</strong> <% = formatInput(strHomepage) %>
    <%
    
	End If

%>
<br /><strong><% = strTxtOccupation %>:</strong> 
     <% If strOccupation = "" OR IsNull(strOccupation) Then Response.Write(strTxtNotGiven) Else Response.Write(strOccupation) %> 
<br /><strong><% = strTxtInterests %>:</strong> 
     <% If strInterests = "" OR IsNull(strInterests) Then Response.Write(strTxtNotGiven) Else Response.Write(strInterests) %> 
<br /> <strong><% = strTxtEmailAddress %>:</strong> <% = strEmail %> 
<br /><strong><% = strTxtFacebook %>:</strong>  <% = formatInput(strFacebookUsername) %>    
<br /><strong><% = strTxtTwitter %>:</strong> <% = formatInput(strTwitterUsername) %>    
<br /><strong><% = strTxtLinkedIn %>:</strong> <% = formatInput(strTwitterUsername) %>
<br /><strong><% = strTxtMSNMessenger %>:</strong> <% = formatInput(strMSNAddress) %>  
<br /><strong><% = strTxtSkypeName %>:</strong> <% = formatInput(strSkypeName) %>  
<br /><strong><% = strTxtYahooMessenger %>:</strong> <% = formatInput(strYahooAddress) %>   
<br /><strong><% = strTxtAIMAddress %>:</strong> <% = formatInput(strAIMAddress) %>   
<br /><strong><% = strTxtICQNumber %>:</strong> <% If not strICQNum = "" Then Response.write(formatInput(strICQNum)) %>
<%   
   
 
  
  	'If there are notes on the user display them
  	If blnSignatures AND (strSignature <> "" AND blnAccSuspended = False) Then
  		
  		%>
<br /><br /> <strong><% = strTxtSignature %>;  </strong>
<br />
<% Response.Write(formatSignature(strSignature)) %> 
 <%
  
	End If




strSQL = "SELECT " & strDbTable & "Thread.IP_addr, " & strDbTable & "Thread.Message_date " & _
        "FROM " & strDbTable & "Thread " & _
        "WHERE (" & strDbTable & "Thread.Author_ID = " & lngProfileNum & ");"
	

'Query the database
rsCommon.Open strSQL, adoCon

'If there are records returned display them
If NOT rsCommon.EOF Then

%>
<br /><br />
<h2><% = strTxtIPAddresses %></h2>
<% = strTxtBelowYouCanFindIPAddressesRecordedInWhenPosting %>
<br />
<%

	'If records are returned loop through and display the IP addresses
	Do While not rsCommon.EOF 

		Response.Write ("<br />") 
		Response.Write (rsCommon("IP_addr"))
		Response.Write (" - ") 
		Response.Write (DateFormat(rsCommon("Message_date")) & " " & strTxtAt & " " & TimeFormat(rsCommon("Message_date")))

	 	'Move to the next record in the recordset 
	    	rsCommon.MoveNext 
	Loop

End If

'Close rs
rsCommon.Close

'Clean up (done down here are session data may need to be saved)
Call closeDatabase()
%>

</body>
</html>