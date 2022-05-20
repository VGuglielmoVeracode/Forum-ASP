<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_send_mail.asp" -->
<!--#include file="functions/functions_edit_post.asp" -->
<!--#include file="functions/functions_format_post.asp" -->
<!--#include file="functions/functions_date_time_format.asp" -->
<!--#include file="includes/emoticons_inc.asp" -->
<!--#include file="language_files/admin_language_file_inc.asp" -->
<!--#include file="functions/functions_hash1way.asp" -->
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
Dim strUsername                 'Holds the users username
Dim strPassword                 'Holds the new users password
Dim strUserCode                 'Holds the unique user code for the user
Dim strEmail                    'Holds the new users e-mail address
Dim intUsersGroupID             'Holds the users group ID
Dim blnShowEmail                'Boolean set to true if the user wishes there e-mail address to be shown
Dim strLocation                 'Holds the new users location
Dim strHomepage                 'Holds the new users homepage if they have one
Dim strAvatar                   'Holds the avatar image
Dim strCheckUsername            'Holds the usernames from the database recordset to check against the new users requested username
Dim blnAutoLogin                'Boolean set to true if the user wants auto login trured on
Dim strImageFileExtension       'holds the file extension
Dim blnAccountReactivate        'Set to true if the users account needs to be reactivated
Dim blnSentEmail                'Set to true if the e-mail has been sent
Dim strEmailBody                'Holds the body of the welcome message e-mail
Dim strSubject                  'Holds the subject of the e-mail
Dim strSignature                'Holds the signature
Dim strICQNum                   'Holds the users ICQ Number
Dim strAIMAddress               'Holds the users AIM address
Dim strMSNAddress               'Holds the users MSN address
Dim strYahooAddress             'Holds the users Yahoo Address
Dim strOccupation               'Holds the users Occupation
Dim strInterests                'Holds the users Interests
Dim dtmDateOfBirth              'Holds the users Date Of Birth
Dim blnPMNotify                 'Set to true if the user want email notification of PM's
Dim strSmutWord                 'Holds the smut word to give better performance so we don't need to keep grabbing it form the recordset
Dim strSmutWordReplace          'Holds the smut word to be replaced with
Dim strMode                     'Holds the mode of the page
Dim blnEmailOK                  'Set to true if e-mail is not already in the database
Dim blnUsernameOK               'Set to true if the username requested does not already exsist
Dim intForumStartingGroup       'Holds the forum starting group ID number
Dim strSalt                     'Holds the salt value for the password
Dim strEncryptedPassword         'Holds the encrypted password
Dim blnPasswordChange           'Holds if the password is changed or not
Dim blnEmailBlocked             'set to true if the email address is blocked
Dim strCheckEmailAddress        'Holds the email address to be checked
Dim lngUserProfileID            'Holds the users ID of the profile to get
Dim blnAdminMode                'Set to true if admin mode is enabled to update other members profiles
Dim blnUserActive               'Set to true if the users membership is active
Dim lngPosts                    'Holds the number of posts the user has made
Dim intDOBYear			'Holds the year of birth
Dim intDOBMonth			'Holds the month of birth
Dim intDOBDay			'Holds the day of birth
Dim strRealName			'Holds the persons real name
Dim strMemberTitle		'Holds the members title
Dim dtmServerTime		'Holds the current server time
Dim lngLoopCounter		'Holds the generic loop counter for page
Dim intUpdatePartNumber		'If an update holds which part to update
Dim strConfirmPassword		'Holds the users old password
Dim blnConfirmPassOK		'Set to false if the conformed pass is not OK
Dim strSkypeName		'Holds the users Skype Name
Dim strSessionKey			'Form ID
Dim blnSuspended		'Holds if user is suspened
Dim strAdminNotes		'Holds admin/modertor info/notes about the user
Dim blnNewsletter		'Set to true if newsletters are selected
Dim strGender			'Holds the users gender
Dim strTempUsername		'Holds a temp username for the user
Dim strTempEmail		'Holds temp email address
Dim blnValidEmail		'Set to false if email is invalid
Dim lngMemberPoints		'Holds the number of points the user has
Dim blnPasswordComplexityOK	'Set if password is complex enough
Dim objRegExp			'used for searches
Dim strCustItem1		'Custom item 1
Dim strCustItem2		'Custom item 2
Dim strCustItem3		'Custom item 3
Dim strFacebookUsername		'Holds the facebook username
Dim strTwitterUsername		'Holds the twitter username
Dim strLinkedInUsername		'Holds the linkedin username
Dim strFormKey
Dim strUserNameFormName
Dim strEmailFormName
Dim strPasswordFormName
Dim blnStopForumSpamFound
Dim blnRealNameOK
Dim blnLocationOK
Dim blnCustRegItemName1OK
Dim blnCustRegItemName2OK
Dim blnCustRegItemName3OK
Dim blnRequiredFieldsValid 
Dim blnSecurityCodeOK
Dim blnIpCountryBanned
Dim strForumName


'Initalise variables
blnSslEnabledPage = True
blnUsernameOK = True
blnEmailOK = True
blnShowEmail = False
blnAutoLogin = True
blnAccountReactivate = False
blnWYSIWYGEditor = True
blnAttachSignature = True
blnPasswordChange = False
blnEmailBlocked = False
blnAdminMode = False
lngUserProfileID = lngLoggedInUserID
blnConfirmPassOK = true
blnNewsletter = False
blnValidEmail = True
blnPasswordComplexityOK = True
strDateFormat = saryDateTimeData(1,0)
blnStopForumSpamFound = False
blnRealNameOK = True
blnLocationOK = True
blnCustRegItemName1OK = True
blnCustRegItemName2OK = True
blnCustRegItemName3OK = True
blnRequiredFieldsValid = True
blnSecurityCodeOK = True
blnIpCountryBanned = False

'Default to short registration form for mobile users
If blnMobileBrowser Then blnLongRegForm = False
	
	
'Make sure that blank cuistom fields are not set to be required, or will throw an error
If strCustRegItemName1 = "" Then blnReqCustRegItemName1 = False
If strCustRegItemName2 = "" Then blnReqCustRegItemName2 = False
If strCustRegItemName3 = "" Then blnReqCustRegItemName3 = False
	


'******************************************
'***	   If banned IP kick user	***
'******************************************	

'If the user is user is using a banned IP redirect to an error page
If bannedIP() Then
	
	'Clean up
	Call closeDatabase()
	
	'Redirect
	Response.Redirect("insufficient_permission.asp?M=IP" & strQsSID3)
End If



'******************************************
'***	     Read in page setup		***
'******************************************

'read in the forum ID number
If isNumeric(Request.QueryString("FID")) Then
	intForumID = IntC(Request.QueryString("FID"))
Else
	intForumID = 0
End If

'Read in the mode of the page
strMode = Trim(Mid(Request.Form("mode"), 1, 7))

'Also see if the admin mode is enabled
If Request("M") = "A" Then blnAdminMode = True

'Check which page part we are displaying and updating if not all
If Request("FPN") Then
	intUpdatePartNumber = IntC(Request("FPN"))
Else
	intUpdatePartNumber = 0
End If



'Get the session ID
strSessionKey = getSessionItem("KEY")
strFormKey = getSessionItem("IDX")



'******************************************
'***  See if this is a new registration	***
'******************************************

'If this is a new registration check the user has accepted the terms of the forum
'Redirect if not been through the registration process
If Request.Form(strSessionKey & "Reg") <> strFormKey AND strMode = "reg" Then

        'Clean up
        Call closeDatabase()

        'Redirect
        Response.Redirect("forum_terms.asp?M=reg&FID=" & intForumID & strQsSID3)
End If




'******************************************
'***	If new reg		  ***
'******************************************

If strMode = "new-reg" OR strMode = "reg" Then 
	
	'If the registration is closed redirect
	If blnRegistrationSuspeneded Then 
		
		'Clean up
	        Call closeDatabase()
	
	        'Redirect
	        Response.Redirect("forum_terms.asp?M=reg&FID=" & intForumID & strQsSID3)
	End If
End If

'If new reg then create a new session key
If strMode = "reg" Then	
	
	'Create new session ID for form ID in case using the one from the registration rules
	strSessionKey = LCase(hexValue(12))
	Call saveSessionItem("KEY", strSessionKey)
	
End If




'Check the user is not registered already and just hitting back on their browser
If (strMode = "new-reg" OR strMode = "reg") AND intGroupID <> 2 Then strMode = ""



'******************************************
'***  Check permision to view page	***
'******************************************

'If the user his not activated their mem
If blnActiveMember = False OR blnBanned Then

        'clean up before redirecting
        Call closeDatabase()

        'redirect to insufficient permissions page
        Response.Redirect("insufficient_permission.asp?M=ACT" & strQsSID3)
End If

'If the user has not logged in or not a new registration then redirect them to the insufficient permissions page
If (intGroupID = 2) AND NOT (strMode = "reg" OR strMode = "new-reg") Then

        'clean up before redirecting
        Call closeDatabase()

        'redirect to insufficient permissions page
        Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If




'********************************************
'***  Check and setup page for admin mode ***
'********************************************

'If the admin mode is enabled see if the user is an admin or moderator
If blnAdminMode Then

        'First see if the user is in a moderator group for any forum
        If blnAdmin = False AND blnModeratorProfileEdit Then

        	'Initalise the strSQL variable with an SQL statement to query the database
		strSQL = "SELECT " & strDbTable & "Permissions.Moderate " & _
	        "FROM " & strDbTable & "Permissions" & strDBNoLock & " " & _
	        "WHERE (" & strDbTable & "Permissions.Group_ID = " & intGroupID & " OR " & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & ") AND  " & strDbTable & "Permissions.Moderate=" & strDBTrue & ";"


                'Query the database
                rsCommon.Open strSQL, adoCon

                'If a record is returned then the user is a moderator in one of the forums
                If NOT rsCommon.EOF Then
                	 blnModerator = True
               'Else this guy is not a moderator
                Else
                	blnModerator = False
                	blnAdminMode = False
                End If

                'Clean up
                rsCommon.Close
        End If


        'Get the profile ID to edit
        lngUserProfileID = LngC(Request("PF"))

        'Turn off email activation if it is enabled as it's not required for an admin edit of a profile
        blnEmailActivation = False


        'If the user is not permitted in to use admin mode send 'em away
        If (blnAdmin = False AND blnModerator = False) Then

                'clean up before redirecting
                Call closeDatabase()

                'redirect to insufficient permissions page
                Response.Redirect("insufficient_permission.asp?FID=" & intForumID & strQsSID3)
        End If
End If





'******************************************
'***    Update or create new member	***
'******************************************

'If the Profile has already been edited then update the Profile
If strMode = "update" OR strMode = "new-reg" Then


	'******************************************
	'***	  Check the session ID		***
	'******************************************

	Call checkFormID(Request.Form("xformID"))
	
	
	'******************************************
	'***	  Check security code		***
	'******************************************

	If strMode = "new-reg" AND blnRegistrationCAPTCHA Then
		'Set the security code OK variable to false
		 If LCase(getSessionItem("SCS")) <> LCase(Request.Form("securityCode")) OR getSessionItem("SCS") = "" Then blnSecurityCodeOK = False
	End If

	'Distroy session variable
	Call saveSessionItem("SCS", "")
	
	
	
	'Recreate encrypted form fields
	strUserNameFormName = "N" & HashEncode("Username" & strFormKey)
	strEmailFormName = "E" & HashEncode("Email" & strFormKey)
	strPasswordFormName = "P" & HashEncode("Password" & strFormKey)
	strFormKey = HashEncode("IDX" & strFormKey)
	
	
	
	
	'******************************************
	'***	  Check the form key		***
	'******************************************
	
	If Request.Form(strSessionKey) <> strFormKey Then
		
		'clean up before redirecting
                Call closeDatabase()

                'redirect to insufficient permissions page
                Response.Redirect("insufficient_permission.asp?M=sID&FID=" & intForumID & strQsSID3)
        End If
        
        'Distroy session variable
        Call saveSessionItem("IDX", "")
	


	'******************************************
	'***  Read in member details from form	***
	'******************************************

        'Read in the users details from the form
        strUsername = Trim(Mid(Request.Form(strUserNameFormName), 1, 20))



        'If part number = 0 (all) or part 1 (reg details) then run this code
        If intUpdatePartNumber = 0 OR intUpdatePartNumber = 1 Then

	        strPassword = Trim(Mid(Request.Form(strPasswordFormName & "1"), 1, 20))
	        strConfirmPassword = Trim(Mid(Request.Form("oldPass"), 1, 20))
	        strEmail = Trim(Mid(Request.Form(strEmailFormName), 1, 60))
	        
	       
	        'Check for a valid email address is enetered
	        If NOT strEmail = "" Then
		        
		        'Check the email address is OK
		        strEmail = emailAddressValidation(strEmail)
		        
		        'If there is no email left beceuase it is not valid then display an error to the user
		        If strEmail = "" Then blnValidEmail = False
		 End If
		 
		 'If there is no email address and email activation is enabled then set the valid email boolen to false
		 If strEmail = "" AND blnEmailActivation Then blnValidEmail = False
       End If



        'If part number = 0 (all) or part 2 (profile details) then run this code
        If intUpdatePartNumber = 0 OR intUpdatePartNumber = 2 Then
        	
        	'Read in custom fields
        	If strCustRegItemName1 <> "" Then strCustItem1 = Trim(Mid(Request.Form("cust1"), 1, 27))
        	If strCustRegItemName2 <> "" Then strCustItem2 = Trim(Mid(Request.Form("cust2"), 1, 27))
        	If strCustRegItemName3 <> "" Then strCustItem3 = Trim(Mid(Request.Form("cust3"), 1, 27))
        	
		'Read in profile details
	        strRealName = Trim(Mid(Request.Form("realName"), 1, 27))
	        strLocation = Trim(Mid(Request.Form("location"), 1, 27))
	        
	        
	        strGender = Trim(Mid(Request.Form("gender"), 1, 10))
	        If blnHomePage Then strHomepage = Trim(Mid(Request.Form("homepage"), 1, 48))
	        If blnSignatures Then
	        	strSignature = Mid(Request.Form("signature"), 1, 210)
	        	blnAttachSignature = BoolC(Request.Form("attachSig"))
	        End If
	        'Check that the ICQ number is a number before reading it in
	        If isNumeric(Request.Form("ICQ")) Then strICQNum = Trim(Mid(Request.Form("ICQ"), 1, 15))
	        strFacebookUsername = Trim(Mid(Request.Form("Facebook"), 1, 60))
	        strTwitterUsername = Trim(Mid(Request.Form("Twitter"), 1, 60))
		strLinkedInUsername = Trim(Mid(Request.Form("LinkedIn"), 1, 60))	
	        strAIMAddress = Trim(Mid(Request.Form("AIM"), 1, 60))
	        strMSNAddress = Trim(Mid(Request.Form("MSN"), 1, 60))
	        strYahooAddress = Trim(Mid(Request.Form("Yahoo"), 1, 60))
	        strSkypeName = Trim(Mid(Request.Form("Skype"), 1, 30))
	        strOccupation = Mid(Request.Form("occupation"), 1, 40)
	        strInterests = Mid(Request.Form("interests"), 1, 130)
	        'Check the date of birth is a date before entering it
	        If Request.Form("DOBday") <> 0 AND Request.Form("DOBmonth") <> 0 AND Request.Form("DOByear") <> 0 Then
	        	dtmDateOfBirth = internationalDateTime(DateSerial(Request.Form("DOByear"), Request.Form("DOBmonth"), Request.Form("DOBday")))
		End If
		
		
		'Validation checks
		If blnRealNameReq AND strRealName = "" Then
			blnRealNameOK = False
			blnRequiredFieldsValid = False
		End If
		
		'Validation checks
		If blnLocationReq AND strLocation = "" Then
			blnLocationOK = False
			blnRequiredFieldsValid = False
		End If
		
		'Validation checks
		If blnReqCustRegItemName1 AND strCustItem1 = "" Then
			blnCustRegItemName1OK = False
			blnRequiredFieldsValid = False
		End If
		
		'Validation checks
		If blnReqCustRegItemName2 AND strCustItem2 = "" Then
			blnCustRegItemName2OK = False
			blnRequiredFieldsValid = False
		End If
		
		'Validation checks
		If blnReqCustRegItemName3 AND strCustItem3 = "" Then
			blnCustRegItemName3OK = False
			blnRequiredFieldsValid = False
		End If
		
		
	End If

	'If part number = 0 (all) or part 3 (forum preferences) then run this code
        If intUpdatePartNumber = 0 OR intUpdatePartNumber = 3 Then

	        If blnWebWizNewsPad Then blnNewsletter = BoolC(Request.Form("newsletter"))
	        blnShowEmail = BoolC(Request.Form("emailShow"))
	        blnPMNotify = BoolC(Request.Form("pmNotify"))
	        blnAutoLogin = BoolC(Request.Form("Login"))
	        strDateFormat = Trim(Mid(Request.Form("dateFormat"), 1, 10))
	        strTimeOffSet = Trim(Mid(Request.Form("serverOffSet"), 1, 1))
	        intTimeOffSet = IntC(Request.Form("serverOffSetHours"))
	        blnReplyNotify = BoolC(Request.Form("replyNotify"))
	        blnWYSIWYGEditor = BoolC(Request.Form("ieEditor"))
	End If



        'If we are in admin mode read in some extras (unless the admin or guest accounts)
        If blnAdminMode AND blnDemoMode = False Then
        	If lngUserProfileID > 2 Then blnUserActive = BoolC(Request.Form("active"))
        	If lngUserProfileID > 2 Then intUsersGroupID = IntC(Request.Form("group"))
        	If isNumeric(Request.Form("posts")) Then lngPosts = LngC(Request.Form("posts"))
        	If isNumeric(Request.Form("points")) Then lngMemberPoints = LngC(Request.Form("points"))
        	strMemberTitle = Trim(Mid(Request.Form("memTitle"), 1, 40))
        	blnSuspended = BoolC(Request.Form("banned"))
        	strAdminNotes = Trim(Mid(removeAllTags(Request.Form("notes")), 1, 255))
        End If



        '******************************************
	'***     Read in the avatar		***
	'******************************************

        'If avatars are enabled then read in selected avatar
        If blnAvatar = True AND (intUpdatePartNumber = 0 OR intUpdatePartNumber = 2) Then

                strAvatar = Trim(Mid(Request.Form("txtAvatar"), 1, 95))

                'If the avatar text box is empty then read in the avatar from the list box
                If strAvatar = "http://" OR strAvatar = "https://" OR strAvatar = "" Then strAvatar = Trim(Request.Form("SelectAvatar"))

                'If there is no new avatar selected then get the old one if there is one
                If strAvatar = "" Then strAvatar = Request.Form("oldAvatar")

                'If the avatar is the blank image then the user doesn't want one
                If strAvatar = strImagePath & "blank.gif" Then strAvatar = ""
        Else
                strAvatar = ""
        End If




        '******************************************
	'***     Clean up member details	***
	'******************************************

        'Clean up user input

        'If part number = 0 (all) or part 2 (profile details) then run this code
        If intUpdatePartNumber = 0 OR intUpdatePartNumber = 2 Then
        	
        	'Custom info
        	If strCustRegItemName1 <> "" Then
        		strCustItem1 = removeAllTags(strCustItem1)
	        	strCustItem1 = formatInput(strCustItem1)
	        End If
	        If strCustRegItemName2 <> "" Then
        		strCustItem2 = removeAllTags(strCustItem2)
	        	strCustItem2 = formatInput(strCustItem2)
	        End If
	        If strCustRegItemName3 <> "" Then
        		strCustItem3 = removeAllTags(strCustItem3)
	        	strCustItem3 = formatInput(strCustItem3)
	        End If
        	
        	'Profile info
	        strRealName = removeAllTags(strRealName)
	        strRealName = formatInput(strRealName)
	        strGender = removeAllTags(strGender)
	        strGender = formatInput(strGender)
	        strLocation = removeAllTags(strLocation)
	        strLocation = formatInput(strLocation)
	        strOccupation = removeAllTags(strOccupation)
	        strOccupation = formatInput(strOccupation)
	        strInterests = removeAllTags(strInterests)
	        strInterests = formatInput(strInterests)
	        

	        'Call the function to format the signature
	        strSignature = FormatPost(strSignature)

	        'Call the function to format forum codes
		strSignature = FormatForumCodes(strSignature)

	        'Call the filters to remove malcious HTML code
	        strSignature = HTMLsafe(strSignature)
	        

	        'If the user has not entered a hoempage then make sure the homepage variable is blank
	        If strHomepage = "http://" OR strHomepage =  "https://" Then strHomepage = ""
	End If
	

	strMemberTitle = removeAllTags(strMemberTitle)
	strMemberTitle = formatInput(strMemberTitle)


	
	
	
	'******************************************
	'***    Check Password Complexity	***
	'******************************************
	
	'Check for passowrd complexity
	If blnEnforceComplexPasswords AND strPassword <> "" Then blnPasswordComplexityOK = passwordComplexity(strPassword, intMinPasswordLength)
		
		
	
        
        
        
        
        
        
        '******************************************
	'*** 	 	Remove bad words	***
	'******************************************

        'Replace swear words with other words with ***
        'Initalise the SQL string with a query to read in all the words from the smut table
        strSQL = "SELECT " & strDbTable & "Smut.* " & _
        "FROM " & strDbTable & "Smut" & strDBNoLock & ";"

        'Open the recordset
        rsCommon.Open strSQL, adoCon
        
         'Create regular experssions object
	Set objRegExp = New RegExp

        'Loop through all the words to check for
        Do While NOT rsCommon.EOF

               
                'Read in the smut words
                strSmutWord = rsCommon("Smut")
                strSmutWordReplace = rsCommon("Word_replace")
                
                'Tell the regular experssions object what to look for
		With objRegExp
			.Pattern = strSmutWord
			.IgnoreCase = True
			.Global = True
		End With
		
		'Ignore errors, incase someone entered an incorrect bad word that breakes regular expressions
		On Error Resume Next

                'Replace the swear words with the words in the database the swear words
                If strMode = "new-reg" AND objRegExp.Execute(strUsername).Count > 0 Then blnUsernameOK = False 'If username contains a smut word then make the user choose another username
                If strCustRegItemName1 <> "" Then strCustItem1 = objRegExp.Replace(strCustItem1, strSmutWordReplace)
        	If strCustRegItemName2 <> "" Then strCustItem2 = objRegExp.Replace(strCustItem2, strSmutWordReplace)
        	If strCustRegItemName3 <> "" Then strCustItem3 = objRegExp.Replace(strCustItem3, strSmutWordReplace)
                strRealName = objRegExp.Replace(strRealName, strSmutWordReplace)
                strGender = objRegExp.Replace(strGender, strSmutWordReplace)
                strSignature = objRegExp.Replace(strSignature, strSmutWordReplace)
                strFacebookUsername = objRegExp.Replace(strFacebookUsername, strSmutWordReplace)
	        strTwitterUsername = objRegExp.Replace(strTwitterUsername, strSmutWordReplace)
		strLinkedInUsername = objRegExp.Replace(strLinkedInUsername, strSmutWordReplace)
                strAIMAddress = objRegExp.Replace(strAIMAddress, strSmutWordReplace)
                strMSNAddress = objRegExp.Replace(strMSNAddress, strSmutWordReplace)
                strYahooAddress = objRegExp.Replace(strYahooAddress, strSmutWordReplace)
                strOccupation = objRegExp.Replace(strOccupation, strSmutWordReplace)
                strInterests = objRegExp.Replace(strInterests, strSmutWordReplace)
                
                'Disable error trapping
		On Error goto 0
                
                 'Move to the next word in the recordset
                rsCommon.MoveNext
        Loop
        
        'Distroy regular experssions object
	Set objRegExp = nothing

        'Release the smut recordset object
        rsCommon.Close
        






	'******************************************
	'***     Check the avatar is OK		***
	'******************************************

        'Remove malicious code form the avatar link or remove it all togtaher if not a web graphic
        If strAvatar <> "" Then
        	
        	'Call the filter for the image
                strAvatar = checkImages(strAvatar)
                strAvatar = formatInput(strAvatar)
        End If




	'******************************************
	'***     Check the username is OK	***
	'******************************************

        'If this is a new reg clean up the username
        If strMode = "new-reg" Then

                'Check there is a username
                If Len(strUsername) < intMinUsernameLength Then blnUsernameOK = False

                'Make sure the user has not entered disallowed usernames
                If InStr(1, strUsername, "admin", vbTextCompare) Then blnUsernameOK = False
        End If

	'******************************************
	'***     Check signature lentgh OK	***
	'******************************************
	
	'Trim signature down to a 255 max characters to prevent database errors
	strSignature = Mid(strSignature, 1, 255)



	



	'******************************************
	'*** 	  Check input if new reg	***
	'******************************************

        'If this is a new reg then check the username and genrate usercode, setup email activation etc.
        If strMode = "new-reg" Then

        	'******************************************
		'***   Check the username is availabe	***
		'******************************************

                'If the username is not already written off then check it's not already gone
                If blnUsernameOK Then
                	
                	'Make username SQL safe
       			strTempUsername = formatSQLInput(strUsername)


                        'Read in the the usernames from the database to check that the username does not already exsist

                        'Initalise the strSQL variable with an SQL statement to query the database
                        strSQL = "SELECT " & strDbTable & "Author.Username " & _
                        "FROM " & strDbTable & "Author" & strDBNoLock & "  " & _
                        "WHERE " & strDbTable & "Author.Username = '" & strTempUsername & "';"

                        'Query the database
                        rsCommon.Open strSQL, adoCon

                        'If there is a record returned from the database then the username is already used
                        If NOT rsCommon.EOF Then blnUsernameOK = False

                        'Close the recordset
                        rsCommon.Close

                      

			'******************************************
			'***   Get the starting group ID	***
			'******************************************

                        'Get the starting group ID number

                        'Initalise the strSQL variable with an SQL statement to query the database
                        strSQL = "SELECT " & strDbTable & "Group.Group_ID " & _
                        "FROM " & strDbTable & "Group" & strDBNoLock & " " & _
                        "WHERE " & strDbTable & "Group.Starting_group = " & strDBTrue & ";"

                        'Query the database
                        rsCommon.Open strSQL, adoCon

                        'Get the forum starting group ID number
                        intForumStartingGroup = CInt(rsCommon("Group_ID"))

                        'Close the recordset
                        rsCommon.Close
                End If


		'******************************************
		'***  Check email domain is not banned	***
		'******************************************

                'Initalise the strSQL variable with an SQL statement to query the database
                strSQL = "SELECT " & strDbTable & "BanList.Email " & _
                "FROM " & strDbTable & "BanList" & strDBNoLock & " " & _
                "WHERE " & strDbTable & "BanList.Email Is Not Null;"

                'Query the database
                rsCommon.Open strSQL, adoCon

                'Loop through the email address and check 'em out
                Do while NOT rsCommon.EOF

                        'Read in the email address to check
                        strCheckEmailAddress = rsCommon("Email")

                        'If a whildcard character is found then check that
                        If Instr(1, strCheckEmailAddress, "*", 1) > 0 Then

	                        'Remove the wildcard charcter from the email address to check
	                        strCheckEmailAddress = Replace(strCheckEmailAddress, "*", "", 1, -1, 1)

	                        'If the banned email and the email entered match up then don't let em sign up
	                        If InStr(1, strEmail, strCheckEmailAddress, 1) Then blnEmailBlocked = True

	                        '2nd check Use the same filters as that on the email address being checked
	        		strCheckEmailAddress = formatInput(strCheckEmailAddress)

	                        'If the banned email and the email entered match up then don't let em sign up
	                        If InStr(1, strEmail, strCheckEmailAddress, 1) Then blnEmailBlocked = True

	                'Else check the actual name doesn't match
	                Else

	                        'If the banned email and the email entered match up then don't let em sign up
	                        If strCheckEmailAddress = strEmail Then blnEmailBlocked = True
	        	End If

                        'Move to the next record
                        rsCommon.MoveNext
                Loop

                'Close recordset
                rsCommon.Close


		'******************************************
		'***  Check email address is availabe	***
		'******************************************

                'If e-mail activation is on then check the email address is not already used
                If blnEmailActivation = True OR blnMemberApprove = True Then
                	
                	'SQL safe format call
       			strTempEmail = formatSQLInput(strEmail)

                        'Initalise the strSQL variable with an SQL statement to query the database
                        strSQL = "SELECT " & strDbTable & "Author.Author_email " & _
                        "FROM " & strDbTable & "Author" & strDBNoLock & " " & _
                        "WHERE " & strDbTable & "Author.Author_email = '" & strTempEmail & "';"

                        'Query the database
                        rsCommon.Open strSQL, adoCon

                        'If there is a record returned from the database then the email address is already used
                        If NOT rsCommon.EOF Then blnEmailOK = False

                        'Close recordset
                        rsCommon.Close

                End If

		'******************************************
		'*** 	     Create a usercode 		***
		'******************************************

                'Calculate a code for the user
                strUserCode = userCode(strUsername)


	'******************************************
	'***   If update, update usercode	***
	'******************************************

        'Else this is an update so just calculate a new usercode
        Else

                'Calculate a new code for the user
                strUserCode = userCode(strLoggedInUsername)

        End If
        
          
          
        '******************************************
	'*** 	  	StopForumSpam		***
	'******************************************
        
        If blnStopForumSpam AND strMode = "new-reg" Then
        
	        'To reduce lookups with StopForumSpam only run check if everything else has gone OK
	         If blnUsernameOK AND blnEmailBlocked = False AND blnEmailOK AND blnValidEmail AND blnRequiredFieldsValid Then
	         
	         	'Run stop forum spam function
	         	Dim blnSpamUsername, blnSpamEmail, blnSpamIP
	         	
	         	'Call stopforumspam function
	         	Call StopForumSpamLookup(strEmail, strUsername,  getIP())
	         	
	         	'If looking for username, IP, and email
	         	If blnStopForumSpamUsername Then
	         		
	         		'If Username, Email, and IP are true set blnStopForumSpamFound to true
	         		If blnSpamUsername AND blnSpamEmail AND blnSpamIP Then blnStopForumSpamFound = True
	         	
	         	'Else only check IP and email address		
	         	Else
	         		'If email and IP found then set blnStopForumSpamFound to true
	         		If blnSpamEmail OR blnSpamIP Then blnStopForumSpamFound = True
	         	End If
	         	
	         	

	         	'If found in StopForumSpamFound database
	         	If blnStopForumSpamFound Then
		         	
		         	'If logging is enabled for new registrations record the failed attempt to the log file
			 	If blnLoggingEnabled AND blnNewRegistrationLogging Then Call logAction("StopForumSpam", "REGISTRATION REJECTION - Username; " & strUsername & " " & blnSpamUsername & " - Email; " & strEmail & " " & blnSpamEmail & " - IP; " & getIP() & " " & blnSpamIP)
			 		
			 	'Update number total registrations StopForumSpam has blocked
			 	lngStopForumSpamBlocked = lngStopForumSpamBlocked + 1
			 	Call addConfigurationItem("SFSpam_no_blocked", lngStopForumSpamBlocked)
			 	
			 	'Update application variables
			 	Application.Lock
				Application(strAppPrefix & "lngStopForumSpamBlocked") = CLng(lngStopForumSpamBlocked)
				Application.UnLock
			 	
			 	'Update last blocked user details
			 	Call addConfigurationItem("SFSpam_last_block_date", internationalDateTime(Now()))
			 	Call addConfigurationItem("SFSpam_last_block_username", strUsername)
			 	Call addConfigurationItem("SFSpam_last_block_email", strEmail)
			 	Call addConfigurationItem("SFSpam_last_block_IP", getIP())
			 	
			 	'If the registratnts email is found and not IP then submit the spammers new IP to StopForumSpam
			 	'If blnSpamEmail = True AND blnSpamIP = False AND strStopForumSpamApiKey <> "" Then
			 	'	
			 	'	Call StopForumSpamSubmit(strEmail, strUsername, getIP())
			 	'	
			 	'End If
	 	
			End If
	 	End If 
	End If
	
	
	
	
	
	
	'******************************************
	'*** 	  IP to Country Block		***
	'******************************************
        
        If NOT strCountryBlockRegList = "" AND strMode = "new-reg" Then
        	
	        'To reduce lookups only run check if everything else has gone OK
	         If blnUsernameOK AND blnEmailBlocked = False AND blnEmailOK AND blnValidEmail AND blnRequiredFieldsValid AND blnStopForumSpamFound = False Then
	         
	         	'See if the IP address location is in country block list
	         	If InStr(strCountryBlockRegList, IpCountryLookup(getIP(), strInstallID)) Then
	         		blnIpCountryBanned = True
	         	End If
	         End If
	End If





	'******************************************
	'*** Read in user details from database ***
	'******************************************

        'Intialise the strSQL variable with an SQL string to open a record set for the Author table
        strSQL = "SELECT " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Group_ID, " & strDbTable & "Author.Username, " & strDbTable & "Author.Real_name, " & strDbTable & "Author.Gender, " & strDbTable & "Author.User_code, " & strDbTable & "Author.Password, " & strDbTable & "Author.Salt, " & strDbTable & "Author.Author_email, " & strDbTable & "Author.Homepage, " & strDbTable & "Author.Location, " & strDbTable & "Author.MSN, " & strDbTable & "Author.Yahoo, " & strDbTable & "Author.ICQ, " & strDbTable & "Author.AIM, " & strDbTable & "Author.Occupation, " & strDbTable & "Author.Interests, " & strDbTable & "Author.DOB, " & strDbTable & "Author.Signature, " & strDbTable & "Author.No_of_posts, " & strDbTable & "Author.Points, " & strDbTable & "Author.No_of_PM, " & strDbTable & "Author.Join_date, " & strDbTable & "Author.Avatar, " & strDbTable & "Author.Avatar_title, " & strDbTable & "Author.Last_visit, " & strDbTable & "Author.Time_offset, " & strDbTable & "Author.Time_offset_hours, " & strDbTable & "Author.Date_format, " & strDbTable & "Author.Show_email, " & strDbTable & "Author.Attach_signature, " & strDbTable & "Author.Active, " & strDbTable & "Author.Rich_editor, " & strDbTable & "Author.Reply_notify, " & strDbTable & "Author.PM_notify, " & strDbTable & "Author.Skype, " & strDbTable & "Author.Login_attempt, " & strDbTable & "Author.Banned, " & strDbTable & "Author.Info, " & strDbTable & "Author.Newsletter, " & strDbTable & "Author.Login_IP, " & strDbTable & "Author.Custom1, " & strDbTable & "Author.Custom2, " & strDbTable & "Author.Custom3, " & strDbTable & "Author.Facebook, " & strDbTable & "Author.Twitter, " & strDbTable & "Author.LinkedIn " &_
	"FROM " & strDbTable & "Author" & strRowLock & " " & _
	"WHERE " & strDbTable & "Author.Author_ID = " & lngUserProfileID & ";"

        'Set the cursor type property of the record set to Forward Only
        rsCommon.CursorType = 0

        'Set the Lock Type for the records so that the record set is only locked when it is updated
        rsCommon.LockType = 3

        'Open the author table
        rsCommon.Open strSQL, adoCon




	'********************************************
	'*** Update the usercode if in admin mode ***
	'********************************************

        'If there is a record and in admin mode update the user code to activate or suspend the member
        If NOT rsCommon.EOF AND blnAdminMode Then

        	'Read in the usercode to check incase we are suspending or unsuspending the account
        	strUserCode = rsCommon("User_code")

        	'If we are suspending the user account then update the user code
        	If (blnUserActive = False OR blnSuspended) AND lngUserProfileID > 2 Then

        	 	strUserCode = userCode(strUsername)
        	End If
        End If



	'********************************************
	'*** Don't let moderator update admin mem ***
	'********************************************

        'Once the author table is open if this is an update and admin mode is on and the updater is a moderator check that the account being updated is not an admin account
        If strMode = "update" AND blnAdminMode AND blnModerator AND NOT rsCommon.EOF Then

                'If the account being updated is an admin account and the updater is only a moderator then send 'em away
                If CInt(rsCommon("Group_ID")) = 1 Then

                        'clean up before redirecting
                        rsCommon.Close
                        Call closeDatabase()

                        'redirect to insufficient permissions page
                        Response.Redirect("insufficient_permission.asp?FID=" & intForumID & strQsSID3)
                End If
        End If


	'******************************************
	'*** 		Encrypt password	***
	'******************************************

        'Encrypt password
	If blnEncryptedPasswords Then

	        If strPassword <> "" Then

	                'If this is a new reg then generate a salt value
	                If strMode = "new-reg" Then
	                        strSalt = getSalt(Len(strPassword))

	                'Else this is an update so get the salt value from the db
	                Else
	                        strSalt = rsCommon("Salt")
	                End If

	                'Concatenate salt value to the password
	                strEncryptedPassword = strPassword & strSalt
	                strConfirmPassword = strConfirmPassword & strSalt

	                'Encrypt the password
	                strEncryptedPassword = HashEncode(strEncryptedPassword)
	                strConfirmPassword = HashEncode(strConfirmPassword)
	        End If

	'Else the password is not set to be encrypted so place the un-encrypted password into the strEncryptedPassword variable
	Else

		strEncryptedPassword = strPassword
	End If




	'******************************************
	'*** 		Update password		***
	'******************************************

	'If this is an update then check the user has not change their password
	If strMode = "update" AND strPassword <> "" Then

	      	'Check the old password matches that of the confirmed password
	        If strConfirmPassword <> rsCommon("Password") AND blnAdminMode = false Then blnConfirmPassOK = false


		'If the password doesn't match that stored in the db then this is a password update
	        If rsCommon("Password") <> strEncryptedPassword AND blnConfirmPassOK Then

			'If encrypted passwords
			If blnEncryptedPasswords Then 
		                
		                'Generate new salt
		                 strSalt = getSalt(Len(strPassword))
	
		         	'Concatenate salt value to the password
		           	strEncryptedPassword = strPassword & strSalt
	
		         	'Re-Genreate encypted password with new salt value
		            	strEncryptedPassword = HashEncode(strEncryptedPassword)
		        
		        'Else if not using encrypted passwords
		        Else
		        	strEncryptedPassword = strPassword
			End If

	                'Set the changed password boolean to true
	                If blnDemoMode = False Then blnPasswordChange = True
	        End If
	  End If





	'******************************************
	'*** 	  Check for email update	***
	'******************************************

        'If e-mail activation is on then check the user has not changed there e-mail address
        If blnEmailActivation AND blnAdmin = False AND (strMode = "update" AND (intUpdatePartNumber = 1 OR intUpdatePartNumber = 0)) Then

                'If the old and new e-mail addresses don't match set the reactivation boolean to true
                If rsCommon("Author_email") <> strEmail Then blnAccountReactivate = True
        End If
        
        
        


	'******************************************
	'*** 	  	Update datbase		***
	'******************************************

        'If this is new reg and the username and email is OK or this is an update then register the new user or update the rs
        If (strMode = "new-reg" AND blnUsernameOK AND blnEmailBlocked = False AND blnEmailOK AND blnSecurityCodeOK AND blnValidEmail AND blnStopForumSpamFound = False AND blnIpCountryBanned = False AND blnPasswordComplexityOK) OR (strMode = "update" AND blnConfirmPassOK AND blnValidEmail) AND blnPasswordComplexityOK AND blnRequiredFieldsValid Then


                'If this is new then create a new rs and reset session variable
                If strMode = "new-reg" Then rsCommon.AddNew


                'Insert the user's details into the rs
                With rsCommon
                
                        If strMode = "new-reg" Then
                        	.Fields("Username") = strUsername
				.Fields("Group_ID") = intForumStartingGroup
				.Fields("Join_date") = internationalDateTime(Now())
				.Fields("Last_visit") = internationalDateTime(Now())
				.Fields("Banned") = False
				.Fields("Info") = "" 'This is to prevent errors in mySQL
				.Fields("No_of_posts") = 0
				.Fields("No_of_PM") = 0
				.Fields("Login_attempt") = 0
				.Fields("Login_IP") = Trim(Mid(getIP(), 1, 50))
			End If



                        'If part number = 0 (all) or part 1 (reg details) then run this code
                        If intUpdatePartNumber = 0 OR intUpdatePartNumber = 1 Then

	                        If (strMode = "update" AND blnPasswordChange = True) OR  strMode = "new-reg" Then .Fields("Password") = strEncryptedPassword
	                        If (strMode = "update" AND blnPasswordChange = True) OR  strMode = "new-reg" Then .Fields("Salt") = strSalt
	                        If blnWindowsAuthentication = False Then .Fields("User_code") = strUserCode
	                        .Fields("Author_email") = strEmail
	                End If




                        'If part number = 0 (all) or part 2 (profile details) then run this code
                        If intUpdatePartNumber = 0 OR intUpdatePartNumber = 2 Then

				If strCustRegItemName1 <> "" Then .Fields("Custom1") = strCustItem1
				If strCustRegItemName2 <> "" Then .Fields("Custom2") = strCustItem2
				If strCustRegItemName3 <> "" Then .Fields("Custom3") = strCustItem3

				.Fields("Real_name") = strRealName
		        	.Fields("Gender") = strGender
		        	.Fields("Location") = strLocation
		       		.Fields("Avatar") = strAvatar


		                'If this is new reg then don't include profile info in the add new
                        	If (blnLongRegForm AND strMode = "new-reg") OR strMode <> "new-reg" Then

		                        .Fields("Homepage") = strHomepage
		                        .Fields("Facebook") = strFacebookUsername
		                        .Fields("Twitter") = strTwitterUsername
		                        .Fields("LinkedIn") = strLinkedInUsername
		                        .Fields("ICQ") = strICQNum
		                        .Fields("AIM") = strAIMAddress
		                        .Fields("MSN") = strMSNAddress
		                        .Fields("Yahoo") = strYahooAddress
		                        .Fields("Skype") = strSkypeName
		                        .Fields("Occupation") = strOccupation
		                        .Fields("Interests") = strInterests
		                        .Fields("DOB") = dtmDateOfBirth
		                        .Fields("Signature") = strSignature
		                        .Fields("Attach_signature") = blnAttachSignature
	                	Else
	                		.Fields("Attach_signature") = true
	                	End If
                	End If




                        'If part number = 0 (all) or part 3 (forum preferences) then run this code
                        If intUpdatePartNumber = 0 OR intUpdatePartNumber = 3 Then

	                        .Fields("Date_format") = strDateFormat
	                        .Fields("Time_offset") = strTimeOffSet
	                        .Fields("Time_offset_hours") = intTimeOffSet
	                        .Fields("Reply_notify") = blnReplyNotify
	                        .Fields("Rich_editor") = blnWYSIWYGEditor
	                        .Fields("PM_notify") = blnPMNotify
	                        .Fields("Show_email") = blnShowEmail
	                        If blnWebWizNewsPad Then .Fields("Newsletter") = blnNewsletter
	                End If




                        'If the e-mail activation is on and this is a new reg or an update and the account needs reactivating then don't activate the account
                        If (((blnEmailActivation OR blnMemberApprove) AND strMode = "new-reg") OR blnAccountReactivate) AND blnModerator = False Then
                                .Fields("Active") = False
                        Else
                                .Fields("Active") = True
                        End If




                        'If the admin mode is enabled then the admin can also update some other member parts
                        If blnAdminMode AND (blnAdmin Or blnModerator) AND strMode = "update" AND blnDemoMode = False Then

                        	If lngUserProfileID > 2 Then .Fields("Active") = blnUserActive

                        	.Fields("Avatar_title") = strMemberTitle
				If blnSuspended Then
					.Fields("Banned") = True
				Else
					.Fields("Banned") = False
				End If
				.Fields("Info") = strAdminNotes

				If isEmpty(lngPosts) = False Then .Fields("No_of_posts") = lngPosts
				If isEmpty(lngMemberPoints) = False Then .Fields("Points") = lngMemberPoints

                        	'If the user is also the admin then let them update some other parts
                        	If blnAdmin AND lngUserProfileID > 2 Then
                        		.Fields("Group_ID") = intUsersGroupID
                		End If
                		
                		'If logging enabled log moderator update user profile
	                 	If blnLoggingEnabled AND blnModeratorLogging Then Call logAction(strLoggedInUsername, "Admin/Moderator Edited Forum Profile of " & strUsername)
                        End If

			
			'Set error trapping
			On Error Resume Next

                        'Update the database with the new user's details (needed for MS Access which can be slow updating)
                        .Update
                        
                        'If an error has occurred write an error to the page
			If Err.Number <> 0 AND strMode = "new-reg" Then 
				Call errorMsg("An error has occurred while writing to the database.", "register_USR", "register.asp")
			ElseIf Err.Number <> 0 Then
				Call errorMsg("An error has occurred while writing to the database.", "update_USR", "register.asp")
			End If

			'Disable error trapping
			On Error goto 0

                        'Re-run the query (required for Access to give it time to update on slower servers)
                        .Requery

	                 'Close rs
	                 .Close
	                 
	                 
	                 'If logging enabled log new registration
	                 If strMode = "new-reg" AND blnLoggingEnabled AND blnNewRegistrationLogging Then Call logAction(strUsername, "New User Registration")
                End With



		'******************************************
		'*** 	     Create usercode cookie	***
		'******************************************

                'Write a cookie with the User ID number so the user logged in throughout the forum
                'But only if not in admin modem and using all parts of part 1 of the reg form
                If (blnAdminMode = False) AND (intUpdatePartNumber = 0 OR intUpdatePartNumber = 1) AND blnWindowsAuthentication = False Then

                        'Write the cookie with the name Forum containing the value UserID number
   			
   			Call saveSessionItem("UID", strUserCode)

                        'Auto Login cookie
                        If blnAutoLogin Then
                                Call setCookie("sLID", "UID", strUserCode, True)
                        'Temp Cookie
                        Else
                        	Call setCookie("sLID", "UID", strUserCode, False)
                        End If
                End If




		'******************************************
		'*** 	   Send activate email   	***
		'******************************************

                'Inititlaise the subject of the e-mail that may be sent in the next if/ifelse statements
                strSubject = strTxtWelcome & " " & strTxtEmailToThe & " " & strMainForumName

                'If the members account needs to be activated or reactivated then send the member a re-activate mail a redirect them to a page to tell them there account needs re-activating
                If ((blnEmailActivation OR blnMemberApprove) AND strMode = "new-reg") OR blnAccountReactivate Then


                	'If new registration we need to get the new users ID from the database
                	If strMode = "new-reg" Then
	                       
	                        'SQL to get the new Author_ID from the database
				strSQL = "SELECT " & strDBTop1 & " " & strDbTable & "Author.Author_ID " & _
				"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
				"ORDER BY " & strDbTable & "Author.Author_ID DESC" & strDBLimit1 & ";"

				'Query database
				rsCommon.Open strSQL, adoCon

	                        'Read back in the user ID for the activation email
	                        lngUserProfileID = CLng(rsCommon("Author_ID"))

	                        'Close rs
	                        rsCommon.Close
	                End If



       
                       'If the admin needs to apporove the member send the activation email to the forum admin
                        If blnMemberApprove Then
                        
                        	'Create admin activation email
	                        strEmailBody = strTxtHi & ", " & _
	                        "<br /><br />" & strTxtEmailNewUserRegistered & " " & strMainForumName & "." & _
	                        "<br /><br />" &  "----------------------------" & _
	                        "<br />" &  strTxtUsername & ": - " & decodeString(strUsername) & _
	                        "<br />" &  strTxtEmailAddress & ": - " & strEmail & _
	                        "<br />" &  strTxtIPLogged & ": - <a href=""https://network-tools.webwiz.net/ip-information.htm?ip=" & Server.URLEncode(getIP()) & """ target=""_blank"">" & getIP() & "</a>" & _	
	                        "<br />" &  "----------------------------" & _
	                        "<br /><br />" & strTxtToViewTheDetailsOf & " " & decodeString(strUsername) & " " & strTxtBeforeApprovingClickTheLinkBelow & ": -" & _
	                        "<br /><a href=""" & strForumPath & "member_profile.asp?PF=" & lngUserProfileID & """>" & strForumPath & "member_profile.asp?PF=" & lngUserProfileID & "</a>" & _
	                         vbCrLf & _
	                        "<br /><br />" & strTxtToViewInfoOnIPAddressOf & " " & decodeString(strUsername) & " " & strTxtForumClickOnTheLinkBelow & ": -" & _
	                        "<br /><a href=""https://network-tools.webwiz.net/ip-information.htm?ip=" & Server.URLEncode(getIP()) & """ target=""_blank"">https://network-tools.webwiz.net/ip-information.htm?ip=" & getIP() & "</a>"
	                        
	                        'If StopForumSpam is enabled and has an API key include link to submit the new memnber as a spammer
	                        If blnStopForumSpam AND strStopForumSpamApiKey <> "" Then
	                        	
	                        	strEmailBody = strEmailBody & vbCrLf & _
		                        "<br /><br />" & strTxtIfYourAreSureThat & " " & decodeString(strUsername) & " " & strTxtIsASmammerSubmitToSpamDatabase & ": -" & _
		                        "<br /><a href=""https://www.stopforumspam.com/add.php?username=" & Server.URLEncode(decodeString(strUsername)) & "&email=" & Server.URLEncode(strEmail) & "&IP=" & Server.URLEncode(getIP()) & "&api_key=" & Server.URLEncode(strStopForumSpamApiKey) & """ target=""_blank"">https://www.stopforumspam.com/add.php?username=" & Server.URLEncode(decodeString(strUsername)) & "&email=" & Server.URLEncode(strEmail) & "&IP=" & Server.URLEncode(getIP()) & "&api_key=" & Server.URLEncode(strStopForumSpamApiKey) & "</a>"
	                        End If
	                        
	                        strEmailBody = strEmailBody & vbCrLf & _
	                        "<br /><br />" & strTxtToActivateTheNewMembershipFor & " " & decodeString(strUsername) & " " & strTxtForumClickOnTheLinkBelow & ": -" & _
	                        "<br /><a href=""" & strForumPath & "admin_activate.asp?USD=" & lngUserProfileID & """>" & strForumPath & "admin_activate.asp?USD=" & lngUserProfileID & "</a>"
	                        
	                        
	                       
                        
                        	'Send the e-mail using the Send Mail function created on the send_mail_function.inc file
                        	blnSentEmail = SendMail(strEmailBody, strTxtForumAdmin, decodeString(strForumEmailAddress), strWebsiteName, decodeString(strForumEmailAddress), strTxtNewMemberActivation, strMailComponent, true)

                        	'If user has an email address send em a welcome email
                        	If blnEmail AND strEmail <> "" Then

	                        	'Initailise the e-mail body variable with the body of the e-mail
		                        strEmailBody = strTxtHi & " " & decodeString(strUsername) & _
		                        vbCrLf & vbCrLf & strTxtEmailThankYouForRegistering & " " & strMainForumName & "." & _
		                        vbCrLf & vbCrLf & strTxtEmailYouCanNowUseOnceYourAccountIsActivatedTheForumAt & " " & strWebsiteName & " " & strTxtEmailForumAt & " " & strForumPath & _
		                        vbCrLf & vbCrLf & "----------------------------" & _
	                        	vbCrLf &  strTxtUsername & ": - " & decodeString(strUsername) & _
		                        vbCrLf & "----------------------------"
		                        If blnEncryptedPasswords Then strEmailBody = strEmailBody & vbCrLf & vbCrLf & strTxtPleaseDontForgetYourPassword

		                        'Send the e-mail using the Send Mail function created on the send_mail_function.inc file
	                        	blnSentEmail = SendMail(strEmailBody, decodeString(strUsername), decodeString(strEmail), strWebsiteName, decodeString(strForumEmailAddress), strSubject, strMailComponent, false)
	                        End If


                        'Send an email to enable the users account to be re-activated
                        ElseIf blnAccountReactivate Then
                        	
                        	'Email subject
                        	strSubject = strMainForumName & " " & strTxtActivationEmail

	                        'Create re-activate email body
	                        strEmailBody = strTxtHi & " " & decodeString(strLoggedInUsername) & _
	                        vbCrLf & vbCrLf & strTxtYourEmailHasChanged & ", " & strMainForumName & ", " & strTxtPleaseUseLinkToReactivate & "." & _
	                        vbCrLf & vbCrLf & strTxtToActivateYourMembershipFor & " " & strMainForumName & " " & strTxtForumClickOnTheLinkBelow & ": -" & _
	                        vbCrLf & vbCrLf & strForumPath & "activate.asp?ID=" & Server.URLEncode(strUserCode) & "&USD=" & lngUserProfileID & "&Type=MemberReactivate"

				'Send the e-mail using the Send Mail function created on the send_mail_function.inc file
                        	blnSentEmail = SendMail(strEmailBody, decodeString(strUsername), decodeString(strEmail), strWebsiteName, decodeString(strForumEmailAddress), strSubject, strMailComponent, false)

			'Else send that this is a new mail account so send activation email
	 		Else

	 			'Create email activate email body
	                        strEmailBody = strTxtHi & " " & decodeString(strUsername) & _
	                        vbCrLf & vbCrLf & strTxtEmailThankYouForRegistering & " " & strMainForumName & "." & _
	                        vbCrLf & vbCrLf & "----------------------------" & _
	                        vbCrLf & strTxtUsername & ": - " & decodeString(strUsername) & _
	                        vbCrLf & "----------------------------" & _
	                        vbCrLf & vbCrLf & strTxtToActivateYourMembershipFor & " " & strMainForumName & " " & strTxtForumClickOnTheLinkBelow & ": -" & _
	                        vbCrLf & vbCrLf & strForumPath & "activate.asp?ID=" & Server.URLEncode(strUserCode) & "&USD=" & lngUserProfileID & "&Type=NewMember"
	                        If blnEncryptedPasswords Then strEmailBody = strEmailBody & vbCrLf & vbCrLf & strTxtPleaseDontForgetYourPassword

				'Send the e-mail using the Send Mail function created on the send_mail_function.inc file
                        	blnSentEmail = SendMail(strEmailBody, decodeString(strUsername), decodeString(strEmail), strWebsiteName, decodeString(strForumEmailAddress), strSubject, strMailComponent, false)
			End If


                        'Reset server Object
                       Call closeDatabase()

			'Redirect if admin activate
			If blnMemberApprove Then
				Response.Redirect("register_confirm.asp?TP=MACT&FID=" & intForumID & strQsSID3)
                        'Redirect the reactivate page
                        ElseIf blnAccountReactivate = True Then
                                Response.Redirect("register_confirm.asp?TP=REACT&FID=" & intForumID & strQsSID3)
                        'Redirect to the activate page
                        Else
                                Response.Redirect("register_confirm.asp?TP=ACT&FID=" & intForumID & strQsSID3)
                        End If


		'******************************************
		'*** 	   Send welcome email   	***
		'******************************************

                'Send the new user a welcome e-mail if e-mail notification is turned on and the user has given an e-mail address
                ElseIf blnEmail AND strEmail <> "" AND strMode = "new-reg" Then

                        'Initailise the e-mail body variable with the body of the e-mail
                        strEmailBody = strTxtHi & " " & decodeString(strUsername) & _
                        vbCrLf & vbCrLf & strTxtEmailThankYouForRegistering & " " & strMainForumName & "." & _
                        vbCrLf & vbCrLf & strTxtEmailYouCanNowUseTheForumAt & " " & strWebsiteName & " " & strTxtEmailForumAt & " " & strForumPath & _
                        vbCrLf & vbCrLf & "----------------------------" & _
                        vbCrLf & strTxtUsername & ": - " & decodeString(strUsername) & _
                        vbCrLf & "----------------------------"
                        If blnEncryptedPasswords Then strEmailBody = strEmailBody & vbCrLf & vbCrLf & strTxtPleaseDontForgetYourPassword
                        

                        'Send the e-mail using the Send Mail function created on the send_mail_function.inc file
                        blnSentEmail = SendMail(strEmailBody, decodeString(strUsername), decodeString(strEmail), strWebsiteName, decodeString(strForumEmailAddress), strSubject, strMailComponent, false)
                End If


		'******************************************
		'*** 	 	 Clean up   		***
		'******************************************

                'Reset server Object
               Call closeDatabase()


		'******************************************
		'*** 	 Redirect to message page	***
		'******************************************

                'Redirect the welcome new user page
                If strMode = "new-reg" Then
                        Response.Redirect("register_confirm.asp?TP=NEW&FID=" & intForumID & strQsSID3)
                'Redirect to the update profile page
                Else
                        Response.Redirect("register_confirm.asp?TP=UPD&FID=" & intForumID & strQsSID3)
                End If

        'Else close rs
        Else
        	rsCommon.Close
        End If
End If




'******************************************
'***         Set the page mode		***
'******************************************

'If this is a new registerant then reset the mode of the page to new
If strMode = "reg" OR strMode = "new-reg" Then

        'set the mode to new
        strMode = "new-reg"

'Else this is an update
Else
        strMode = "update"
End If




'******************************************
'***     Get the user details from db	***
'******************************************

'If this is a profile update get the users details to update
If strMode = "update" Then

        'Read the various forums from the database
        'Initalise the strSQL variable with an SQL statement to query the database
        strSQL = "SELECT " & strDbTable & "Author.* " & _
	"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Author.Author_ID = " & lngUserProfileID

        'Query the database
        rsCommon.Open strSQL, adoCon

        'If there is no matching profile returned by the recordset then redirect the user to the main forum page
        If rsCommon.EOF Then

                'Reset server Object
                rsCommon.Close
                Call closeDatabase()

                Response.Redirect("default.asp" & strQsSID1)
        End If

        'Read in the new user's profile from the recordset
        strUsername = rsCommon("Username")
        strRealName = rsCommon("Real_name")
        If strCustRegItemName1 <> "" Then strCustItem1 = rsCommon("Custom1")
	If strCustRegItemName2 <> "" Then strCustItem2 = rsCommon("Custom2")
	If strCustRegItemName3 <> "" Then strCustItem3 = rsCommon("Custom3")
        strGender = rsCommon("Gender")
        If NOT isNull(rsCommon("Author_email")) Then strEmail = formatInput(rsCommon("Author_email"))
        If blnWebWizNewsPad Then blnNewsletter = CBool(rsCommon("Newsletter"))
        blnShowEmail = CBool(rsCommon("Show_email"))
        If NOT isNull(rsCommon("Homepage")) Then strHomepage = formatInput(rsCommon("Homepage"))
        If NOT isNull(rsCommon("Location")) Then strLocation = rsCommon("Location")
        strSignature = rsCommon("Signature")
        strAvatar = rsCommon("Avatar")
        strMemberTitle = rsCommon("Avatar_title")
        strDateFormat = rsCommon("Date_format")
        strTimeOffSet = rsCommon("Time_offset")
        intTimeOffSet = CInt(rsCommon("Time_offset_hours"))
        blnReplyNotify = CBool(rsCommon("Reply_notify"))
        blnAttachSignature = CBool(rsCommon("Attach_signature"))
        blnWYSIWYGEditor = CBool(rsCommon("Rich_editor"))
        If NOT isNull(rsCommon("Facebook")) Then  strFacebookUsername = formatInput(rsCommon("Facebook"))
        If NOT isNull(rsCommon("Twitter")) Then  strTwitterUsername = formatInput(rsCommon("Twitter"))
        If NOT isNull(rsCommon("LinkedIn")) Then  strLinkedInUsername = formatInput(rsCommon("LinkedIn"))
        If NOT isNull(rsCommon("ICQ")) Then  strICQNum = formatInput(rsCommon("ICQ"))
        If NOT isNull(rsCommon("AIM")) Then strAIMAddress = formatInput(rsCommon("AIM"))
        If NOT isNull(rsCommon("MSN")) Then strMSNAddress = formatInput(rsCommon("MSN"))
        If NOT isNull(rsCommon("Yahoo")) Then strYahooAddress = formatInput(rsCommon("Yahoo"))
        If NOT isNull(rsCommon("Skype")) Then strSkypeName = formatInput(rsCommon("Skype"))
        strOccupation = rsCommon("Occupation")
        strInterests = rsCommon("Interests")
        dtmDateOfBirth = rsCommon("DOB")
        blnPMNotify = CBool(rsCommon("PM_notify"))

        'If we are in admin mode then read on extra user details
        If blnAdminMode Then
                intUsersGroupID = CInt(rsCommon("Group_ID"))
                blnUserActive = CBool(rsCommon("Active"))
                If isNull(rsCommon("No_of_posts")) Then lngPosts = 0 Else lngPosts = CLng(rsCommon("No_of_posts"))
		If isNull(rsCommon("Points")) Then lngMemberPoints = 0 Else lngMemberPoints = CLng(rsCommon("Points"))
                blnSuspended = CBool(rsCommon("Banned"))
                strAdminNotes = rsCommon("Info")
        End If

        'Reset Server Objects
        rsCommon.Close


        'If admin mode is on and the user is only a moderator and the edited account is an admin account then the modertor can not edit the account
        If blnAdminMode AND blnModerator AND intUsersGroupID = 1 Then


                'clean up before redirecting
                Call closeDatabase()

                'redirect to insufficient permissions page
                Response.Redirect("insufficient_permission.asp?FID=" & intForumID & strQsSID3)
        End If


        'Split the date of biith into the various parts
        If isDate(dtmDateOfBirth) Then
	        intDOBYear = Year(dtmDateOfBirth)
		intDOBMonth = Month(dtmDateOfBirth)
		intDOBDay = Day(dtmDateOfBirth)
	End If
End If



'******************************************
'***  	    De-code signature		***
'******************************************

'Covert the signature back to forum codes
If strSignature <> "" Then  strSignature = EditPostConvertion(strSignature)



'**********************************
'***  	   Form Key ID		***
'********************************** 

'Create a form key ID (done for extra security)
strFormKey = LCase(hexValue(14))
Call saveSessionItem("IDX", strFormKey)

'Create encrypted form fields
strUserNameFormName = "N" & HashEncode("Username" & strFormKey)
strEmailFormName = "E" & HashEncode("Email" & strFormKey)
strPasswordFormName = "P" & HashEncode("Password" & strFormKey)
strFormKey = HashEncode("IDX" & strFormKey)



'Set bread crumb trail
If strMode = "update" Then
	strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtEditProfile
Else
	strBreadCrumbTrail = strBreadCrumbTrail & strNavSpacer & strTxtRegisterNewUser
End If

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% If strMode = "update" Then Response.Write(strTxtEditProfile) Else Response.Write(strTxtRegisterNewUser) %></title>

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

<!-- Check the from is filled in correctly before submitting -->
<script language="JavaScript">

//Function to check form is filled in correctly before submitting
function CheckForm () {

        //Initialise variables
        var errorMsg = "";
        var errorMsgLong = "";
	var formArea = document.getElementById('frmRegister');
<%
'If this is new reg then make sure the user eneters a username and password
If strMode ="new-reg" Then

	%>
        if (formArea.<% = strUserNameFormName %>.value.length < <% = intMinUsernameLength %>){errorMsg += "\n<% = strTxtErrorUsernameChar & " " &  intMinUsernameLength & " " & strTxtCharacters %>";}
        if (formArea.<% = strPasswordFormName %>1.value.length < <% = intMinPasswordLength %>){errorMsg += "\n<% = strTxtErrorPasswordChar & " " & intMinPasswordLength  & " " & strTxtCharacters %>";}<%

'If this is an update only check the password length if the user is enetring a new password
ElseIf (intUpdatePartNumber = 0 OR intUpdatePartNumber = 1) AND blnWindowsAuthentication = False AND (blnMemberAPI = False OR blnMemberAPIDisableAccountControl = False) Then

	%>
        if ((formArea.<% = strPasswordFormName %>1.value.length < 0) && (formArea.<% = strPasswordFormName %>1.value.length < <% = intMinPasswordLength %>)){errorMsg += "\n<% = strTxtErrorPasswordChar & " " & intMinPasswordLength  & " " & strTxtCharacters %>";}<%

End If

'If this is not showing the reg part or all the form then don't run the password and email check js
If (intUpdatePartNumber = 0 OR intUpdatePartNumber = 1) AND blnWindowsAuthentication = False AND (blnMemberAPI = False OR blnMemberAPIDisableAccountControl = False) Then

	'Check both passwords are the same
	%>
        if ((formArea.<% = strPasswordFormName %>1.value) != (formArea.<% = strPasswordFormName %>2.value)){
                errorMsg += "\n<% = strTxtErrorPasswordNoMatch %>";
                formArea.<% = strPasswordFormName %>1.value = "";
                formArea.<% = strPasswordFormName %>2.value = "";
        }
        if (<%

	'If e-mail activation is on check that the e-mail address entered is correct
	If blnEmailActivation = True Then

	        Response.Write("formArea." & strEmailFormName & ".value == """" || ")
	Else

	        Response.Write("formArea." & strEmailFormName & ".value.length >0 && ")
	End If
                %>(formArea.<% = strEmailFormName %>.value.indexOf("@",0) == -1||formArea.<% = strEmailFormName %>.value.indexOf(".",0) == -1)) {
                errorMsg +="\n<% = strTxtErrorValidEmail %>";
<%
	'If e-mail activation is not on display a long error message to the user if they enter an incorrect e-mail addres
	If NOT blnEmailActivation = True Then Response.Write("          errorMsgLong += ""\n- " & strTxtErrorValidEmailLong & """; ")
%>
        }
        if (!(formArea.<% = strEmailFormName %>.value == formArea.email2.value)){
                errorMsg +="\n<% = strTxtErrorConfirmEmail %>";
                formArea.email2.focus();
        }<%
 
 	'Only display the show email check if showing full prfile form
	If intUpdatePartNumber = 0 Then
%>
        if (formArea.<% = strEmailFormName %>.value == "" && formArea.emailShow[0].checked == true){
                errorMsgLong += "\n- <% = strTxtErrorNoEmailToShow %>";
                formArea.emailShow[1].checked = true
                formArea.<% = strEmailFormName %>.focus();
        }
<%
	End If
End If


'If this is new reg then make sure the user eneters a username and password
If strMode ="new-reg" AND blnRegistrationCAPTCHA Then
	%>
	//Check for a security code
        if (formArea.securityCode.value == ''){
                errorMsg += "\n<% = strTxtErrorSecurityCode %>";
        }<%

End If



'If real name required
If blnRealNameReq AND (intUpdatePartNumber = 0 OR intUpdatePartNumber = 2) Then
%>
        if (formArea.realName.value == ''){
                errorMsg += "\n<% = strTxtRealNameError %>";
        }<%
        
End If
	
'If location is required
If blnLocationReq AND (intUpdatePartNumber = 0 OR intUpdatePartNumber = 2) Then
%>
        
        //Check for a location code
        if (formArea.location.value == ''){
                errorMsg += "\n<% = strTxtLocationError %>";
        }<%
        
End If


'If custom field 1 is required
If blnReqCustRegItemName1 AND strCustRegItemName1 <> "" AND (intUpdatePartNumber = 0 OR intUpdatePartNumber = 2) Then
%>
        if (formArea.cust1.value == ''){
                errorMsg += "\n<% = strCustRegItemName1 %>   -  <% = strYouMustEnterYour & " " & strCustRegItemName1 %>";
        }<%
        
End If

'If custom field 2 is required
If blnReqCustRegItemName2 AND strCustRegItemName2 <> "" AND (intUpdatePartNumber = 0 OR intUpdatePartNumber = 2) Then
%>
        if (formArea.cust2.value == ''){
                errorMsg += "\n<% = strCustRegItemName2 %>   -  <% = strYouMustEnterYour & " " & strCustRegItemName2 %>";
        }<%
        
End If

'If custom field 3 is required
If blnReqCustRegItemName3 AND strCustRegItemName3 <> "" AND (intUpdatePartNumber = 0 OR intUpdatePartNumber = 2) Then
%>
        if (formArea.cust3.value == ''){
                errorMsg += "\n<% = strCustRegItemName3 %>   -  <% = strYouMustEnterYour & " " & strCustRegItemName3 %>";
        }<%
        
End If


'If long reg form is not on then don't need to check the lengh of the signature, (real name, and location optional)
If ((blnLongRegForm AND strMode = "new-reg") OR (strMode <> "new-reg")) AND (intUpdatePartNumber = 0 OR intUpdatePartNumber = 2) Then

	'If signtaures are enabled check teh user has eneterd one
 	If blnSignatures Then       
%>        
        if (formArea.signature.value.length > 200){
                errorMsg += "\n<% = strTxtErrorSignatureToLong %>";
                errorMsgLong += "\n- <% = strTxtYouHave %> " + document.frmRegister.signature.value.length + " <% = strTxtCharactersInYourSignatureToLong %>";
        }<%
	
	End If
End If
%>
        //If there is aproblem with the form then display an error
        if ((errorMsg != "") || (errorMsgLong != "")){
                msg = "<% = strTxtErrorDisplayLine %>\n\n";
                msg += "<% = strTxtErrorDisplayLine1 %>\n";
                msg += "<% = strTxtErrorDisplayLine2 %>\n";
                msg += "<% = strTxtErrorDisplayLine %>\n\n";
                msg += "<% = strTxtErrorDisplayLine3 %>\n";

                errorMsg += alert(msg + errorMsg + "\n" + errorMsgLong);
                return false;
        }
        formArea.xformID.value='<% = strSessionKey %>';
        formArea.action = 'register.asp?FID=<% = Server.HTMLEncode(intForumID) %><% = strQsSID3 %>';
        formArea.target = '_self';
        return true;
}

function characterCounter(charNoBox, textFeild) {
	document.getElementById(charNoBox).value = document.getElementById(textFeild).value.length;
}

function OpenPreviewWindow(targetPage){

	var formArea = document.getElementById('frmRegister');
	now = new Date
	
	//Open the window first
   	winOpener('','preview',1,1,680,400)

   	//Now submit form to the new window
   	formArea.action = targetPage + '?ID=' + now.getTime();
	formArea.target = 'preview';
	formArea.submit();
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
<!-- #include file="includes/header.asp" -->
<!-- #include file="includes/status_bar_header_inc.asp" -->
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
 <tr>
  <td align="left"><h1><% If strMode = "update" Then Response.Write(strTxtEditProfile) Else Response.Write(strTxtRegisterNewUser) %></h1></td>
</tr>
</table>
<br /><%

'If this is an update and email notify is on show link to email subcriptions
If strMode = "update" AND lngUserProfileID <> 2 Then

%>
<table class="basicTable" cellspacing="0" cellpadding="0" align="center"> 
 <tr> 
  <td class="tabTable">
   <a href="member_control_panel.asp<% If blnAdminMode Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtControlPanel %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>member_control_panel.<% = strForumImageType %>" border="0" alt="<% = strTxtControlPanel %>" /> <% = strTxtControlPanel %></a>
   <a href="register.asp<% If blnAdminMode Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtProfile2 %>" class="tabButtonActive">&nbsp;<img src="<% = strImagePath %>profile.<% = strForumImageType %>" border="0" alt="<% = strTxtProfile2 %>" /> <% = strTxtProfile2 %></a><%
 
	If blnEmail Then

%>
   <a href="email_notify_subscriptions.asp<% If blnAdminMode Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtSubscriptions %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>subscriptions.<% = strForumImageType %>" border="0" alt="<% = strTxtSubscriptions %>" /> <% = strTxtSubscriptions %></a><%
	End If


	'Only disply other links if not in admin mode
	If blnAdminMode = False AND blnActiveMember AND blnPrivateMessages Then

%>
   <a href="pm_buddy_list.asp<% = strQsSID1 %>" title="<% = strTxtBuddyList %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>buddy_list.<% = strForumImageType %>" border="0" alt="<% = strTxtBuddyList %>" /> <% = strTxtBuddyList %></a><%

	End If


	'If file/image uploads
	If blnAttachments OR blnImageUpload Then

%>
   <a href="file_manager.asp<% If blnAdminMode Then Response.Write("?PF=" & lngUserProfileID & "&M=A" & strQsSID2) Else Response.Write(strQsSID1) %>" title="<% = strTxtFileManager %>" class="tabButton">&nbsp;<img src="<% = strImagePath %>file_manager.<% = strForumImageType %>" border="0" alt="<% = strTxtFileManager %>" /> <% = strTxtFileManager %></a><%

	End If
	

%>
  </td>
 </tr>
</table>
<br /><%

End If



'If an error has occurred display what the error is, for those without JS
If blnUsernameOK = False OR blnEmailOK = False OR blnEmailBlocked OR blnConfirmPassOK = False OR blnSecurityCodeOK = False OR blnValidEmail = False OR blnPasswordComplexityOK = False OR blnStopForumSpamFound OR blnRequiredFieldsValid = False OR blnIpCountryBanned Then

	%>
<table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
    <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong><% = strTxtError %></strong></td>
  </tr>
  <tr>
    <td><%

         'If the username is already gone diaply an error message pop-up
        If blnUsernameOK = False Then Response.Write(Replace(strTxtUsrenameGone, "\n\n", "<br />") & "<br /><br />")

        'If the email address is invalid, display an error message
        If blnValidEmail = False Then Response.Write(Replace(strTxtTheEmailAddressEnteredIsInvalid, ".\n\n", "<br />") & "<br /><br />")
        
        'If the email address is used up and email activation is on, display an error message
        If blnEmailOK = False Then Response.Write(Replace(strTxtEmailAddressAlreadyUsed, "\n\n", "<br />") & "<br /><br />")

        'If the email address or domain is blocked
        If blnEmailBlocked = True Then Response.Write(Replace(strTxtEmailAddressBlocked, "\n\n", "<br />") & "<br /><br />")

        'If the confirmed password is incorrect
        If blnConfirmPassOK = False Then Response.Write(Replace(strTxtConformOldPassNotMatching, "\n\n", "<br />") & "<br /><br />")
	
	'If password not complex enough
	If blnPasswordComplexityOK = False Then Response.Write(Replace(strTxtPasswordNotComplex, "\n\n", "<br />") & "<br /><br />")
	
	'If found on StopFourmSpam	
	If blnStopForumSpamFound Then Response.Write(Replace(strTxtInformationGiveFoundInSpamDatabaseAndRejected, "\n\n", "<br />") & "<br /><br />")
	
	'If IP location is banned	
	If blnIpCountryBanned Then Response.Write(Replace(strTxtRegistrationToThisForumIsNotPermittedFromYourCountry, "\n\n", "<br />") & "<br /><br />")
		
	'If no real name
	If blnRealNameOK = False Then Response.Write(Replace(strTxtRealNameError, "\n\n", "<br />") & "<br /><br />")
		
	'If no location
	If blnLocationOK = False Then Response.Write(Replace(strTxtLocationError, "\n\n", "<br />") & "<br /><br />")
	
	'If custom fields are required
	If blnCustRegItemName1OK = False Then Response.Write(Replace(strYouMustEnterYour & " " & strCustRegItemName1, "\n\n", "<br />") & "<br /><br />")
	If blnCustRegItemName2OK = False Then Response.Write(Replace(strYouMustEnterYour & " " & strCustRegItemName2, "\n\n", "<br />") & "<br /><br />")
	If blnCustRegItemName3OK = False Then Response.Write(Replace(strYouMustEnterYour & " " & strCustRegItemName3, "\n\n", "<br />") & "<br /><br />")
		
	 'If the security code is incorrect
        If blnSecurityCodeOK = False Then Response.Write(Replace(strTxtSecurityCodeDidNotMatch, "\n\n", "<br />") & "<br /><br />")
	
%></td>
  </tr>
</table>
<br /><%

End If


%>
<form method="post" name="frmRegister" id="frmRegister" action="register.asp?FID=<% = Server.HTMLEncode(intForumID) %><% = strQsSID2 %>" onReset="return confirm('<% = strResetFormConfirm %>');">
<table cellspacing="1" cellpadding="10" class="tableBorder" align="center"><%




'************************************
'****    Registration Details    ****
'************************************

'If part number = 0 (all) or part 1 (reg details) then show reg details
If intUpdatePartNumber = 0 OR intUpdatePartNumber = 1 Then

     %>
  <tr class="tableLedger">
   <td colspan="2"><% = strTxtRegistrationDetails %></td>
  </tr>
  <tr class="tableSubLedger">
   <td colspan="2"><span class="smText">*<% = strTxtRequiredFields %></span></td>
  </tr>
  <tr class="tableRow">
   <td width="50%"><% = strTxtUsername %>*<br /><span class="smText"><% = strTxtProfileUsernameLong  %></span></td>
   <td width="50%" valign="top"><%

	'If this is a new registration display a filed for the username
	If strMode = "new-reg" Then

        	%><input type="text" name="<% = strUserNameFormName %>" size="15" maxlength="20" value="<% = strUsername %>" autocomplete="off" tabindex="1" /><%

	Else
      		Response.Write(strUsername & "<input type=""hidden"" name=""" & strUserNameFormName  & """ value=""" &  strUsername & """ />")
	End If
	
	
	'Don't show password field when using windows authentication or member API
	If blnWindowsAuthentication = False AND (blnMemberAPI = False OR blnMemberAPIDisableAccountControl = False) Then

%></td>
   </tr>
   <tr class="tableRow">
    <td><% If strMode = "new-reg" Then Response.Write(strTxtPassword & "*") Else Response.Write(strTxtNewPassword) %></td>
    <td valign="top"><input type="password" name="<% = strPasswordFormName %>1" id="<% = strPasswordFormName %>1" size="15" maxlength="20" value="" autocomplete="off" tabindex="2"<% If strMode ="update" AND blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
   </tr>
   <tr class="tableRow">
    <td><% If strMode = "new-reg" Then Response.Write(strTxtRetypePassword & "*") Else Response.Write(strTxtRetypeNewPassword) %></td>
    <td><input type="password" name="<% = strPasswordFormName %>2" id="<% = strPasswordFormName %>2" size="15" maxlength="20" value="" autocomplete="off" tabindex="3"<% If strMode ="update" AND blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
   </tr><%
      		'If update confirm old pass if changing password
      		If strMode ="update" AND blnAdminMode = false Then
%>
   <tr class="tableRow">
    <td><% Response.Write(strTxtConfirmOldPass) %></td>
    <td><input type="password" name="oldPass" id="oldPass" size="15" maxlength="20" value="" autocomplete="off" tabindex="4"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
   </tr><%

		End If
	End If

	%>
   <tr class="tableRow">
    <td><% = strTxtEmail %><%

	'If email or admin activation is on then tell the user for a real email address
	If blnEmailActivation OR blnMemberApprove Then

	        If strMode = "new-reg" Then
	                Response.Write("*<br /><span class=""smText"">" & strTxtEmailRequiredForActvation & "</span><br />")
	        Else
	                Response.Write("*<br /><span class=""smText"">" & strTxtCahngeOfEmailReactivateAccount & "</span><br />")
	        End If
	Else
	        Response.Write("         <br /><span class=""smText"">" & strTxtProfileEmailLong & "</span><br />")
	End If

         %></td>
    <td valign="top"><input type="email" name="<% = strEmailFormName %>" id="<% = strEmailFormName %>" size="30" maxlength="60" value="<% = strEmail %>" tabindex="5" <% If blnMemberAPI AND blnMemberAPIDisableAccountControl Then Response.Write(" readonly=""readonly""") %> /></td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtConfirmEmail %></td>
    <td valign="top"><input type="email" name="email2" id="email2" size="30" maxlength="60" value="<% = strEmail %>" tabindex="5" <% If blnMemberAPI AND blnMemberAPIDisableAccountControl Then Response.Write(" readonly=""readonly""") %> /></td>
   </tr><%
   

End If


'*********************************
'****      Security Code      ****
'*********************************

'If this is a new reg then ask for a seurity code
If strMode = "new-reg" AND blnRegistrationCAPTCHA Then

     %>
   <tr class="tableLedger">
    <td colspan="2"><% = strTxtSecurityCodeConfirmation %></td>
   </tr>
   <tr class="tableRow">
    <td width="50%" valign="top"><% = strTxtUniqueSecurityCode %><br /><span class="smText"><% = strTxtEnterCAPTCHAcode %></span></td>
    <td width="50%" valign="top" tabindex="8"><!--#include file="includes/CAPTCHA_form_inc.asp" --></td>
   </tr><%

End If





'***********************************************
'****    Profile Information (not required?) ****
'***********************************************

If intUpdatePartNumber = 0 OR intUpdatePartNumber = 2 Then

     %>
   <tr class="tableLedger">
    <td colspan="2"><% = strTxtProfileInformation %></td>
   </tr><%
   
   
  	 '***************************
	'****   Custom Reg Items ****
	'****************************
	
	'If custom field 1 is required
	If strCustRegItemName1 <> "" Then
 %> 
   <tr class="tableRow">
    <td width="50%"><% = strCustRegItemName1 %><% If blnReqCustRegItemName1 Then Response.Write("*") %></td>
    <td width="50%"><input type="text" name="cust1" id="cust1" size="30" maxlength="27" value="<% = strCustItem1 %>" tabindex="9" /></td>
   </tr><%
   	
	End If
	
	'If custom field 2 is required
	If strCustRegItemName2 <> "" Then
 %> 
   <tr class="tableRow">
    <td width="50%"><% = strCustRegItemName2 %><% If blnReqCustRegItemName2 Then Response.Write("*") %></td>
    <td width="50%"><input type="text" name="cust2" id="cust2" size="30" maxlength="27" value="<% = strCustItem2 %>" tabindex="10" /></td>
   </tr><%
   	
	End If
	
	'If custom field 3 is required
	If strCustRegItemName3 <> "" Then
 %> 
   <tr class="tableRow">
    <td width="50%"><% = strCustRegItemName3 %><% If blnReqCustRegItemName3 Then Response.Write("*") %></td>
    <td width="50%"><input type="text" name="cust3" id="cust3" size="30" maxlength="27" value="<% = strCustItem3 %>" tabindex="11" /></td>
   </tr><%
   	
	End If
   
   
%>   
   <tr class="tableRow">
    <td width="50%"><% = strTxtRealName %><% If blnRealNameReq Then Response.Write("*") %></td>
    <td width="50%"><input type="text" name="realName" id="realName" size="30" maxlength="27" value="<% = strRealName %>" tabindex="12" /></td>
   </tr>
   <tr class="tableRow">
    <td width="50%"><% = strTxtGender %></td>
    <td width="50%">
     <select name="gender" id="gender" tabindex="13">
      <option value=""<% If strGender = "" Or strGender = null Then Response.Write(" selected") %>><% = strTxtPrivate %></option>
      <option value="<% = strTxtMale %>"<% If strGender = strTxtMale Then Response.Write(" selected") %>><% = strTxtMale %></option>
      <option value="<% = strTxtFemale %>"<% If strGender = strTxtFemale Then Response.Write(" selected") %>><% = strTxtFemale %></option>
     </select>
    </td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtLocation %><% If blnLocationReq Then Response.Write("*") %></td>
    <td>
     <input type="text" name="location" id="location" size="15" maxlength="15" value="<% = strLocation %>" tabindex="14" />
    </td>
   </tr><%

	'If new reg don't show everything
	If ((blnLongRegForm AND strMode = "new-reg") OR strMode <> "new-reg") then
		
		'If the homepgae can be allowed
		If blnHomePage Then

%>
   <tr class="tableRow">
    <td width="50%"><% = strTxtHomepage %></td>
    <td width="50%"><input type="text" name="homepage" size="30" maxlength="48" value="<% If strHomepage = "" Then Response.Write "http://" Else Response.Write(strHomepage) %>" tabindex="15" /></td>
   </tr><%
   
		End If

%>
   <tr class="tableRow">
    <td><% = strTxtFacebook %></td>
    <td><input type="text" name="Facebook" id="Facebook" size="30" maxlength="60" value="<% = strFacebookUsername %>" tabindex="18" /></td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtTwitter %></td>
    <td><input type="text" name="Twitter" id="Twitter" size="30" maxlength="60" value="<% = strTwitterUsername %>" tabindex="18" /></td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtLinkedIn %></td>
    <td><input type="text" name="LinkedIn" id="LinkedIn" size="30" maxlength="60" value="<% = strLinkedInUsername %>" tabindex="18" /></td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtMSNMessenger %></td>
    <td><input type="text" name="MSN" id="MSN" size="30" maxlength="60" value="<% = strMSNAddress %>" tabindex="18" /></td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtSkypeName %></td>
    <td><input type="text" name="skype" id="skype" size="30" maxlength="30" value="<% = strSkypeName %>" tabindex="20" /></td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtYahooMessenger %></td>
    <td><input type="text" name="Yahoo" id="Yahoo" size="30" maxlength="60" value="<% = strYahooAddress %>" tabindex="19" /></td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtAIMAddress %></td>
    <td><input type="text" name="AIM" id="AIM" size="30" maxlength="60" value="<% = strAIMAddress %>" tabindex="17" /></td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtICQNumber %></td>
    <td><input type="text" name="ICQ" id="ICQ" size="15" maxlength="15" value="<% = strICQNum %>" tabindex="16" /></td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtOccupation %></td>
    <td><input type="text" name="occupation" id="occupation" size="30" maxlength="40" value="<% = strOccupation %>" tabindex="21" /></td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtInterests %></td>
    <td><input type="text" name="interests" id="interests" size="30" maxlength="130" value="<% = strInterests %>" tabindex="22" /></td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtDateOfBirth %></td>
    <td><% = strTxtDay %>
     <select name="DOBday" id="DOBday" tabindex="23">
      <option value="0" <% If intDOBDay = 0 Then Response.Write("selected") %>>----</option><%

		'Create lists day's for birthdays
		For lngLoopCounter = 1 to 31
			Response.Write(VbCrLf & "     <option value=""" & lngLoopCounter & """")
			If intDOBDay = lngLoopCounter Then Response.Write("selected")
			Response.Write(">" & lngLoopCounter & "</option>")
		Next

%>
    </select>
    <% = strTxtCMonth %>
    <select name="DOBmonth" id="DOBmonth" tabindex="24">
      <option value="0" <% If intDOBMonth = 0 Then Response.Write("selected") %>>---</option><%

		'Create lists of days of the month for birthdays
		For lngLoopCounter = 1 to 12
			Response.Write(VbCrLf & "      <option value=""" & lngLoopCounter & """")
			If intDOBMonth = lngLoopCounter Then Response.Write("selected")
			Response.Write(">" & lngLoopCounter & "</option>")
		Next

%>
     </select>
     <% = strTxtCYear %>
     <select name="DOByear" id="DOByear" tabindex="25">
      <option value="0" <% If intDOBYear = 0 Then Response.Write("selected") %>>-----</option><%

		'Create lists of years for birthdays
		For lngLoopCounter = CInt(Year(Now()))-99 to CInt(Year(Now()))-6
			Response.Write(VbCrLf & "      <option value=""" & lngLoopCounter & """")
			If intDOBYear = lngLoopCounter Then Response.Write("selected")
			Response.Write(">" & lngLoopCounter & "</option>")
		Next

%>
     </select>
    </td>
   </tr><%

End If

	'------------- Avatar ---------------

	'If avatars are enabled then let the user select an avatar
	If blnAvatar = True Then
%>
   <tr class="tableRow">
    <td valign="top"><% = strTxtSelectAvatar %><br /><span class="smText"><% = strTxtSelectAvatarDetails %>.</span></td>
    <td valign="top" height="2" >
    <table width="290" border="0" cellspacing="0" cellpadding="1">
     <tr>
      <td width="168">
<select name="SelectAvatar" id="SelectAvatar" size="4" onchange="(avatar.src = SelectAvatar.options[SelectAvatar.selectedIndex].value) && (txtAvatar.value='<% If strSslEnabled = "Enabled" Then Response.Write("https://") Else Response.Write("http://") %>') && (oldAvatar.value='')" tabindex="26">
        <option value="<% = strImagePath %>blank.gif"><% = strTxtNoneSelected %></option>
        <!-- #include file="includes/select_avatar.asp" -->
       </select>
      </td>
      <td width="122" align="center"><img src="<%

		'If there is an avatar then display it
		If strAvatar <> "" Then
		     	Response.Write(strAvatar)
		Else
			Response.Write(strImagePath & "blank.gif")
		End If
                %>" name="avatar" id="avatar" />
       <input type="hidden" name="oldAvatar" id="oldAvatar" value="<% = strAvatar %>"/></td>
      </tr>
      <tr>
       <td width="168"><input type="text" name="txtAvatar" id="txtAvatar" size="30" maxlength="95" value="<%

		'If the avatar is the persons own then display the link
		If InStr(1, strAvatar, "http://") > 0 OR InStr(1, strAvatar, "https://") > 0 Then
			Response.Write(strAvatar)
			
		ElseIf strSslEnabled = "Enabled" Then
			Response.Write("https://")
		Else
			Response.Write("http://")
		End If
        %>" onchange="oldAvatar.value=''" tabindex="27" /></td>
      <td width="122"><input type="button" name="preview" id="preview" value="<% = strTxtPreview %>" onclick="avatar.src=txtAvatar.value" tabindex="28" /></td>
     </tr>
    </table><%

		'If avatar uploading is enabled and the user is registered then have a link to it
		If blnAvatarUploadEnabled AND intGroupID <> 2 AND blnActiveMember Then

	%>
    <a href="javascript:winOpener('upload_avatars.asp<% = strQsSID1 %>','avatars',0,1,700,385)" class="smLink"><% = strTxtAvatarUpload %></a>
<%
		End If
%>
    </td>
   </tr><%
	End If

'-----------------------------------------------


	'If new reg don't show everything
	If ((blnLongRegForm AND strMode = "new-reg") OR strMode <> "new-reg") then
		
		'Only show signtaures if enabled
		If blnSignatures Then

%>
   <tr class="tableRow">
    <td valign="top"><% = strTxtSignature %><br /><span class="smText"><% = strTxtSignatureLong %>&nbsp;(max 200 characters)
     <br />
     <br />
     <br />
     <a href="javascript:winOpener('BBcodes.asp<% = strQsSID1 %>','codes',1,1,610,500)" class="smLink"><% = strTxtForumCodes %></a> <% = strTxtForumCodesInSignature %></span><%
     
     	'If rel=nofollow the display a message
     	If blnNoFollowTagInLinks Then Response.Write("<br /><span class=""smText"">" & strTxtNoFollowAppliedToAllLinks & ".</span>")
%></td>
    <td valign="top" height="2">
     <textarea name="signature" id="signature" cols="30" rows="3" onKeyDown="characterCounter('sigChars', 'signature');" onKeyUp="characterCounter('sigChars', 'signature');" tabindex="29"><% = strSignature %></textarea>
     <br />
     <input size="3" value="0" name="sigChars" id="sigChars" maxlength="3" />
     <input onclick="characterCounter('sigChars', 'signature');" type="button" value="<% = strTxtCharacterCount %>" name="Count" />&nbsp;&nbsp;<span class="smText"><a href="javascript:OpenPreviewWindow('signature_preview.asp')" class="smLink"><% = strTxtSignaturePreview %></a>
    </td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtAlwaysAttachMySignature %></td>
    <td><% = strTxtYes %><input type="radio" name="attachSig" value="true" <% If blnAttachSignature = True Then Response.Write "checked" %> tabindex="30" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="attachSig" value="false" <% If blnAttachSignature = False Then Response.Write "checked" %> tabindex="31" /></td>
   </tr><%
		End If

	End If
End If




'*********************************
'****    Forum Preferences    ****
'*********************************

'If part number = 0 (all) or part 3 (forum preferences) then show reg details
If intUpdatePartNumber = 0 OR intUpdatePartNumber = 3 Then

     %>
   <tr class="tableLedger">
    <td colspan="2"><% = strTxtForumPreferences %></td>
   </tr><%

     	'If this is an update and only showing part 3 of the form with no email address entered don't show the 'show email' part of the form
     	If (intUpdatePartNumber = 3 AND strEmail <> "") OR intUpdatePartNumber = 0 Then

     		'If Newsletter is enabled
        	If blnWebWizNewsPad Then
        		%>
   <tr class="tableRow">
    <td width="50%"><% = strTxtNewsletterSubscription %><br /><span class="smText"><% = strTxtSignupToRecieveNewsletters & " " & strWebsiteName %></span></td>
    <td width="50%" valign="top"><% = strTxtYes %><input type="radio" name="newsletter" id="newsletter" value="True" <% If blnNewsletter = True Then Response.Write "checked" %> tabindex="32" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="newsletter" id="newsletter" value="False" <% If blnNewsletter = False Then Response.Write "checked" %> tabindex="33" /></td>
   </tr><%
       		 End If

%>
   <tr class="tableRow">
    <td width="50%"><% = strTxtShowHideEmail %><br /><span class="smText"><% = strTxtShowHideEmailLong %></span></td>
    <td width="50%" valign="top"><% = strTxtYes %><input type="radio" name="emailShow" id="emailShow" value="True" <% If blnShowEmail = True Then Response.Write "checked" %> tabindex="34" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="emailShow" id="emailShow" value="False" <% If blnShowEmail = False Then Response.Write "checked" %> tabindex="35" />
   </td>
   </tr><%

	End If

	'If email notify is on give them a choice to receive mail or not
	If blnEmail = True Then
		%>
   <tr class="tableRow">
    <td width="50%"  class="text"><% = strTxtNotifyMeOfReplies %><br /><span class="smText"><% = strTxtSendsAnEmailWhenSomeoneRepliesToATopicYouHavePostedIn %></span></td>
    <td width="50%" valign="top"><% = strTxtYes %><input type="radio" name="replyNotify" id="replyNotify" value="True" <% If blnReplyNotify = True Then Response.Write "checked" %> tabindex="36" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="replyNotify" id="replyNotify" value="False" <% If blnReplyNotify = False Then Response.Write "checked" %> tabindex="37" /></td>
   </tr><%

        	'If private messageing is also on let them decide if they want to receive email notification when they get em
        	If blnPrivateMessages = True Then
        		%>
   <tr class="tableRow">
    <td width="50%"><% = strTxtNotifyMeOfPrivateMessages %></td>
    <td width="50%" valign="top"><% = strTxtYes %><input type="radio" name="pmNotify" id="pmNotify" value="True" <% If blnPMNotify = True Then Response.Write "checked" %> tabindex="38" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="pmNotify" id="pmNotify" value="False" <% If blnPMNotify = False Then Response.Write "checked" %> tabindex="39" /></td>
   </tr><%
        	End If
	End If

	'If the IE WYSIWYG Editor is on let the user select if they want to use it or not
	If blnRTEEditor = True Then
%>
   <tr class="tableRow">
    <td width="50%"><% = strTxtEnableTheWindowsIEWYSIWYGPostEditor %></td>
    <td width="50%" valign="top"><% = strTxtYes %><input type="radio" name="ieEditor" id="ieEditor" value="True" <% If blnWYSIWYGEditor = True Then Response.Write "checked" %> tabindex="40" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="ieEditor" id="ieEditor" value="False" <% If blnWYSIWYGEditor = False Then Response.Write "checked" %> tabindex="41" /></td>
   </tr><%
	End If

     %>
   <tr class="tableRow">
    <td width="50%"><% = strTxtProfileAutoLogin %><br /><span class="smText"><% = strTxtAutologinOnlyAppliesToSession %></span></td>
    <td width="50%" valign="top"><% = strTxtYes %><input type="radio" name="Login" id="Login" value="True" <% If blnAutoLogin = True Then Response.Write "checked" %> tabindex="42" />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="Login" id="Login" value="False" <% If blnAutoLogin = False Then Response.Write "checked" %> tabindex="43" /></td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtTimezone %><br /><span class="smText"><% = strTxtPresentServerTimeIs %><%

	'Get the current server time
	dtmServerTime = Now()

	'Make sure that the time and date format function isn't effected by the server time off set
	If strTimeOffSet = "-" Then
		dtmServerTime = DateAdd("h", + intTimeOffSet, dtmServerTime)
	ElseIf strTimeOffSet = "+" Then
		dtmServerTime = DateAdd("h", - intTimeOffSet, dtmServerTime)
	End If

	'Display the current server time
	Response.Write(stdDateFormat(dtmServerTime, True) & " " & strTxtAt & " " & TimeFormat(dtmServerTime))

%></span></td>
    <td valign="top">
     <select name="serverOffSet" id="serverOffSet" tabindex="44">
      <option value="+" <% If strTimeOffSet = "+" Then Response.Write("selected") %>>+</option>
      <option value="-" <% If strTimeOffSet = "-" Then Response.Write("selected") %>>-</option>
     </select>
    <select name="serverOffSetHours" tabindex="45"><%

	'Create list of time off-set
	For lngLoopCounter = 0 to 24
		Response.Write(VbCrLf & "      <option value=""" & lngLoopCounter & """")
		If intTimeOffSet = lngLoopCounter Then Response.Write("selected")
		Response.Write(">" & lngLoopCounter & "</option>")
	Next

%>
     </select> <% = strTxtHours %>
    </td>
   </tr>
   <tr class="tableRow">
   <td><% = strTxtDateFormat %></td>
    <td valign="top">
     <select name="dateFormat" tabindex="46">
      <option value="dd/mm/yy" <% If strDateFormat = "dd/mm/yy" Then Response.Write("selected") %>><% = strTxtDayMonthYear %></option>
      <option value="mm/dd/yy" <% If strDateFormat = "mm/dd/yy" Then Response.Write("selected") %>><% = strTxtMonthDayYear %></option>
      <option value="yy/mm/dd" <% If strDateFormat = "yy/mm/dd" Then Response.Write("selected") %>><% = strTxtYearMonthDay %></option>
      <option value="yy/dd/mm" <% If strDateFormat = "yy/dd/mm" Then Response.Write("selected") %>><% = strTxtYearDayMonth %></option>
     </select>
    </td>
   </tr><%

End If




'*********************************************
'****    Admin and Moderator Functions    ****
'*********************************************

'If the admin mode is enabled then place some extra options in the edit profile (unless this is the Guest or Admin accounts)
If blnAdminMode AND (blnAdmin Or blnModerator) Then

     %>
   <tr class="tableLedger">
    <td colspan="2"><a name="admin"></a><% = strTxtAdminModeratorFunctions %></td>
   </tr><%

     	'Don't allow changing group if admin or guest account
     	If lngUserProfileID > 2 Then
     %>
   <tr class="tableRow">
    <td width="50%"><% = strTxtUserIsActive %></td>
    <td width="50%"><% = strTxtYes %><input type="radio" name="active" id="active" value="True" <% If blnUserActive = True Then Response.Write "checked" %> tabindex="47"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="active" id="active" value="False" <% If blnUserActive = False Then Response.Write "checked" %> tabindex="48"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
   </tr>
   <tr class="tableRow">
    <td width="50%"><% = strTxtSuspendUser %></td>
    <td width="50%"><% = strTxtYes %><input type="radio" name="banned" id="banned" value="True" <% If blnSuspended = True Then Response.Write "checked" %> tabindex="49"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />&nbsp;&nbsp;<% = strTxtNo %><input type="radio" name="banned" id="banned" value="False" <% If blnSuspended = False Then Response.Write "checked" %> tabindex="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
   </tr><%

	        'Only allow admin to change the member group
	        If blnAdmin Then


	                'Get the forum groups from the database so admin can change the members group

	                'Initlise SQL query
	                 strSQL = "" & _
	                "SELECT " & strDbTable & "Group.*, " & strDbTable & "LadderGroup.* " & _
			"FROM " & strDbTable & "Group " & _
			"LEFT JOIN " & strDbTable & "LadderGroup ON " & strDbTable & "Group.Ladder_ID = " & strDbTable & "LadderGroup.Ladder_ID " & _
			"ORDER BY " & strDbTable & "LadderGroup.Ladder_Name ASC, " & strDbTable & "Group.Minimum_posts ASC, " & strDbTable & "Group.Group_ID ASC;"
	

	                'Query the database
	                rsCommon.Open strSQL, adoCon

	                'If there are groups then disply them
	                If NOT rsCommon.Eof Then


     %>
   <tr class="tableRow">
    <td><% = strTxtGroup %></td>
    <td>
     <select name="group" id="group" tabindex="51"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>><%

	                        'Loop round to display all the groups
	                        Do While NOT rsCommon.EOF

	                                Dim intSelGroupID
	                                Dim strSelGroupName
	                                Dim blnSelSpecialGroup
	                                Dim lngSelMinimumRankPosts
					Dim strLadderGroup
					 
	                                'Read in the recordset
	                                intSelGroupID = CInt(rsCommon("Group_ID"))
	                                strSelGroupName = rsCommon("Name")
	                                blnSelSpecialGroup = CBool(rsCommon("Special_rank"))
	                                lngSelMinimumRankPosts = CLng(rsCommon("Minimum_posts"))
	                                strLadderGroup = rsCommon("Ladder_Name")

	                                'Display the selection
	                                Response.Write("      <option value=""" & intSelGroupID & """")

	                                'If this is the group the member is part of then have it slected
	                                If intUsersGroupID = intSelGroupID Then Response.Write(" selected")

	                                'Display the end of the select option
	                                If blnSelSpecialGroup Then
	                                        Response.Write(">" & strSelGroupName & "</option>" & vbCrLf)
	                                Else
	                                        Response.Write(">" & strSelGroupName & " - " & strTxtRankLadderGroup & " '" & strLadderGroup & "' " & strTxtMinPosts & " " & lngSelMinimumRankPosts & "</option>" & vbCrLf)
	                                End If

	                                'Move to the next record
	                                rsCommon.MoveNext

	                        Loop
%>
     </select>
    </td>
   </tr><%

                	End If
                End If
	End If
     %>
   <tr class="tableRow">
    <td><% = strTxtMemberTitle %></td>
    <td ><input type="text" name="memTitle" id="memTitle" size="30" maxlength="40" value="<% = strMemberTitle %>" tabindex="52"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtNumberOfPoints %></td>
    <td><input type="text" name="points" id="points" size="4" maxlength="7" value="<% = lngMemberPoints %>" tabindex="53"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
   </tr>
   <tr class="tableRow">
    <td><% = strTxtNumberOfPosts %></td>
    <td><input type="text" name="posts" id="posts" size="4" maxlength="7" value="<% = lngPosts %>" tabindex="54"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
   </tr>
   <tr class="tableRow">
    <td valign="top"><% = strTxtAdminNotes %><br /><span class="smText"><% = strTxtAdminNotesAbout %>.</span></td>
    <td><textarea name="notes" id="notes" cols="30" rows="4" onKeyDown="characterCounter('notesChars', 'notes');" onKeyUp="characterCounter('notesChars', 'notes');" tabindex="55"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>><% = strAdminNotes %></textarea>
    	<br />
     <input size="3" value="0" name="notesChars" id="notesChars" maxlength="3" />
     <input onclick="characterCounter('notesChars', 'notes');" type="button" value="<% = strTxtCharacterCount %>" name="Count" />
    </td>
   </tr><%

End If

'If there is a privacy notice display it
If NOT strRegPrivacyNotice = "" Then
%>
   
    <tr class="tableLedger">
    <td colspan="2"><% = strTxtYourInformation %></td>
   </tr>
   <tr class="tableRow">
    <td colspan="2">
    	<% = strRegPrivacyNotice %>
    </td>
   </tr><%
   
End If
  
%>   
   <tr class="tableBottomRow">
    <td colspan="2" align="center"><%

'If this is admin mode then set the admin stuff up
If blnAdminMode AND (blnAdmin Or blnModerator) Then

        %>
     <input type="hidden" name="M" id="M" value="A" />
     <input type="hidden" name="PF" id="PF" value="<% = lngUserProfileID %>" /><%
End If


	Response.Write(vbCrLf & "     <input type=""hidden"" name=""" & strSessionKey & """ id=""" & strSessionKey & """ value=""" & strFormKey & """ />")

%>
     <input type="hidden" name="mode" id="mode" value="<% = strMode %>" />
     <input type="hidden" name="FPN" id="FPN" value="<% = intUpdatePartNumber %>" />
     <input type="hidden" name="xformID" id="xformID" value="" />
     <input type="submit" name="Submit" id="Submit" value="<% If strMode = "new-reg" Then Response.Write(strTxtRegister) Else Response.Write(strTxtUpdateProfile) %>" onclick="return CheckForm();" tabindex="60" />
     <input type="reset" name="Reset" id="Reset" value="<% = strTxtResetForm %>" tabindex="61" />
    </td>
   </tr>
  </table>
 </form>
<br />
<div align="center"><%

'Release server objects
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
</div><%


'If the username is already gone display an error message pop-up
If blnUsernameOK = False Then
        Response.Write("<script  language=""JavaScript"">")
        Response.Write("alert('" & strTxtUsrenameGone & "');")
        Response.Write("</script>")

End If



'If the email address invalid display error message, display an error message
If blnValidEmail = False Then
        Response.Write("<script  language=""JavaScript"">")
        Response.Write("alert('" & strTxtTheEmailAddressEnteredIsInvalid & "');")
        Response.Write("</script>")
End If

'If the email address is used up and email activation is on, display an error message
If blnEmailOK = False Then
        Response.Write("<script  language=""JavaScript"">")
        Response.Write("alert('" & strTxtEmailAddressAlreadyUsed & "');")
        Response.Write("</script>")
End If

'If the email address or domain is blocked
If blnEmailBlocked Then
        Response.Write("<script  language=""JavaScript"">")
        Response.Write("alert('" & strTxtEmailAddressBlocked & "');")
        Response.Write("</script>")
End If


'If the confirmed password is incorrect
If blnConfirmPassOK = False Then
        Response.Write("<script  language=""JavaScript"">")
        Response.Write("alert('" & strTxtConformOldPassNotMatching & "');")
        Response.Write("</script>")
End If

'If passowrd not complex
If blnPasswordComplexityOK = False Then
	Response.Write("<script  language=""JavaScript"">")
        Response.Write("alert('" & strTxtPasswordNotComplex & "');")
        Response.Write("</script>")
End If

'If the security code did not match
If blnSecurityCodeOK = False Then
        Response.Write("<script  language=""JavaScript"">")
        Response.Write("alert('" & strTxtSecurityCodeDidNotMatch & "');")
        Response.Write("</script>")
End If


'If found at stopforumspam
If blnStopForumSpamFound Then
	Response.Write("<script  language=""JavaScript"">")
        Response.Write("alert('" & strTxtInformationGiveFoundInSpamDatabaseAndRejected & "');")
        Response.Write("</script>")
End If


%>
<!-- #include file="includes/footer.asp" -->