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




'******************************************
'*** 		 Login User		***
'******************************************

'Function to login a user from a form
Private Function loginUser(ByVal strUsername, ByVal strPassword, ByVal blnCAPTCHArequired, strType)


	'Key to login response
	'0 = Login Failed
	'1 = Login OK
	'2 = CAPTCHA Code OK
	'3 = CAPTCHA Code Incorrect
	'4 = CAPTHCA required
	
	
	Dim blnSecurityCodeOK
	Dim lngUserID
	Dim blnActive
	Dim strNewUserCode
	Dim strLagacyPassword
	
	
	'Initilise
	loginUser = 0	'Initilise the login as a fail, changed if all parts correct
	blnSecurityCodeOK = True

	
	'Clean up for SQL
	strUsername = formatSQLInput(strUsername)
	
	'Read in the legacy password format which is lower case
	strLagacyPassword = LCase(strPassword)	
	
	

	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Author.Password, " & strDbTable & "Author.Salt, " & strDbTable & "Author.Username, " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.User_code, " & strDbTable & "Author.Active, " & strDbTable & "Author.Login_attempt, " & strDbTable & "Author.Last_visit, " & strDbTable & "Author.Login_IP " & _
	"FROM " & strDbTable & "Author" & strRowLock & " " & _
	"WHERE " & strDbTable & "Author.Username = '" & strUsername & "';"

	'Set the Lock Type for the records so that the record set is only locked when it is updated
	rsCommon.LockType = 3


	'Set error trapping
	On Error Resume Next
		
	'Query the database
	rsCommon.Open strSQL, adoCon

	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "loginUser()_get_USR_login", "functions_login.asp")
				
	'Disable error trapping
	On Error goto 0	
	
	
	'If a member is returned then checkout the members record to see if they can login
	If NOT rsCommon.EOF Then
		
			
		'Read in the login attempts
		intLoginAttempts = CInt(rsCommon("Login_attempt"))
				
		'Increment login attempts
		intLoginAttempts = intLoginAttempts + 1
				


		'Only encrypt password if this is enabled
		If blnEncryptedPasswords Then
			
			'Encrypt password so we can check it against the encypted password in the database
			'Read in the salt
			strPassword = strPassword & rsCommon("Salt")
	
			'Encrypt the entered password
			strPassword = HashEncode(strPassword)
			
			
			'For backward compatibility with older versions lower case the password
			'Read in the salt
			strLagacyPassword = strLagacyPassword & rsCommon("Salt")
				
			'Encrypt the entered password
			strLagacyPassword = HashEncode(strLagacyPassword)
			
		End If
		


		'Check the encrypted password is correct, if it is get the user ID and set a cookie
		If (strPassword = rsCommon("Password") OR strLagacyPassword = rsCommon("Password")) AND blnSecurityCodeOK Then

			'Only save the user login if CAPTCHA is NOT required, or CAPTCHA is correct
			If (blnSecurityCodeOK AND blnCAPTCHArequired) OR blnCAPTCHArequired = false Then
				
				'Read in the users ID number and whether they want to be automactically logged in when they return to the forum
				lngUserID = CLng(rsCommon("Author_ID"))
				strUsername = rsCommon("Username")
				strNewUserCode = rsCommon("User_code")
				blnActive = CBool(rsCommon("Active"))
				
				'Read in the last login date/time for this user
				If isDate(rsCommon("Last_visit")) Then 
					dtmLastVisitDate = CDate(rsCommon("Last_visit"))
					Call saveSessionItem("LV", internationalDateTime(dtmLastVisitDate))
				End If
				
				
				
				'Set error trapping
				On Error Resume Next
					
				'For extra security create a new user code for the user if this feature is enabled
				If blnNewUserCode AND blnWindowsAuthentication = False Then
					
					'Create new user code for user
					If blnActive Then strNewUserCode = userCode(strUsername)
					
					'Save the new usercode back to the database and reset login attempts
					If blnActive Then rsCommon.Fields("User_code") = strNewUserCode 'Only do this if the users account is active, otherwise their activation email will fail
				End If	
				
				'Reset login count	
				rsCommon.Fields("Login_attempt") = 0
				
				'Save login IP
				rsCommon.Fields("Login_IP") = Trim(Mid(getIP(), 1, 50))
				
				'Update the database
				rsCommon.Update
					
				'If an error has occurred write an error to the page
				If Err.Number <> 0 Then	Call errorMsg("An error has occurred while writing to the database.", "loginUser()_update_USR_Code", "functions_login.asp")
						
				'Disable error trapping
				On Error goto 0
				
				
				
				'Save the users login ID to the session variable
				Call saveSessionItem("UID", strNewUserCode)
				
				

				
				'If logging in an admin save an admin session code
				If strType = "admin" Then 
					
					Call saveSessionItem("AID", strNewUserCode)
				
				'If not an admin section login update the Anonymous cookie
				Else
					'Save to session if the user is browsing annonymously, 1 = Anonymous, 0 = Shown
					If BoolC(Request.Form("NS")) = False Then
						Call saveSessionItem("NS", "1")
					Else
						Call saveSessionItem("NS", "0")
					End If
				End If
				
				
				
				'If the user has selected auto login set a cookie for the user on their machine
				If blnAutoLogin Then
					
					'Write a login cookie to keep the user logged in
					Call setCookie("sLID", "UID", strNewUserCode, True)
					
					'If not admin mode update annoymous user
					If NOT strType = "admin" Then
						'Write a cookie saying if the user is browsing anonymously, 1 = Anonymous, 0 = Shown
						If BoolC(Request.Form("NS")) = False Then
							Call setCookie("sLID", "NS", "1", True)'Anonymous
						Else
							Call setCookie("sLID", "NS", "0", True)'Shown
						End If
					End If
				
				'Else non auto login
				Else 
					'Write a login cookie to prevent users having issues being logged out on bad servers
					Call setCookie("sLID", "UID", strNewUserCode, False)
				
				End If
				
				
				'Set the login response to 1 for OK
				loginUser = 1
			End If
			
			
		
		'Else the login was incorrect
		Else
			
			'Set error trapping
			On Error Resume Next
			
			'Update the login attempts in the database
			rsCommon.Fields("Login_attempt") = intLoginAttempts
			
			'For extra security create a new user code (auto-login code) for the user if more than the set un-sucessful login attempts
			'This should make it harder for a hacker if they are attempting mutiple methods of gaining control of an account
			'It will also mean the account holder is forced to log back in again, so they will then be informed when logging in of the login attempts on their account which will alert them to the presence of the attempt on their account
			If intLoginAttempts => intIncorrectLoginAttempts AND blnWindowsAuthentication = False Then rsCommon.Fields("User_code") = userCode(strUsername)
			
			'Update the database
			rsCommon.Update
			
			'If an error has occurred write an error to the page
		  	If Err.Number <> 0 Then	Call errorMsg("An error has occurred while writing to the database.", "loginUser()_update_login_attempts", "functions_login.asp")
				
			'Disable error trapping
			On Error goto 0
			
			'If the CAPTCHA check has failed inform the user, but for extra security don't let them know if the login has failed
			If blnSecurityCodeOK = False Then
				loginUser = 3
			
			'Else let the user know the login failed
			Else
				loginUser = 0
			End If
		End If
	End If

	'Reset Server Objects
	rsCommon.Close
	
	
	'Clear out any read/unread posts and start again for this user now they have logged in as we can now use their last visit date from their account
	If loginUser = 1 Then
		Application("sarryUnReadPosts" &  strSessionID) = ""
		Session("sarryUnReadPosts") = ""
		Session("dtmUnReadPostCheck") = ""
				
		'Get read/unread posts
		Call UnreadPosts()
	End If
	
End Function










'******************************************
'*** 		Get User Data		***
'*****************************************



'Sub procedure to get the user data for users
Public Sub getUserData(ByVal strSessionItem)


	'Read in user ID from the application session
	strLoggedInUserCode = getSessionItem(strSessionItem)
	
	
	'Read in users ID number from the auto login cookie (if not an admin area ID)
	If strLoggedInUserCode = "" AND strSessionItem = "UID" Then strLoggedInUserCode = Trim(Mid(getCookie("sLID", "UID"), 1, 44))

	
	
	'If the member API is enabled log the user in from an existing member login system
	If blnMemberAPI Then
		'If username stored in the session has changed run the API Login again as this is a different user
		If Session("ForumUSER") <> Session("USER") AND strSessionItem = "UID" Then strLoggedInUserCode = existingMemberAPI()
	End If
	
	'If windows authentication is enabled then log the user in using from windows
	If blnWindowsAuthentication AND strLoggedInUserCode = "" AND strSessionItem = "UID" Then strLoggedInUserCode = windowsAuthentication()
	
	
	
	'If a cookie exsists on the users system then read in there username from the database
	If NOT strLoggedInUserCode = "" Then
	
		'Make the usercode SQL safe
		strLoggedInUserCode = formatSQLInput(strLoggedInUserCode)
	
		'Initalise the strSQL variable with an SQL statement to query the database
		strSQL = "SELECT " & strDbTable & "Author.Username, " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.Group_ID, " & strDbTable & "Author.Active, " & strDbTable & "Author.Signature, " & strDbTable & "Author.Author_email, " & strDbTable & "Author.Date_format, " & strDbTable & "Author.Time_offset, " & strDbTable & "Author.Time_offset_hours, " & strDbTable & "Author.Reply_notify, " & strDbTable & "Author.Attach_signature, " & strDbTable & "Author.Rich_editor, " & strDbTable & "Author.Last_visit, " & strDbTable & "Author.No_of_PM, " & strDbTable & "Author.Inbox_no_of_PM, " & strDbTable & "Author.Banned, " & strDbTable & "Group.Image_uploads, " & strDbTable & "Group.File_uploads, " & strDbTable & "Group.Signatures, " & strDbTable & "Group.URLs, " & strDbTable & "Group.Images, " & strDbTable & "Group.Private_Messenger, " & strDbTable & "Group.Chat_Room " & _
		"FROM " & strDbTable & "Author" & strDBNoLock & ", " & strDbTable & "Group" & strDBNoLock & " " & _
		"WHERE " & strDbTable & "Author.Group_ID  = " & strDbTable & "Group.Group_ID " & _
			"AND " & strDbTable & "Author.User_code = '" & strLoggedInUserCode & "';"
	
		'Set error trapping
		On Error Resume Next
	
		'Query the database
		rsCommon.Open strSQL, adoCon
				
		'If an error has occurred write an error to the page
		If Err.Number <> 0 Then	Call errorMsg("An error has occurred while reading user data from the database.", "getUserData()_get_member_data", "functions_login.asp")
			
		'Disable error trapping
		On Error goto 0
	
		'If the database has returned a record then run next bit
		If NOT rsCommon.EOF Then
	
			'Read in the users details from the recordset
			strLoggedInUsername = rsCommon("Username")
			intGroupID = rsCommon("Group_ID")
			lngLoggedInUserID = CLng(rsCommon("Author_ID"))
			blnActiveMember = CBool(rsCommon("Active"))
			strDateFormat = rsCommon("Date_format")
			strTimeOffSet = rsCommon("Time_offset")
			intTimeOffSet = CInt(rsCommon("Time_offset_hours"))
			blnReplyNotify = CBool(rsCommon("Reply_notify"))
			blnAttachSignature = CBool(rsCommon("Attach_signature"))
			blnWYSIWYGEditor = CBool(rsCommon("Rich_editor"))
			strLoggedInUserEmail = rsCommon("Author_Email")
			If NOT isNull(rsCommon("No_of_PM")) Then intNoOfPms = CInt(rsCommon("No_of_PM")) Else intNoOfPms = 0  
			If NOT isNull(rsCommon("Inbox_no_of_PM")) Then intNoOfInboxPms = CInt(rsCommon("Inbox_no_of_PM")) Else intNoOfInboxPms = 0  	
			If isDate(rsCommon("Last_visit")) Then dtmUserLastVisitDate = CDate(rsCommon("Last_visit")) Else dtmUserLastVisitDate = Now()
			If rsCommon("Signature") <> Trim("") Then blnLoggedInUserSignature = True
			blnBanned = CBool(rsCommon("Banned"))
			blnAttachments = CBool(rsCommon("File_uploads"))
			blnImageUpload = CBool(rsCommon("Image_uploads"))
			blnGroupSignatures = CBool(rsCommon("Signatures")) 
			blnGroupURLs = CBool(rsCommon("URLs"))
			blnGroupImages = CBool(rsCommon("Images"))
		
			
			'If private messages are enabled see if the member can use it or not
			If blnPrivateMessages Then
				If CBool(rsCommon("Private_Messenger")) = False Then blnPrivateMessages = False
			End If
			
			'If chat room is enabled see if the member can use it or not
			If blnChatRoom Then
				If CBool(rsCommon("Chat_Room")) = False Then blnChatRoom = False
			End If
			
			If blnDemoMode Then
				blnAttachments = True
				blnImageUpload = True
			
			End If
			
			'See if the user has entered an email address
			If strLoggedInUserEmail <> Trim("") Then blnLoggedInUserEmail = True
	
	
			'Read in the Last Visit Date for the user from the db if we haven't already
			If dtmUserLastVisitDate > dtmLastVisitDate OR dtmLastVisitDate = "" Then
				dtmLastVisitDate = dtmUserLastVisitDate
				Call saveSessionItem("LV", internationalDateTime(dtmUserLastVisitDate))
			End If
	
			
				
			'If the Last Visit date in the db date is older than 1 minutes for the user then update it
			'Set to every 1 minutes to save on the number of db updates required
			If dtmUserLastVisitDate < DateAdd("n", -1, Now()) Then
	
				'Initilse sql statement
			 	strSQL = "UPDATE " & strDbTable & "Author" & strRowLock & " " & _
				"SET " & strDbTable & "Author.Last_visit = " & formatDbDate(Now()) & _
				"WHERE " & strDbTable & "Author.Author_ID = " & lngLoggedInUserID & ";"
				
				'Set error trapping
				On Error Resume Next
	
				'Write to database
				adoCon.Execute(strSQL)
				
				'If an error has occurred write an error to the page
				If Err.Number <> 0 Then	Call errorMsg("An error has occurred while writing to the database.", "getUserData()_update_last_visit", "functions_login.asp")
			
				'Disable error trapping
				On Error goto 0
	
			End If
	
			'If the members account is not active or suspended then set there group to 2 (Guest Group)
			If blnActiveMember = False OR blnBanned Then intGroupID = 2
	
			'Set the Guest boolean to false
			blnGuest = False
		End If
	
		'Clean up
		rsCommon.Close
	End If
	
	
	'Call the sub for read/unread posts now so that it uses the last visit date from the database
	If Session("dtmUnReadPostCheck") = "" Then Call UnreadPosts()
End Sub
%>