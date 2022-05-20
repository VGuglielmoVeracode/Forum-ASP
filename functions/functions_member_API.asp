<!--#include file="functions_hash1way.asp" -->
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






'This file is used if you want to use an existing site login system to log users into Web Wiz Forums.



'****************************************************************************************

'How the Web Wiz Forums Member API works
'---------------------------------------

'By setting the session variables below they are passed to Web Wiz Forums login API when 
'the member enters the forum. If the member is not already in the Web Wiz Forums own database 
'they are automatically entered if they are already in the database they are logged in with 
'that members details, if their password or email address has changed since their last visit 
'to the forum this will also be updated to keep the forum in sync with your own websites 
'login system.

'The Member API relies on the security of your own login system, if your login system is
'not secure then nor will Web Wiz Forums.

'****************************************************************************************



'****************************************************************************************

'To use the Web Wiz Forums member API in your existing member login system you need to 
'set the following session variables for the members username, password, email address 
'so that they can be passed across to Web Wiz Forums to login in your member

'	Session("USER") = Member_Username
'	Session("PASSWORD") = Member_Password
'	Session("EMAIL") = Member_Email

'Replace 'Member_Username' with the username of your member
'Replace 'Member_Password' with the password of your member
'Replace 'Member_Email' with the password of your member (this one is optional)

'****************************************************************************************



'Enable Member API
'-----------------
'Set the variable below to 'True' if you are using the member API to log users in from your 
'own websites login system.

Const blnMemberAPI = False



'Set auto-login cookie
'------------------------------
'Set this veriable to true if you want an auto-login cookie set on the users browser, so when they 
'return to the forum at a later date they are automatically logged into the forum

Const blnMemberAPIautoLoginCookie = False




'Disable Forums Account Control
'------------------------------
'Set the variable below to 'True' if you wish to disable the forums own login, logout, 
'registration system as well as password and email address changes.
'This is useful if you want to control membership and account settings through your own system.

Const blnMemberAPIDisableAccountControl = False



'Login from your own website
'---------------------------
'Enter the URL below between the "" quotes to your web sites login page. When the user clicks 
'the 'login' link in the forum they will be taken to your own websites login page. 
'Use the full URL eg. "http://www.myweb.com/login.htm"

Const strMemberAPILoginURL = ""



'Register on your own website
'----------------------------
'Enter the URL below between the "" quotes to your websites registration pages. When the 
'users clicks on the 'register' link in the forum they will be taken to your own websites 
'registration pages.
'Use the full URL eg. "http://www.myweb.com/register.htm"

Const strMemberAPIRegistrationURL = ""



'Logout from the forum and your own website
'------------------------------------------
'Enter the URL below between the "" quotes to your web sites logout page. When the user clicks 
'the 'logout' link in the forum they will be taken to your own websites logout page. 
'Use the full URL eg. "http://www.myweb.com/logout.htm"

Const strMemberAPILogoutURL = ""

'****************************************************************************************





'The code below logs the user in when using the member API, unless you are an advanced user
'or developer do not change the code below

'Sub procedure to login and create accounts on Web Wiz Forums when the member API is enabled
Private Function existingMemberAPI()

	Dim strUsername
	Dim strPassword
	Dim strPasswordCheck
	Dim strEmail
	Dim intForumStartingGroup
	Dim strSalt
	Dim strNewUserCode
	Dim lngUserID
	Dim blnActive


	'Get windows authentcated username
	strUsername = Session("USER")
	Session("ForumUSER") = Session("USER")

	'Get the user password
	strPassword = Trim(Session("PASSWORD"))
	
	'Get the user email
	strEmail = LCase(Trim(Session("EMAIL")))
	
	
	'Exit function if no username and password for the user
	If strUsername = "" Then Exit Function
	
	
	'Check for SQL injections
	strUsername = formatSQLInput(strUsername)
	
	
	'Read in the user data from database if it exists
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Author.Password, " & strDbTable & "Author.Salt, " & strDbTable & "Author.Username, " & strDbTable & "Author.Author_email, " & strDbTable & "Author.Author_ID, " & strDbTable & "Author.User_code, " & strDbTable & "Author.Active, " & strDbTable & "Author.Login_attempt " & _
	"FROM " & strDbTable & "Author" & strRowLock & " " & _
	"WHERE " & strDbTable & "Author.Username = '" & strUsername & "';"

	'Set the Lock Type for the records so that the record set is only locked when it is updated
	rsCommon.LockType = 3


	'Set error trapping
	On Error Resume Next
		
	'Query the database
	rsCommon.Open strSQL, adoCon

	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "existingMemberAPI()_get_USR_login", "functions_member_API.asp")
				
	'Disable error trapping
	On Error goto 0	
	
	
	'If NOT EOF then login the user
	If NOT rsCommon.EOF Then
		
		
		'Read in the users ID number and whether they want to be automactically logged in when they return to the forum
		lngUserID = CLng(rsCommon("Author_ID"))
		strUsername = rsCommon("Username")
		strNewUserCode = rsCommon("User_code")
		blnActive = CBool(rsCommon("Active"))
		
		
		
		
		'Only encrypt password if this is enabled
		If blnEncryptedPasswords Then
			
			'Read in the salt value from the database
			strSalt = rsCommon("Salt")
				
			'Encrypt password so we can check it against the encypted password in the database
			'Read in the salt
			strPasswordCheck = strPassword & strSalt
		
			'Encrypt the entered password
			strPasswordCheck = HashEncode(strPasswordCheck)
		
		'Else the password is not encrypted
		Else
			strPasswordCheck = strPassword
		End If
	
		
		'If the password doest match that on record we need to create a new password to save to db
		If NOT strPasswordCheck = rsCommon("Password") Then 
	
			'Create a salt value
			strSalt = getSalt(5)
			
			'Only encrypt password if this is enabled
			If blnEncryptedPasswords Then
				
				'Encrypt the entered password
				strPassword = HashEncode(strPassword & strSalt)
			End If
		End If
		
		
		
		'For extra security create a new user code for the user
		strNewUserCode = userCode(strUsername)
	
		'Set error trapping
		On Error Resume Next
				
		'Save the new usercode back to the database and the password incase it has been changed
		rsCommon.Fields("User_code") = strNewUserCode
		rsCommon.Fields("Password") = strPassword
		rsCommon.Fields("Salt") = strSalt
		If strEmail <> "" Then rsCommon.Fields("Author_email") = strEmail
		rsCommon.Update
				
		'If an error has occurred write an error to the page
		If Err.Number <> 0 Then	Call errorMsg("An error has occurred while writing to the database.", "existingMemberAPI()_update_pass", "functions_member_API.asp")
					
		'Disable error trapping
		On Error goto 0
	
		'Close recordset
		rsCommon.Close
	
	
	
	
	
	
	'Else the user is not in the database so they must be a new user
	Else
	
		'Close recordset
		rsCommon.Close
		
		
		'Create a usercode for the new user
		strNewUserCode = userCode(strUsername)
		
		'Create a salt value
		strSalt = getSalt(5)
		
		'If no password then create one for the user
		If strPassword = "" Then hexValue(10)
		
		'Only encrypt password if this is enabled
		If blnEncryptedPasswords Then
			
			'Encrypt the entered password
			strPassword = HashEncode(strPassword & strSalt)
		
		
		'Else the password is not encrypted, but check for SQL injections
		Else 
			strPassword = formatSQLInput(strPassword)
		End If
		
		 'Check for SQL injections
		strEmail = formatSQLInput(strEmail)
		
		
		
		
		'We need to get the start user group ID from the database
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
                
               
		
		
		
		'Create SQL to insert new user in database
		strSQL = "INSERT INTO " & strDbTable & "Author (" & _
		"Group_ID, " & _
		"Username, " & _
		"User_code, "
		If strDatabaseType = "mySQL" Then strSQL = strSQL & "Password, " Else strSQL = strSQL &  "[Password], "
		strSQL = strSQL & "Salt, " & _
		"Author_email, " & _
		"Show_email, " & _
		"Attach_signature, " & _
		"Time_offset, " & _
		"Time_offset_hours, " & _
		"Rich_editor, " & _
		"Date_format, " & _
		"Active, " & _
		"Reply_notify, " & _
		"PM_notify, " & _
		"No_of_posts, " & _
		"Signature, " & _
		"Join_date, " & _
		"Last_visit, " & _
		"Login_attempt, " & _
		"Banned, " & _
		"Info " & _
		") " & _
		"VALUES " & _
		"('" & intForumStartingGroup & "', " & _
		"'" & strUserName & "', " & _
		"'" & strNewUserCode & "', " & _
		"'" & strPassword & "', "  & _
		"'" & strSalt & "', " & _
		"'" & strEmail & "', " & _
		strDBFalse & ", " & _
		strDBFalse & ", " & _
		"'" & strTimeOffSet & "', " & _
		"'" & intTimeOffSet & "', " & _
		strDBTrue & ", " & _
		"'dd/mm/yy', " & _
		strDBTrue & ", " & _
		strDBFalse & ", " & _
		strDBFalse & ", " & _
		"'0', " & _
		"'', " & _
		strDatabaseDateFunction & ", " & _
		strDatabaseDateFunction & ", " & _
		"'0', " & _
		strDBFalse & ", " & _
		"'');"
		
			
		'Set error trapping
		On Error Resume Next
	
		'Write to database
		adoCon.Execute(strSQL)
				
		'If an error has occurred write an error to the page
		If Err.Number <> 0 Then	Call errorMsg("An error has occurred while writing to the database.", "existingMemberAPI()_add_new_user", "functions_member_API.asp")
			
		'Disable error trapping
		On Error goto 0
	
	End If
	
	
	
	
	'Save the users login ID to the session variable
	Call saveSessionItem("UID", strNewUserCode)
	Call saveSessionItem("NS", "0")
		
	'Write auto login cookie (set for 1 year)
	If blnMemberAPIautoLoginCookie Then
		Call setCookie("sLID", "UID", strNewUserCode, True)
		Call setCookie("sLID", "NS", "0", True)
	End If
	
	'Return function
	existingMemberAPI = strNewUserCode
End Function
%>