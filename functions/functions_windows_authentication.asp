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




'This file is used if you want to use Windows Authentication/Active Directory to log users into Web Wiz Forums.

'This file will auto check if the user is already in the forums database, if not a new 
'member will be created for the windows authenticated user.

'The members username in the forum will use the name part of the windows authentication username,
'if the forum admin would like to change the members name to something different, they can log
'into the forums online 'control panel' and the 'Change Username' tool to change the members name


'Admin Control Panel - IMPORTANT
'The admin control panel can only be accessed through the built in admin account once Windows Authentication/Active Directory
'is enabled. You would need to point your browser at the 'admin.asp' page and login with the built in admin account credentials.
'Other member accounts can NOT be set as admin accounts once enabled, if you wish to give users extra powers you would need to
'make them moderators. 'Do NOT rename the built in admin account with the same name as your Windows Authentication/Active Directory 
'login name.


'IMPORTANT
'Windows Authentication/Active Directory MUST be enabled from a clean install you can NOT change to this type of login at a later stage.
'Do NOT try and import members or add them manually, you MUST let the users be auto added to your forum by the Web Wiz Forums software.



'*** PLEASE NOTE, THIS FEATURE IS NOT AVAILABLE IN THE FREE EXPRESS EDITION ****



'Set the variable below to true if you are using windows authentication
blnWindowsAuthentication = False



'Sub to login and create member accounts when using windows authentication
Private Function windowsAuthentication()


	Dim strAuthenticatedUser
	Dim strAuthenticatedPass
	Dim strUserName
	Dim strSQL
	Dim intForumStartingGroup
	Dim strSalt


	'Get windows authentcated username
	strAuthenticatedUser = Request.ServerVariables("AUTH_USER")
	
	'If the method above fails use the following
	If strAuthenticatedUser = "" Then strAuthenticatedUser = Request.ServerVariables("LOGON_USER")
		
	'Get the user password
	strAuthenticatedPass = Request.ServerVariables("AUTH_PASSWORD")
	
	
	
	'If the windows authenticated username has not been passed across display an error message
	If strAuthenticatedUser = "" Then
		Call errorMsg("An error has occurred while reading Authenticated User's, username from Windows Server.<br />Please check that are using a Windows Authenticated Login System and you are NOT browsing this site anonymously.", "windowsAuthentication()_get_AUTH_USER", "functions_windows_authentication.asp")
	End If
	
	
	
	'Create a salt value
	strSalt = getSalt(5)
	
	'Only encrypt password if this is enabled
	If blnEncryptedPasswords AND strAuthenticatedPass <> "" Then
		
		'Encrypt the entered password
		strAuthenticatedPass = HashEncode(strAuthenticatedPass & strSalt)
	End If
	
	
	'If there is no password for the user then just place in a random value
	If strAuthenticatedPass = "" Then strAuthenticatedPass = LCase(hexValue(10))
	
	
	
	
	'See if user is in forum db by looking for the windows authentication ID in the User_code field
	strSQL = "SELECT " & strDbTable & "Author.User_code " & _
	"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Author.User_code = '" & formatSQLInput(strAuthenticatedUser) & "';"
	
	'Set error trapping
	On Error Resume Next
	
	'Query the database
	rsCommon.Open strSQL, adoCon
				
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("An error has occurred while reading data from the database.", "windowsAuthentication()_get_member_data", "functions_windows_authentication.asp")
			
	'Disable error trapping
	On Error goto 0
		
	
	'If not in forum db add user to db
	If rsCommon.EOF AND strAuthenticatedUser <> "" Then
		
		'Close rs
		rsCommon.Close
		
		'Use the last part of the windows authentication (bit without domain) as the forum username
		If InStrRev(strAuthenticatedUser, "\") = 0 Then
			strUserName = strAuthenticatedUser
		Else
			strUserName = Mid(strAuthenticatedUser, InStrRev(strAuthenticatedUser, "\")+1, Len(strAuthenticatedUser))
		End If
		
		
		
		'Check the user is not already in the database to prevent crashes later
		strSQL = "SELECT " & strDbTable & "Author.Username " & _
		"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
		"WHERE " & strDbTable & "Author.Username = '" & formatSQLInput(strUserName) & "';"
		
		'Set error trapping
		On Error Resume Next
		
		'Query the database
		rsCommon.Open strSQL, adoCon
		
		'If an error has occurred write an error to the page
		If Err.Number <> 0 Then	Call errorMsg("An error has occurred while reading data from the database.", "windowsAuthentication()_get_check_username", "functions_windows_authentication.asp")
				
		'Disable error trapping
		On Error goto 0
		
		'If the user is in the database already then change the username to the AD domain\username
		If NOT rsCommon.EOF Then strUserName = strAuthenticatedUser
	
		'Close rs
		rsCommon.Close
		
		
		
		
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
		"'" & formatSQLInput(strUserName) & "', " & _
		"'" & formatSQLInput(strAuthenticatedUser) & "', " & _
		"'" & formatSQLInput(strAuthenticatedPass) & "', "  & _
		"'" & strSalt & "', " & _
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
		If Err.Number <> 0 Then	Call errorMsg("An error has occurred while writing to the database.", "windowsAuthentication()_add_new_user", "functions_windows_authentication.asp")
			
		'Disable error trapping
		On Error goto 0
	
	Else
		rsCommon.Close
	End If
	
	
	'Login user
	
	'Create a forum session for the user to keep them logged
	Call saveSessionItem("UID", strAuthenticatedUser)
	Call saveSessionItem("NS", "0")
	
	'Write auto login cookie (can improve tracking for some browsers)
	Call setCookie("sLID", "UID", strAuthenticatedUser, True)
	Call setCookie("sLID", "NS", "0", True)
	
	
	'Return function
	windowsAuthentication = strAuthenticatedUser

End Function
%>