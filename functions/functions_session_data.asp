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
'***  Read in session application data	***
'******************************************

'Sub procedure to read in the session data from application variable or create new one if it doesn't exist
Private Sub getSessionData()



	Dim intSessionArrayPass
	Dim intLastArrayPostionPointer
	Dim strNewSessionID
	Dim blnCookiesDetected
	Dim blnFoundSession
	Dim strIP
	Dim strDate
	Dim intArraySize
	Dim strNewSessionIDKey
	Dim sarryStaleSessions
	Dim intStaleSessionLoop


	'Initialise  variables
	
	intLastArrayPostionPointer = 0
	blnCookiesDetected = false
	blnFoundSession = false
	strNewSessionID = createForumSessionID()
	strQsSID = ""
	strQsSID1 = ""
	strQsSID2 = ""
	strQsSID3 = ""


	'If Session IP checking is enabled then include the IP address with the session
	If blnSessionIpCheck Then
		'Get IP Address
		strIP = getIP()
		
		
		'If IPv6 then use the network block address /48
		If InStr(strIP, ":") Then
			strIP = IPv6NetBlock(strIP)
			
		'Use only the first 2 parts of the IP address to prevent errors when using mutiple proxies (eg. AOL uses)
		Else
			strIP = Mid(strIP, 1, (InStr(InStr(1, strIP, ".", 1)+1, strIP, ".", 1)))
		End If
	'Else IP checking is disabled
	Else
		strIP = " "
	End If


	'Read in the session ID, if available, 
	'Use cookies (check before querystring)
	If getCookie("sID", "SID") <> "" Then
		'Set the cookie detection variable to true
		blnCookiesDetected = true

		'Get the session ID from cookie
		strSessionID = LCase((Trim(getCookie("sID", "SID"))))

	'Else if no cookies, or cookies not working use querystrings
	ElseIf Request.QueryString("SID") <> "" Then

		'Get the session ID from querystring
		strSessionID = LCase(Trim(Request.QueryString("SID")))
	End If



	'Session array lookup table
	'0 = Session ID
	'1 = IP address
	'2 = Time last accessed
	'3 = Session data
	
	
	
	'*******************************
	'*** Database Held Sessions ****
	'*******************************
	
	'Read in the session data from the database
	If blnDatabaseHeldSessions Then
		
		'Get all sssion data from database
		strSQL = "SELECT " & strDbTable & "Session.Session_ID, " & strDbTable & "Session.IP_address, " & strDbTable & "Session.Last_active, " & strDbTable & "Session.Session_data " & _
		"FROM " & strDbTable & "Session" & strDBNoLock & ";"
		
		'Set error trapping
		On Error Resume Next
	
		'Get recordset
		rsCommon.Open strSQL, adoCon
				
		'If an error has occurred write an error to the page
		If Err.Number <> 0 Then Call errorMsg("An error has occurred while executing SQL query on database.", "getSessionData()_get_session_data", "functions_session_data.asp")
			
		'Disable error trapping
		On Error goto 0
		
		
		'If records returned place them in the array
		If NOT rsCommon.EOF Then
			sarySessionData = rsCommon.GetRows()
		'Else create an array
		Else
			ReDim sarySessionData(3,0)
		End If
		
		'Close recordset
		rsCommon.Close
		
		
		'Get array size using isArray to prevent errors.
		If isArray(sarySessionData) Then 
			intArraySize = CLng(UBound(sarySessionData, 2))
			intLastArrayPostionPointer = intArraySize
		Else
			intArraySize = 0
		End If
		
		'Iterate through array
		For intSessionArrayPass = 0 To intArraySize
		
			'If user has a session, read in data, and update last access time (for security we also check the IP address)
			If sarySessionData(0, intSessionArrayPass) = strSessionID AND sarySessionData(1, intSessionArrayPass) = strIP Then
	
				'If using a database for session data we need to update the last access time in the database
				'Only update if older than 3 minutes to cut down on database hits (this date is also updated when saving session data, reducing the amount of times it needs to be updated)
				If CDate(sarySessionData(2, intSessionArrayPass)) < DateAdd("n", -3, Now()) Then
				
					'Initilse sql statement
				 	strSQL = "UPDATE " & strDbTable & "Session" & strRowLock & " " & _
					"SET " & strDbTable & "Session.Last_active = " & formatDbDate(Now()) & " " & _
					"WHERE " & strDbTable & "Session.Session_ID = '" & sarySessionData(0, intSessionArrayPass) & "';"
					
					'Set error trapping
					On Error Resume Next
		
					'Write to database
					adoCon.Execute(strSQL)
					
					'If an error has occurred write an error to the page
					If Err.Number <> 0 Then	Call errorMsg("An error has occurred while writing to the database.", "getSessionData()_update_last_active_date", "functions_session_data.asp")
				
					'Disable error trapping
					On Error goto 0
					
				End If
				
				
				'Set blnFoundSession to true
				blnFoundSession = true
	
				'Read in session data
				strSessionData = sarySessionData(3, intSessionArrayPass)
				
				'Update last access time
				sarySessionData(2, intSessionArrayPass) = internationalDateTime(Now())
			End If
			

			'If creating a new session for the user and the session ID already exists create a new one
			If strNewSessionID = sarySessionData(0, intSessionArrayPass) Then
				strNewSessionID = createForumSessionID()
				intSessionArrayPass = 0
			End If
			
		Next
			
		
		
			
		'Remove stale read/unread session data from OS memory
		
		'SQL to get stale sessions from the database
		strSQL = "SELECT " & strDbTable & "Session.Session_ID " & _
		" FROM " & strDbTable & "Session" & strDBNoLock & " " & _
		"WHERE " & strDbTable & "Session.Last_active < " & formatDbDate(DateAdd("n", -intSessionTimeout, Now())) & ";"
	
		'Set error trapping
		On Error Resume Next
	
		'Open RS
		rsCommon.Open strSQL, adoCon
				
		'If an error has occurred write an error to the page
		If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database", "getSessionData()_get_stale_session_data", "functions_session_data.asp")
				
		'Disable error trapping
		On Error goto 0 
		
		'If records returned place them in the array
		If NOT rsCommon.EOF Then
			
			'Place rs into array
			sarryStaleSessions = rsCommon.GetRows()
		
			'Loop through the stale sessions to remove 
			For intStaleSessionLoop = 0 to UBound(sarryStaleSessions, 2)
				
				'Remove unread/read array for this session
				Application("sarryUnReadPosts" &  sarryStaleSessions(0, intStaleSessionLoop)) = ""
				
				'SQL to delete stale entries from the database
				strSQL = "DELETE FROM " & strDbTable & "Session" & strRowLock & " " & _
				"WHERE " & strDbTable & "Session.Session_ID = '" & sarryStaleSessions(0, intStaleSessionLoop) & "';"
				
				'Set error trapping
				On Error Resume Next
				
				'Execute SQL
				adoCon.Execute(strSQL)
						
				'If an error has occurred write an error to the page
				If Err.Number <> 0 Then	Call errorMsg("An error has occurred while writing to the database.", "getSessionData()_delete_session_data", "functions_session_data.asp")
					
				'Disable error trapping
				On Error goto 0 
			Next
			
		End If	
		
		'Close RS
		rsCommon.Close	
		
		
		
		
		
		'If the user doesn't have a session create one
		If blnFoundSession = false Then
			
			'Create an ID key for the session
			strNewSessionIDKey = LCase(hexValue(12))
		
			'Get the array size
			intArraySize = CInt(UBound(sarySessionData, 2))
			
			'Increarment the array size
			intArraySize = intArraySize + 1
	
			'ReDimesion the array
			ReDim Preserve sarySessionData(3, intArraySize)
	
			'Update the new array position with the new session details
			sarySessionData(0, intArraySize) = strNewSessionID
			sarySessionData(1, intArraySize) = strIP
			sarySessionData(2, intArraySize) = internationalDateTime(Now())
			sarySessionData(3, intArraySize) = "KEY=" & strNewSessionIDKey & ";"
	
			'Initilise the session id variable
			strSessionID = strNewSessionID
			strSessionData = "KEY=" & strNewSessionIDKey & ";"
	
			'Create a cookie and querystring with the session ID
			Call setCookie("sID", "SID", strNewSessionID, False)
		
		
		
		
			'SQL to update the database with the new session
		 	strSQL = "INSERT INTO " & strDbTable & "Session (" &_
			"Session_ID, " & _
			"IP_address, " & _
			"Last_active, " & _
			"Session_data " & _
			") " & _
			"VALUES " & _
			"('" & strNewSessionID & "', " & _
			"'" & strIP & "', " & _
			formatDbDate(Now()) & ", " & _
			"'KEY=" & strNewSessionIDKey & ";' " & _
			");"
			
			'Set error trapping
			On Error Resume Next

			'Write to database
			adoCon.Execute(strSQL)
			
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then	Call errorMsg("An error has occurred while writing to the database.", "getSessionData()_save_new_session_data", "functions_session_data.asp")
		
			'Disable error trapping
			On Error goto 0
		
		End If
			
			
		
		
		
	
	
	
	
	
	'***********************************
	'*** App Veriable Held Sessions ****
	'***********************************
	
	'Read in the session array from the application variable	
	Else
		
		'Lock the application so that no other user can try and update the application level variable at the same time
		Application.Lock
		
		'Read in the session from the application
		If isArray(Application(strAppPrefix & "sarySessionData")) Then
			sarySessionData = Application(strAppPrefix & "sarySessionData")
	
		'Else create an array
		Else
			ReDim sarySessionData(3,0)
		End If

		'Get array size using isArray to prevent errors.
		If isArray(sarySessionData) Then 
			intArraySize = CInt(UBound(sarySessionData, 2))
			intLastArrayPostionPointer = intArraySize
		Else
			intArraySize = 0
		End If
	
		'Response.Write(intArraySize)
	
		'Iterate through array
		For intSessionArrayPass = 0 To intArraySize
	
			'Check that the array position is not over 20 minutes old and remove them, AND part to make sure we are not re-checking out dated sessions that are already moved
			If CDate(sarySessionData(2, intSessionArrayPass)) < DateAdd("n", -intSessionTimeout, Now()) AND intSessionArrayPass < intLastArrayPostionPointer Then
			
				'First remove any unread post arrays that maybe stored in the memory for this session
				Application("sarryUnReadPosts" &  sarySessionData(0, intSessionArrayPass)) = ""
				
				'Check that the array postion pointer is not for an outdated session (AND part for error handling as don't want intLastArrayPostionPointer to be less than 0)
				If CDate(sarySessionData(2, intLastArrayPostionPointer)) < DateAdd("n", -intSessionTimeout, Now()) AND intLastArrayPostionPointer > 0 Then intLastArrayPostionPointer = intLastArrayPostionPointer - 1
	
				'Swap this array postion with the last in the array
				sarySessionData(0, intSessionArrayPass) = sarySessionData(0, intLastArrayPostionPointer)
				sarySessionData(1, intSessionArrayPass) = sarySessionData(1, intLastArrayPostionPointer)
				sarySessionData(2, intSessionArrayPass) = sarySessionData(2, intLastArrayPostionPointer)
				sarySessionData(3, intSessionArrayPass) = sarySessionData(3, intLastArrayPostionPointer)
				
				'Decrement the last array pointer
				If intLastArrayPostionPointer > 0 Then intLastArrayPostionPointer = intLastArrayPostionPointer - 1
	
	
	
			'ElseIf user has a session, read in data, and update last access time (for security we also check the IP address)
			ElseIf sarySessionData(0, intSessionArrayPass) = strSessionID AND sarySessionData(1, intSessionArrayPass) = strIP Then
				
				'Set blnFoundSession to true
				blnFoundSession = true
	
				'Read in session data
				strSessionData = sarySessionData(3, intSessionArrayPass)
				
				
				'Update last access time
				sarySessionData(2, intSessionArrayPass) = internationalDateTime(Now())
			End If
	
			'If creating a new session for the user and the session ID already exists create a new one
			If strNewSessionID = sarySessionData(0, intSessionArrayPass) Then
				strNewSessionID = createForumSessionID()
				intSessionArrayPass = 0
			End If
		Next
	
	
		'Remove the last array position as it is no-longer needed
		If intArraySize > intLastArrayPostionPointer Then ReDim Preserve sarySessionData(3, intLastArrayPostionPointer)
		
	
	
	
		'If the user doesn't have a session create one
		If blnFoundSession = false Then
			
			'Create an ID key for the session
			strNewSessionIDKey = LCase(hexValue(12))
			
			'Get the array size
			intArraySize = CInt(UBound(sarySessionData, 2))
			
			'Increarment the array size
			intArraySize = intArraySize + 1
	
			'ReDimesion the array
			ReDim Preserve sarySessionData(3, intArraySize)
	
			'Update the new array position with the new session details
			sarySessionData(0, intArraySize) = strNewSessionID
			sarySessionData(1, intArraySize) = strIP
			sarySessionData(2, intArraySize) = internationalDateTime(Now())
			sarySessionData(3, intArraySize) = "KEY=" & strNewSessionIDKey & ";"
	
			'Initilise the session id variable
			strSessionID = strNewSessionID
			strSessionData = "KEY=" & strNewSessionIDKey & ";"
	
			'Create a cookie and querystring with the session ID
			Call setCookie("sID", "SID", strNewSessionID, False)
		End If
		
		
		'Update the application level variable
		Application(strAppPrefix & "sarySessionData") = sarySessionData
	
		'Unlock the application
		Application.UnLock
		

	End If






	'If cookies are not detected setup to use querystrings to pass around the session ID
	'For better Search Engine indexing don't use Session Querystring if detected as Search Robot
	If blnCookiesDetected = false Then
		strQsSID = strSessionID			'For form entries etc.
		strQsSID1 = "?SID=" & strSessionID	'For ? querystrings
		strQsSID2 = "&amp;SID=" & strSessionID	'For &amp; querystrings
		strQsSID3 = "&SID=" & strSessionID	'For & querystrings - redirects
	End If

End Sub







'******************************************
'***  Get application session data	***
'******************************************

'Function to read in application session data for those without cookies
Private Function getSessionItem(ByVal strSessionKey)

	Dim saryUserSessionData
	Dim intSessionArrayPass
	
	'Don't run code if a search engine spider unless search engine sessions are enabled
	If (blnSearchEngineSessions = True AND strOSType = "Search Robot") OR NOT strOSType = "Search Robot" Then
	
		'Append '=' to the end of the session key to make full session key (eg. key=)
		strSessionKey = strSessionKey & "="
	
		'Split the session data up into an array
		saryUserSessionData = Split(strSessionData, ";")
	
		'Loop through array to get the required data
		For intSessionArrayPass = 0 to UBound(saryUserSessionData)
			If InStr(saryUserSessionData(intSessionArrayPass), strSessionKey) Then
				'Return the data item
				getSessionItem = Replace(saryUserSessionData(intSessionArrayPass), strSessionKey, "", 1, -1, 1)
			End If
		Next
	End If

End Function







'******************************************
'***  Save application session data	***
'******************************************

'Sub procedure to save application session data for those without cookies
Private Sub saveSessionItem(ByRef strSessionKey, ByRef strSessionKeyValue)

	'Response.write("session_updated")

	Dim saryUserSessionData
	Dim intSessionArrayPass
	Dim strNewSessionData
	Dim intItemArrayPass

	'Don't run code if a search engine spider unless search engine sessions are enabled
	If (blnSearchEngineSessions = True AND strOSType = "Search Robot") OR NOT strOSType = "Search Robot" Then
		
		'Read in the application session for the user and update the session data in it
		For intSessionArrayPass = 0 To UBound(sarySessionData, 2)
	
			'If we find the users session data update it
			If sarySessionData(0, intSessionArrayPass) = strSessionID Then
	
				'Split the session data up into an array
				saryUserSessionData = Split(sarySessionData(3, intSessionArrayPass), ";")
	
				'Loop through array and do NOT add the updated key to session data
				For intItemArrayPass = 0 to UBound(saryUserSessionData)
	
					If InStr(saryUserSessionData(intItemArrayPass), strSessionKey) = 0 AND saryUserSessionData(intItemArrayPass) <> "" Then
	
						'Create session data string
						strNewSessionData = strNewSessionData & saryUserSessionData(intItemArrayPass) & ";"
					End If
				Next
	
				'Add the updated or new key to session string
				If strSessionKeyValue <> "" Then strNewSessionData = strNewSessionData & strSessionKey & "=" & strSessionKeyValue
	
				'Update the array
				sarySessionData(3, intSessionArrayPass) = ";" & strNewSessionData
				
				
				'If using a database save the session data to database
				If blnDatabaseHeldSessions AND ((blnSearchEngineSessions = True AND strOSType = "Search Robot") OR NOT strOSType = "Search Robot") Then
					
					'Make sure session data is SQL safe to prevent SQL injections
					strNewSessionData = ";" & formatSQLInput(strNewSessionData)

					'Initilse sql statement
				 	strSQL = "UPDATE " & strDbTable & "Session" & strRowLock & " " & _
					"SET " & strDbTable & "Session.Last_active = " & formatDbDate(Now()) & ", " & strDbTable & "Session.Session_data = '" & strNewSessionData & "' " & _
					"WHERE " & strDbTable & "Session.Session_ID = '" & strSessionID & "';"
					
					'Set error trapping
					On Error Resume Next
		
					'Write to database
					adoCon.Execute(strSQL)
					
					'If an error has occurred write an error to the page
					If Err.Number <> 0 Then	Call errorMsg("An error has occurred while writing to the database.", "saveSessionItem()_update_session_data", "functions_session_data.asp")
				
					'Disable error trapping
					On Error goto 0
					
				
				'Else save the sesison data to the application array
				Else
	
					'Lock the application so that no other user can try and update the application level variable at the same time
					Application.Lock
		
					'Update the application level session data for the user
					Application(strAppPrefix & "sarySessionData") = sarySessionData
		
					'Unlock the application
					Application.UnLock
				End If
	
				'Exit for loop
				Exit For
			End If
		Next
	End If
End Sub




'******************************************
'***  		Create Session ID     *****
'******************************************

Private Function createForumSessionID()

  	Dim dblNowDate 
  	Dim strClientIP
  	
  	'Get the time as a double number
  	dblNowDate = CDbl(Now())
  	dblNowDate = Mid(dblNowDate, InStr(1, dblNowDate, ".")+1, Len(dblNowDate))
  	
	
  	'Application session number to ensure that the sessions ID is unquie
	Application("SessionNum") = Application("SessionNum") + 1
  	
	'Calculate a code for the user
	createForumSessionID = Application("SessionNum") & "-" & LCase(hexValue(20)) & dblNowDate
	
End Function
%>