<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="functions/functions_upload.asp" -->
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



'Dimension veriables
Dim strMode 		'Holds the page mode (eg admin)
Dim lngUserProfileID	'Holds the profile ID of the user
Dim strFileName		'Holds the file name to be deleted
Dim saryFileUploadTypes	'Array of upload file types
Dim objFSO		'File system object
Dim blnFileDeleted	'Set to true if file deleted


'Initialise
blnExtensionOK = False
blnFileDeleted = false


'If the user is not allowed kick 'em
If bannedIP() OR  blnActiveMember = False OR blnBanned OR intGroupID = 2 OR (blnAttachments = false AND blnImageUpload = false) Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If



'Check the session ID to stop CSRF
Call checkFormID(Request.QueryString("XID"))



'Read in the mode of the page
strMode = Trim(Mid(Request.QueryString("M"), 1, 1))

'If this is not an admin but in admin mode then see if the user is a moderator
If blnAdmin = False AND strMode = "A" AND blnModeratorProfileEdit Then
	
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


'Clean up
Call closeDatabase()




'Get the user ID of the memebr being edited by the admin and set to the users upload directory
If (blnAdmin OR (blnModerator AND LngC(Request.QueryString("PF")) > 2)) AND strMode = "A" Then
	
	lngUserProfileID = LngC(Request.QueryString("PF"))
	strUploadFilePath = strUploadOriginalFilePath & "/" & lngUserProfileID

'Get the logged in ID number
Else
	lngUserProfileID = lngLoggedInUserID
End If


'Read in the file name to be deleted
strFileName = decodeString(Request.QueryString("fileName"))


'Now to protect the server really need to do allot of checking on the file name passed across
strFileName = formatFileName(strFileName)

'Create an array of upload file types
saryFileUploadTypes = Split(Trim(strImageTypes & ";" & strUploadFileTypes), ";")

'Make sure the file extension is OK
blnExtensionOK = fileExtension(strFileName, saryFileUploadTypes)



'File deletion (if extension is OK)
If blnExtensionOK Then
	
	'Create an instance of the FSO object
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
			
	
	'See if file exsists
	If objFSO.FileExists(Server.MapPath(strUploadFilePath & "\" & strFileName)) Then
					
		'Delete file
		objFSO.DeleteFile Server.MapPath(strUploadFilePath & "\" & strFileName), true
		
		'Set to true
		blnFileDeleted = True
	End If
			
	'Release the FSO object
	Set objFSO = Nothing
End If

'Go back to file manager
If strMode = "A" Then 
	Response.Redirect("file_manager.asp?DL=" & blnFileDeleted & "&PF=" & lngUserProfileID & "&M=A" & strQsSID3)
Else
	Response.Redirect("file_manager.asp?DL=" & blnFileDeleted & strQsSID3)
End If

%>