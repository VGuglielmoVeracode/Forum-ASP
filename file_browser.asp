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





Response.Expires = -1
Response.ExpiresAbsolute = Now() - 2
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"


'Dim variables
Dim objFSO			'Holds the file system object
Dim objFSOfolder		'Holds the FSO file object
Dim objFSOsubFolder
Dim objFSOfile
Dim saryAllowedFileTypes	'Holds the allowd file types
Dim intExtensionLoopCounter	'Loop counter to check file extensions
Dim strFileName			'Holds the file name
Dim strFileType			'Holds the file type
Dim intFileSize			'Holds the file size
Dim strFileIcon			'Holds the icon for the file
Dim strFileExtension		'Holds the file extension
Dim intElementIDno		'Holds the element ID number
Dim strSubFolderName		'Holds the name of the subfolder
Dim strFolderPath		'Holds the path to the folder
Dim strSubFolderUp		'Hollds the path to the folder above
Dim strMode			'Holds the page mode
Dim lngUserProfileID		'Holds the profile ID of the user
Dim blnUploadFolderExsist	'Set to true if the user has an upload folder
Dim intParentStripLoop		'Loop variable


'Initialise variables
intElementIDno = 0
strSubFolderName = Request.QueryString("sub")
strUploadFileTypes = strUploadFileTypes & ";" & strImageTypes
blnUploadFolderExsist = False



'If the user is not allowed kick 'em
If bannedIP() OR  blnActiveMember = False OR blnBanned OR intGroupID = 2 OR (blnAttachments = false AND blnImageUpload = false AND blnAvatarUploadEnabled = False) Then
	
	'Clean up
	Call closeDatabase()

	'Redirect
	Response.Redirect("insufficient_permission.asp" & strQsSID1)
End If




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






'Get the user ID of the memebr being edited by the admin and set to the users upload directory
If (blnAdmin OR (blnModerator AND LngC(Request.QueryString("PF")) > 2)) AND strMode = "A" Then
	
	lngUserProfileID = LngC(Request.QueryString("PF"))
	strUploadFilePath = strUploadOriginalFilePath & "/" & lngUserProfileID

'Get the logged in ID number
Else
	lngUserProfileID = lngLoggedInUserID
End If



'Check the user has an upload folder, if they have uploaded annything
blnUploadFolderExsist = userUploadFolder(strUploadFilePath)


'Reset Server Objects
Call closeDatabase()


'Get what we are looking for

'Get the image types allowed
saryAllowedFileTypes = Split(Trim(strUploadFileTypes), ";")
	
'Get the file path
strFolderPath = strUploadFilePath


'See if this is a subfolder being looked in
If strSubFolderName <> "" Then
	
	'Replace any \ with \
	strSubFolderName = Replace(strSubFolderName, "/", "\", 1, -1, 1)

	'Loop through and remove any parent paths that could course a security issue
	For intParentStripLoop = 0 To 1
		
		'Look for ..\
		If Instr(1, strSubFolderName, "..\", 1) Then
			strSubFolderName = Replace(strSubFolderName, "..\", "", 1, -1, 1)
			intParentStripLoop = 0 'Loop again incase there are anymore
		End If
		
		'Remove any .\
		If Instr(1, strSubFolderName, ".\", 1) Then
			strSubFolderName = Replace(strSubFolderName, ".\", "", 1, -1, 1)
			intParentStripLoop = 0 'Loop again incase there are anymore
		End If
	Next
	
	'Get the complete folder path to the subfolder in the upload directory
	strFolderPath = strFolderPath &  strSubFolderName
	

	'Calculate one folder up path
	strSubFolderUp = Mid(strSubFolderName, 1, (Len(strSubFolderName) - Len(Mid(strSubFolderName, InstrRev(strSubFolderName, "\")))))
	
End If



%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strTxtFileExplorer %></title>

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

<script language="JavaScript">
	
//Function to get subfolder
function subFolder(sub){
	self.document.location.href = 'RTE_file_browser.asp?look=<% = Server.URLEncode(Request.QueryString("look")) %>&sub=<% = Replace(strSubFolderName, "\", "\\", 1, -1, 1) %>\\' + sub + '<% = strQsSID2 %>';
}

//Function to get subfolder
function upFolder(){
	self.document.location.href = 'RTE_file_browser.asp?look=<% = Server.URLEncode(Request.QueryString("look")) %>&sub=<% = Replace(strSubFolderUp, "\", "\\", 1, -1, 1) %><% = strQsSID2 %>';
}

//Function to preview image
function upadatePreview(fileName, fileSize, fileType){
	//Get the file extension to check
	var extension = fileName;
	extension = extension.slice(extension.lastIndexOf('.'), extension.length).toLowerCase();

	//Display file if of the correct type
	if (extension==".txt" 
	 || extension==".text" 
	 || extension==".htm" 
	 || extension==".html" 
	 || extension==".gif" 
	 || extension==".jpg" 
	 || extension==".jpeg" 
	 || extension==".png" 
	 || extension==".bmp"){
		try {
			self.parent.document.getElementById('prevWindow').contentWindow.location.href = '<% = Replace(strFolderPath, "\", "/", 1, -1, 1) %>/' + fileName;	
		}catch(exception){
		}
	
	}else{
		self.parent.document.getElementById("prevWindow").contentWindow.location.href="RTE_popup_link_preview.asp?b=0<% = strQsSID2 %>";
	
	}
	
	self.parent.document.getElementById('fileNameDisplay').innerHTML = fileName;
	self.parent.document.getElementById('fileName').value = fileName;
	self.parent.document.getElementById('fileType').innerHTML = fileType;
	self.parent.document.getElementById('fileSize').innerHTML = fileSize + 'KB';
	self.parent.document.getElementById('fileDownload').innerHTML = '<a href="<% = Replace(strFolderPath, "\", "/", 1, -1, 1) %>/' + fileName + '" target="_blank">' + fileName + '</a>';
}

//Function to hover file item
function overIcon(iconItem) {
	iconItem.style.backgroundColor='#CCCCCC';
}

//Function to moving off file item
function outIcon(iconItem) {
	iconItem.style.backgroundColor='#FFFFFF';
}
</script>

<style type="text/css">
<!--
.fileText {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 11px;
	color: #000000;
}
html,body { 
	border: 0px; 
}
-->
</style>
</head>
<body bgcolor="#FFFFFF" leftmargin="2" topmargin="2" marginwidth="2" marginheight="2">
 <table width="100%"  border="0" cellspacing="0" cellpadding="3"><%

'If the user has an upload folder display the contents
If blnUploadFolderExsist Then
	
	'Create the file system object
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	'Create a folder object with the contents of the folder
	Set objFSOfolder = objFSO.GetFolder(Server.MapPath(strFolderPath))

	
	'Else show an icon for Parent Directory
	If strSubFolderName <> "" Then
		intElementIDno = intElementIDno + 1
			
		Response.Write(vbCrLf & "  <tr onMouseover=""overIcon(this)"" onMouseout=""outIcon(this)"" onclick=""upFolder()"" style=""cursor: default;"">" & _
			       vbCrLf & "    <td colspan=""3"" class=""fileText""><img src=""" & strImagePath & "icon_up_folder.gif"" align=""absbottom"" title=""" & strTxtParentDirectory & """ border=""0"">" & strTxtParentDirectory & "</td>" & _
			       vbCrLf & "  </tr>")
	
	End If
	
	'Show any sub folders
	For Each objFSOsubFolder In objFSOfolder.SubFolders
		
		If NOT objFSOsubFolder.Name = "temp" AND NOT objFSOsubFolder.Name = "_vti_cnf" Then
			intElementIDno = intElementIDno + 1
				
			Response.Write(vbCrLf & "  <tr onMouseover=""overIcon(this)"" onMouseout=""outIcon(this)"" onclick=""subFolder(document.getElementById('ico" & intElementIDno & "').title)"" style=""cursor: default;"">" & _
				       vbCrLf & "    <td colspan=""3"" class=""fileText""><img src=""" & strImagePath & "icon_folder.gif"" align=""absbottom"" id=""ico" & intElementIDno & """ title=""" & objFSOsubFolder.Name & """ border=""0"">"  & objFSOsubFolder.Name & "</td>" & _
			 	       vbCrLf & "  </tr>")
		End If
	Next
	
	
	
	'Loop through all the files in the folder
	For Each objFSOfile in objFSOfolder.Files
			
		'Loop through to check if the file has an allowed extension
		For intExtensionLoopCounter = 0 To UBound(saryAllowedFileTypes)
			
			'If the extension is allowed show the file
			If LCase(objFSO.GetExtensionName(objFSOfile.Name)) = saryAllowedFileTypes(intExtensionLoopCounter) AND NOT objFSOfile.Name = "folder_info.xml" Then
				
				'Initilse the icon file with unknown file type
				strFileIcon = "icon_unknown.gif"
				intElementIDno = intElementIDno + 1
			
				'Read in details
				strFileName = objFSOfile.Name
				strFileType = objFSOfile.Type
				intFileSize = LngC(objFSOfile.Size / 1024)
				strFileExtension = LCase(objFSO.GetExtensionName(objFSOfile.Name))
				
				'Check the length of the file name is not to long
				If Len(strFileName) > 21 Then
					strFileName = Trim(Mid(strFileName, 1, 23)) & "..." & strFileExtension
				End If
				
				'Check the length of the file type is not to long
				If Len(strFileType) > 11 Then
					strFileType = Trim(Mid(strFileType, 1, 11)) & "..."
				End If
				
				'Get the icon for the file type
				Select Case strFileExtension
					Case "jpg"
						strFileIcon = "icon_jpg.gif"
					Case "jpeg"
						strFileIcon = "icon_jpg.gif"
					Case "gif"
						strFileIcon = "icon_gif.gif"
					Case "bmp"
						strFileIcon = "icon_bmp.gif"
					Case "png"
						strFileIcon = "icon_png.gif"
					Case "doc"
						strFileIcon = "icon_doc.gif"
					Case "htm"
						strFileIcon = "icon_htm.gif"
					Case "html"
						strFileIcon = "icon_htm.gif"
					Case "rtf"
						strFileIcon = "icon_doc.gif"
					Case "txt"
						strFileIcon = "icon_txt.gif"
					Case "text"
						strFileIcon = "icon_txt.gif"
					Case "zip"
						strFileIcon = "icon_zip.gif"
					Case "rar"
						strFileIcon = "icon_zip.gif"
					Case "tar"
						strFileIcon = "icon_zip.gif"
					Case "exe"
						strFileIcon = "icon_exe.gif"
					Case "pdf"
						strFileIcon = "icon_pdf.gif"
				
				End Select
				
				
			
				Response.Write(vbCrLf & "  <tr onMouseover=""overIcon(this)"" onMouseout=""outIcon(this)"" OnClick=""upadatePreview('" & objFSOfile.Name & "', " & intFileSize & ", '" & objFSOfile.Type & "')"" style=""cursor: default;"">" & _ 
				 	       vbCrLf & "    <td class=""fileText"" width=""58%""><img src=""" & strImagePath & strFileIcon & """ align=""absbottom"" id=""ico" & intElementIDno & """ border=""0"" alt=""" & objFSOfile.Name & """ title=""" & objFSOfile.Name & """>" & strFileName & "</td>" & _
				 	       vbCrLf & "    <td class=""fileText"" width=""15%"">" & intFileSize & "KB</td>" & _
				 	       vbCrLf & "    <td class=""fileText"" width=""27%"">" & strFileType & "</td>" & _
				 	       vbCrLf & "  </tr>")
				 	       
				'As we have displayed this file jump out the for loop as we don't need to check any other extensions
				Exit For
			End If
		Next
			
	
	Next
	
	'Distroy objects
	Set objFSOsubFolder = Nothing
	Set objFSOfile = Nothing	
	Set objFSOfolder = Nothing
	Set objFSO = Nothing  

'Else no user upload folder
Else
	Response.Write(vbCrLf & "  <tr>" & _
			vbCrLf & "    <td colspan=""3"" class=""fileText"">" & strTxtUploadFolderEmpty & "</td>" & _
			vbCrLf & "  </tr>")

End If



%>
  </table>
 </body>
</html>