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





'Upload global variables

Dim strImageName		'Holds the file name
Dim blnExtensionOK		'Set to false if the extension of the file is not allowed
Dim lngErrorFileSize		'Holds the file size if the file is not saved because it is to large
Dim dblErrorAllotedFileSpace	'Holds the alloted space size error
Dim blnFileExists		'Set to true if the file already exists
Dim blnFileSpaceExceeded	'Set to true if the alloted file space is exceeded
Dim blnSecurityScanFail		'Set to true if security scan fails


'Intiliase global variables
blnExtensionOK = True
blnFileExists = False
blnFileSpaceExceeded = False
blnSecurityScanFail = False
lngErrorFileSize = 0
dblErrorAllotedFileSpace = 0







'******************************************
'***	   File Upload Function        ****
'******************************************

'Function to upload a file
Private Function fileUpload(ByVal strUploadType)

	'Dimension variables
	Dim objUpload		
	Dim strNewFileName	
	Dim strOriginalFileName	
	Dim objFSO
	Dim objTextStream
	Dim strTempFile
	Dim strExtension
	Dim saryFileUploadTypes
	Dim lngMaxFileSize
	Dim lngLoopCounter
	Dim objAspJpeg
	
	
	
	'Make sure the user has a folder to upload to
	createUserFolder(strUploadFilePath)
	
	
	
	'First check the user has not gone over their alloted space
	'Get used space
	dblErrorAllotedFileSpace = folderSize(strUploadFilePath)
	
	'Check to see if the user has gone over the alloted space
	If CDbl(dblErrorAllotedFileSpace) > CDbl(intUploadAllocatedSpace) Then
		blnFileSpaceExceeded = True
		Exit Function
	End If
	
	
	
	
	'Get the file types we are uploading
	If strUploadType = "file" Then
		lngMaxFileSize = lngUploadMaxFileSize
		saryFileUploadTypes = Split(Trim(strUploadFileTypes), ";")
	ElseIf strUploadType = "image" Then
		lngMaxFileSize = lngUploadMaxImageSize
		saryFileUploadTypes = Split(Trim(strImageTypes), ";")
	End If
	
	'If no file type of extensions set then leave now
	If isArray(saryFileUploadTypes) = False Then 
		blnExtensionOK = False
		Exit Function
	End If





	'******************************************
	'***	     Upload components         ****
	'******************************************

	'Select which upload component to use
	Select Case strUploadComponent


		'******************************************
		'***     Persits AspUpload component   ****
		'******************************************

		'Persits AspUpload upload component - tested with version 3.0
		Case "AspUpload", "AspUpload2"
		
			'Set error trapping
			On Error Resume Next

			'Create upload object
			Set objUpload = Server.CreateObject("Persits.Upload.1")
			
			'If AspUpload 3.x or above get the progress ID for the progress bar
			If strUploadComponent = "AspUpload" Then objUpload.ProgressID = Request.QueryString("PID")
				
			'Check to see if an error has occurred
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then	Call errorMsg("An error has occurred while uploading file/image.<br />Please check the Persits AspUpload Component is installed on the server.", "create_AspUpload_object", "functions_upload.asp")
		
			'Disable error trapping
			On Error goto 0
			
		
			With objUpload
			
				'If ASPUpload 3.x and UTF-8 set the code page to 65001 for UTF-8 support in AspUpload
				If strUploadComponent = "AspUpload" AND LCase(strPageEncoding) = "utf-8" Then
					.CodePage = 65001
				End If

				'make sure files arn't over written
				.OverwriteFiles = False

				'We need to save the file to memory before we can find out anything about it
				.Save

				'Get the file name
				strNewFileName = .Files(1).ExtractFileName

				'Filter file name to remove anything that isn't allowed by the filters
				strNewFileName = formatFileName(strNewFileName)

				'Check the file size is not above the max allowed size, this is done using a function not the compoent to stop an exception error
				lngErrorFileSize = fileSize(.Files(1).Size, lngMaxFileSize)

				'Loop through all the allowed extensions and see if the file has one
				blnExtensionOK = fileExtension(strNewFileName, saryFileUploadTypes)
				
				'Check if file exsists
				blnFileExists = .FileExists(Server.MapPath(strUploadFilePath) & "\" & strNewFileName)
				
				'If file exists and duplicate file names allowed rename the file
				If blnUploadDuplicateFiles and blnFileExists Then 
					strNewFileName = renameDuplicateFileName(strNewFileName)
					blnFileExists = False
				End If
				
					
				'If the file is OK save it to disk
				If lngErrorFileSize = 0 AND blnExtensionOK AND blnFileExists = False Then

					'Save the file to disk with new file name
					'** SaveAsVirtual is used as SaveAs is often disabled by the web host **
					.Files(1).SaveAsVirtual strUploadFilePath & "\" & strNewFileName
				
					'Pass the filename back
					fileUpload = strNewFileName


				'Else if it is not OK delete the uploaded file
				Else
					.Files(1).Delete

				End If

			End With

			'Clean up
			Set objUpload = Nothing




		'******************************************
		'***         Dundas Upload component   ****
		'******************************************

		'Dundas upload component free from http://www.dundas.com - tested with version 2.0
		Case "Dundas"
		
			'Set error trapping
			On Error Resume Next

			'Create upload object
			Set objUpload = Server.CreateObject("Dundas.Upload")
			
			'Check to see if an error has occurred
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then	Call errorMsg("An error has occurred while uploading file/image.<br />Please check the Dundas Upload Component is installed on the server.", "create_Dundas_Upload_object", "functions_upload.asp")
		
			'Disable error trapping
			On Error goto 0
			
		
			With objUpload

				'Make sure we are using a virtual directory for script
				.UseVirtualDir = True

				'Make sure the file names are not unique at this time
				.UseUniqueNames = False

				'Save the file first to memory
				.SaveToMemory()

				'Get the file name, the path mehod will be empty as we are saving to memory so use the original file path of the users system to get the name
				strNewFileName = .GetFileName(.Files(0).OriginalPath)

				
				'Filter file name to remove anything that isn't allowed by the filters
				strNewFileName = formatFileName(strNewFileName)

				'Check the file size is not above the max allowed size, this is done using a function not the compoent to stop an exception error
				lngErrorFileSize = fileSize(.Files(0).Size, lngMaxFileSize)

				'Loop through all the allowed extensions and see if the file has one
				blnExtensionOK = fileExtension(strNewFileName, saryFileUploadTypes)
				
				'Check if file exists
				blnFileExists = .FileExists(strUploadFilePath & "\" & strNewFileName)
				
				'If file exists and duplicate file names allowed rename the file
				If blnUploadDuplicateFiles and blnFileExists Then 
					strNewFileName = renameDuplicateFileName(strNewFileName)
					blnFileExists = False
				End If

				'If the file is OK save it to disk
				If lngErrorFileSize = 0 AND blnExtensionOK AND blnFileExists = False Then

					
					'Save the file to disk
					.Files(0).SaveAs strUploadFilePath & "/" & strNewFileName

					'Pass the filename back
					fileUpload = strNewFileName
				End If
			End With

			'Clean up
			Set objUpload = Nothing




		'******************************************
		'***  SoftArtisans FileUp component    ****
		'******************************************

		'SA FileUp upload component - tested with version 4
		Case "fileUp"
		
			'Set error trapping
			On Error Resume Next

			'Create upload object
			Set objUpload = Server.CreateObject("SoftArtisans.FileUp")
			
			'Check to see if an error has occurred
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then	Call errorMsg("An error has occurred while uploading file/image.<br />Please check the SoftArtisans FileUp Component is installed on the server.", "create_SoftArtisans_FileUp_object", "functions_upload.asp")
			
			'Disable error trapping
			On Error goto 0
			
			
			With objUpload

				'Over write files or an exception will occur if it already exists
				.OverWriteFiles = True

				'Set the upload path
				.Path = Server.MapPath(strUploadFilePath)

				'Get the file name, the path mehod will be empty as we are saving to memory so use the original file path of the users system to get the name
				strNewFileName = Mid(.UserFilename, InstrRev(.UserFilename, "\") + 1)

				
				'Filter file name to remove anything that isn't allowed by the filters
				strNewFileName = formatFileName(strNewFileName)

				'Check the file size is not above the max allowed size, this is done using a function not the compoent to stop an exception error
				lngErrorFileSize = fileSize(.TotalBytes, lngMaxFileSize)

				'Loop through all the allowed extensions and see if the file has one
				blnExtensionOK = fileExtension(strNewFileName, saryFileUploadTypes)
				
				'Create the file system object
				Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
				
				'Check if file exsists
				blnFileExists = objFSO.FileExists(Server.MapPath(strUploadFilePath) & "\" & strNewFileName)
				
				'If file exists and duplicate file names allowed rename the file
				If blnUploadDuplicateFiles and blnFileExists Then 
					strNewFileName = renameDuplicateFileName(strNewFileName)
					blnFileExists = False
				End If
				
				'Drop FSO as no longer needed
				Set objFSO = Nothing
				

				'If the file is OK save it to disk
				If lngErrorFileSize = 0 AND blnExtensionOK AND blnFileExists = False Then

					'Save the file to disk
					.SaveAs strNewFileName

					'Pass the filename back
					fileUpload = strNewFileName
				End If

			End With

			'Clean up
			Set objUpload = Nothing




		'******************************************
		'***  	AspSmartUpload component       ****
		'******************************************

		'AspSmartUpload upload component free from http://www.aspsmart.com
		Case "aspSmart"
		
			'Set error trapping
			On Error Resume Next

			'Create upload object
			Set objUpload = Server.CreateObject("aspSmartUpload.SmartUpload")
			
			'Check to see if an error has occurred
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then	Call errorMsg("An error has occurred while uploading file/image.<br />Please check the Asp Smart Upload Component is installed on the server.", "create_AspSmartUpload_object", "functions_upload.asp")
			
			'Disable error trapping
			On Error goto 0
			

			With objUpload

				'Make sure we are using a virtual directory
				.DenyPhysicalPath = True

				'Save the file first to memory
				.Upload

				'Get the file name, the path mehod will be empty as we are saving to memory so use the original file path of the users system to get the name
				strNewFileName = .Files(1).Filename

				'Filter file name to remove anything that isn't allowed by the filters
				strNewFileName = formatFileName(strNewFileName)

				'Check the file size is not above the max allowed size
				lngErrorFileSize = fileSize(.Files(1).Size, lngMaxFileSize)

				'Loop through all the allowed extensions and see if the file has one
				blnExtensionOK = fileExtension(strNewFileName, saryFileUploadTypes)
				
				'Create the file system object
				Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
				
				'Check if file exsists
				blnFileExists = objFSO.FileExists(Server.MapPath(strUploadFilePath) & "\" & strNewFileName)
				
				'If file exists and duplicate file names allowed rename the file
				If blnUploadDuplicateFiles and blnFileExists Then 
					strNewFileName = renameDuplicateFileName(strNewFileName)
					blnFileExists = False
				End If
				
				'Drop FSO as no longer needed
				Set objFSO = Nothing

				'If the file is OK save it to disk
				If lngErrorFileSize = 0 AND blnExtensionOK AND blnFileExists = False Then
					
					'Save the file to disk
					.Files(1).SaveAs strUploadFilePath & "/" & strNewFileName

					'Pass the filename back
					fileUpload = strNewFileName
				End If

			End With

			'Clean up
			Set objUpload = Nothing



		'******************************************
		'***     AspSimpleUpload component     ****
		'******************************************

		'ASPSimpleUpload component
		Case "AspSimple"

			'Dimension variables
			Dim file	'Holds the FSO file object

			'Set error trapping
			On Error Resume Next

			'Create upload object
			Set objUpload = Server.CreateObject("ASPSimpleUpload.Upload")
			
			'Check to see if an error has occurred
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then	Call errorMsg("An error has occurred while uploading file/image.<br />Please check the ASPSimpleUpload is installed on the server.", "create_AspSimpleUpload_object", "functions_upload.asp")
		
			'Disable error trapping
			On Error goto 0
			
			With objUpload

				'Get the file name
				strOriginalFileName = .ExtractFileName(.Form("file"))

				'Save the amended file name
				strNewFileName = "TMP" & hexValue(7) & "_" & strOriginalFileName
				
				'Filter file name to remove anything that isn't allowed by the filters
				strNewFileName = formatFileName(strNewFileName)

				'Save the file to disk first so we can check it
				Call .SaveToWeb ("file", strUploadFilePath & "\" & strNewFileName)

				'Create the file system object
				Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

				'Create a file object with the file details
				Set file = objFSO.GetFile(Server.MapPath(strUploadFilePath) & "\" & strNewFileName)

				'Check the file size is not above the max allowed size, this is done using a function not the compoent to stop an exception error
				lngErrorFileSize = fileSize(file.Size, lngMaxFileSize)


				'Place the original file name back in the new filename variable
				strNewFileName = strOriginalFileName
				
				'Filter file name to remove anything that isn't allowed by the filters
				strNewFileName = formatFileName(strNewFileName)


				'Loop through all the allowed extensions and see if the file has one
				blnExtensionOK = fileExtension(strNewFileName, saryFileUploadTypes)

				'Check if file exsists
				blnFileExists = objFSO.FileExists(Server.MapPath(strUploadFilePath) & "\" & strNewFileName)
				
				'If file exists and duplicate file names allowed rename the file
				If blnUploadDuplicateFiles and blnFileExists Then 
					strNewFileName = renameDuplicateFileName(strNewFileName)
					blnFileExists = False
				End If

				'If the file is OK save it to disk
				If lngErrorFileSize = 0 AND blnExtensionOK AND blnFileExists = False Then

					'Save the file to disk
					Call .SaveToWeb("file", strUploadFilePath & "/" & strNewFileName)

					'Pass the filename back
					fileUpload = strNewFileName
				End If
				
				'Delete the original file
				file.Delete

			End With

			'Clean up
			Set file = Nothing
			Set objFSO = Nothing
			Set objUpload = Nothing




		'******************************************
		'***     Without component     ****
		'******************************************

		'Without component
		Case "NoComponent"

			'Set error trapping
			On Error Resume Next
			%>
			<!--#include FILE="upload_5xsoft.inc"-->
			<%

			'Create upload object
			Set objUpload = new upload_5xsoft
			
			'Check to see if an error has occurred
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then	Call errorMsg("An error has occurred while uploading file/image.<br />Please check uploading without component.", "create_NoComponent_object", "functions_upload.asp")
		
			'Disable error trapping
			On Error goto 0

			'Create a file object with the file details
			Set file = objUpload.file("file")

			'Check the file size is not above the max allowed size, this is done using a function not the compoent to stop an exception error
			lngErrorFileSize = fileSize(file.FileSize, lngMaxFileSize)

				
			'Filter file name to remove anything that isn't allowed by the filters
			strNewFileName = formatFileName(file.FileName)


			'Loop through all the allowed extensions and see if the file has one
			blnExtensionOK = fileExtension(strNewFileName, saryFileUploadTypes)

			'Create the file system object
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

			'Check if file exsists
			'blnFileExists = objFSO.FileExists(Server.MapPath(strUploadFilePath) & "\" & strNewFileName)

			'If the file is OK save it to disk
			If lngErrorFileSize = 0 AND blnExtensionOK Then 'AND blnFileExists = False 

				'Save the file to disk
				file.saveAs Server.MapPath(strUploadFilePath) & "\" & strNewFileName

				'Pass the filename back
				fileUpload = strNewFileName

			End If				

			'Clean up
			Set file = Nothing
			Set objFSO = Nothing
			Set objUpload = Nothing

	End Select
	
	
	
	
	'********************************************
	'***  Check and shrink image dimensions  ****
	'********************************************
	
	'If an image shrink the dimentions
	If (intMaxImageWidth > 0 OR intMaxImageHeight > 0) AND strUploadType = "image" AND AspJpegImage(strNewFileName) Then
		
		'Create the file system object
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		
		'Check to make sure file exsists
		If objFSO.FileExists(Server.MapPath(strUploadFilePath) & "\" & strNewFileName) Then

			'Create instance of Persits AspJPEG
			Set objAspJpeg = Server.CreateObject("Persits.Jpeg")
	
	
			'Maintain aspect ratio of image
			objAspJpeg.PreserveAspectRatio = True
	                   
	                 'As the aspect ration is kept the image is checked twice for reszing, once for height and the other for width
	                   
			'Open image to resize width
			objAspJpeg.Open Server.MapPath(strUploadFilePath) & "\" & strNewFileName

			'If more than max width shrink
			If objAspJpeg.OriginalWidth > intMaxImageWidth Then
				
				'Set the new image width
				objAspJpeg.Width = intMaxImageWidth
				
				'Save re-sized image back to disk
				objAspJpeg.Save Server.MapPath(strUploadFilePath) & "\" & strNewFileName
				
				'Open image again to resize height
				objAspJpeg.Open Server.MapPath(strUploadFilePath) & "\" & strNewFileName
			End If  
			
			
			'If more than max height shrink
			If objAspJpeg.OriginalHeight > intMaxImageHeight Then
				
				'Set the new image height
				objAspJpeg.Height = intMaxImageHeight
				
				'Save re-sized image back to disk
				objAspJpeg.Save Server.MapPath(strUploadFilePath) & "\" & strNewFileName 
			End If  
			
			
			'Clean up
			Set objAspJpeg = Nothing
		
		End If
		
		'Clean up
		Set objFSO = Nothing
	
	End If
	
	
	
	
	'******************************************
	'***  Security check for MIME change   ****
	'******************************************
	
	'Read in the uploaded file to make sure that the user is not trying to sneak through a change of content type in an image etc.
	If blnUploadSecurityCheck Then
		'Get the file extension
		If InStr(strNewFileName, ".") Then
			strExtension = Lcase(Mid(strNewFileName, InStrRev(strNewFileName, "."), 5))
		Else
			strExtension = "."
		End If
		
		'Don't run if text based file
		If strExtension <> ".txt" AND strExtension <> ".text" AND strExtension <> ".xml" AND strExtension <> ".css" AND strExtension <> ".htm" AND strExtension <> ".html" Then
		
			'Create the file system object
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		
			'Check to make sure file exsists
			If objFSO.FileExists(Server.MapPath(strUploadFilePath) & "\" & strNewFileName) Then
				
				'Create a file object with the file details
				Set file = objFSO.GetFile(Server.MapPath(strUploadFilePath) & "\" & strNewFileName)
				
				' Open the file for reading (1) as an ascii file (0)
				Set objTextStream = file.OpenAsTextStream(1, 0)
			
				'Read in line by line and check the content type is not altered
				Do While Not objTextStream.AtEndOfStream
					strTempFile = strTempFile & objTextStream.readline
				Loop
				
				'Clean up (done now to prevent a permissions error later)
				Set objTextStream = nothing
				
				'Trim and set as lower case
				strTempFile = Trim(strTempFile)
				
				'If file is empty (do this before using 'replace' in the next block to prevent errors if the file is empty)
				If strTempFile = "" Then 
					
					blnSecurityScanFail = True
					
				'Else the file does contain something so check it doesn't contain malicious code
				Else
				
					'LCase
					strTempFile = LCase(strTempFile)
				
					'Remove spaces and tabs
					strTempFile = Replace(strTempFile, Chr(9), "", 1, -1, 1) 'Tabs
					strTempFile = Replace(strTempFile, " ", "", 1, -1, 1)
					
					%><!--#include file="unsafe_upload_content_inc.asp" --><%
					
				
					'If the file type is an image do some futher checking
					If strExtension = ".gif" OR strExtension = ".jpg" OR strExtension = ".png" OR strExtension = ".jpeg" OR strExtension = ".jpe" OR strExtension = ".tiff" OR strExtension = ".bmp" Then
						
						'Loop through the array of disallowed HTML tags
						For lngLoopCounter = LBound(saryUnSafeHTMLtags) To UBound(saryUnSafeHTMLtags)
						
								
							'If the disallowed HTML is found set the file as not being allowed
							If Instr(1, strTempFile,  saryUnSafeHTMLtags(lngLoopCounter), 1) Then
								'For testing purposes
								'Response.Write(" - " & saryUnSafeHTMLtags(lngLoopCounter))
								'Response.End
								
								'Set the security scan fail boolen to true
								blnSecurityScanFail = True
							End If
						Next
					End If
				End If
				
				'If security scan fails then delete the image
				If blnSecurityScanFail Then 
					file.Delete
					strNewFileName = ""
				End If
					
				
			End If
			
			'Clean up
			Set file = Nothing
			Set objFSO = Nothing
		End If
	End If

End Function





'******************************************
'***	Check file size function       ****
'******************************************
Function fileSize(ByVal lngFileSize, ByVal lngMaxFileSize)

	'If the file size is to large place the present file size in then return the file size
	If CLng(lngFileSize / 1024) > lngMaxFileSize Then

		fileSize = CLng(lngFileSize / 1024)

	'Else set the return value to 0
	Else
		fileSize = 0
	End If

End Function





'******************************************
'***	Check file ext. function       ****
'******************************************
Function fileExtension(ByVal strFileName, ByVal saryFileUploadTypes)

	'Dimension varibles
	Dim intExtensionLoopCounter

	'Intilaise return value
	fileExtension = False

	'Loop through all the allowed extensions and see if the file has one
	For intExtensionLoopCounter = 0 To UBound(saryFileUploadTypes)

		If LCase(Right(strFileName, Len(saryFileUploadTypes(intExtensionLoopCounter))+1)) = "." & LCase(saryFileUploadTypes(intExtensionLoopCounter)) Then fileExtension = True
	Next

End Function





'******************************************
'***	Format file names      	       ****
'******************************************
'Format file names to strip caharacters that will otherwise be stripped by the filters producing dead links
Private Function formatFileName(ByVal strInputEntry)

	'Dimension variable
	Dim intLoopCounter 	'Holds the loop counter
	

	'Loop through the ASCII characters 0 to 31
	For intLoopCounter = 0 to 31
		strInputEntry = Replace(strInputEntry, CHR(intLoopCounter), "", 1, -1, 0)
	Next
	
	'Windows illegal filename characters
	strInputEntry = Replace(strInputEntry, "/", "", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "\", "", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, ":", "", 1, -1, 0) 'This also removes : from the time set for IOS images
	strInputEntry = Replace(strInputEntry, ";", "", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "*", "", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "?", "", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, """", "", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "<", "", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, ">", "", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "|", "", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "'", "", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, ",", "", 1, -1, 1)
	
	'Replace space with underscore
	strInputEntry = Replace(strInputEntry, " ", "_", 1, -1, 1)

	'Strip others that would otherwise later be stripped by the image/file link filters and prevent the file/image displaying
	strInputEntry = Replace(strInputEntry, "[", "", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "]", "", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "(", "", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, ")", "", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "{", "", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "}", "", 1, -1, 1)

	'Return
	formatFileName = strInputEntry
End Function




'******************************************
'***	Rename Duplicate Files         ****
'******************************************

Private Function renameDuplicateFileName(ByVal strInputFimeNameEntry)

	Dim strAppendDateFormat
	
	
	'Get the Date in International format
	strAppendDateFormat = internationalDateTime(Now())
	
	
	'Change date format to nicer file format
	strAppendDateFormat = Replace(strAppendDateFormat, ":", "-", 1, -1, 0)
	strAppendDateFormat = Replace(strAppendDateFormat, " ", "_", 1, -1, 0)


	'Append the date time to end of file name
	strInputFimeNameEntry = Replace(strInputFimeNameEntry, ".", "_" & strAppendDateFormat & ".")
	

	'Return
	renameDuplicateFileName = strInputFimeNameEntry
		
End Function		






'**********************************************
'***   Create a folder for uploads 	   ****
'**********************************************

Private Sub createUserFolder(ByVal strFolder)	

	Dim objFSO
	Dim objUserXMLfile
	Dim strFolderUserName
	Dim lngFolderUserID
	
	'Set error trapping
	On Error Resume Next
		
	'Creat an instance of the FSO object
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	'Check to see if an error has occurred
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("An error has occurred while uploading file/image.<br />Please check the File System Object (FSO) is installed on the server.", "create_FSO_object", "functions_upload.asp")

	'Disable error trapping
	On Error goto 0
	
	'If a folder doesn't exist for this user create one
	If NOT objFSO.FolderExists(Server.MapPath(strFolder)) Then
		
		'Get the user ID from the end of the file path
		lngFolderUserID = CLng(Right(strFolder, (Len(strFolder) - Instr(strFolder, "/"))))
		
		
		'If the user dosen't have a folder create them one
		'Make sure the folder doesn't already exsist (we already do this above, but some people still get an error, so we do it again)
		If Not objFSO.FolderExists(Server.MapPath(strFolder)) Then objFSO.CreateFolder(Server.MapPath(strFolder))
		
		
		'Read in the username of this user from the database as it is needed for the XML file containing data on the folder
		strSQL = "SELECT " & strDbTable & "Author.Username " & _
		"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
		"WHERE " & strDbTable & "Author.Author_ID = " & lngFolderUserID & ";"
	
		'Query the database
		rsCommon.Open strSQL, adoCon
		
		If NOT rsCommon.EOF Then strFolderUserName = rsCommon("Username")
		
		'Close RS
		rsCommon.Close
		
		
		'Create an XML file with user details; TODO, add feature to be able to add notes
		Set objUserXMLfile = objFSO.CreateTextFile(Server.MapPath(strFolder) & "\folder_info.xml", True) 
		
		'Add Contents
		objUserXMLfile.WriteLine("<?xml version=""1.0"" encoding=""utf-8""?>" & _
				vbCrLf & "<folder>" & _
				vbCrLf & " <created>" & internationalDateTime(Now()) & "</created>" & _
				vbCrLf & " <owner>" & _
				vbCrLf & "  <uid>" & lngFolderUserID & "</uid>" & _
				vbCrLf & "  <username>" & strFolderUserName & "</username>" & _
				vbCrLf & " </owner>" & _
				vbCrLf & "</folder>")
		
		'Close
		objUserXMLfile.Close
		Set objUserXMLfile = Nothing
		
	End If
	
	'Release the FSO object
	Set objFSO = Nothing
	
End Sub




'**********************************************
'***   Check if user has upload folder   ****
'**********************************************

Private Function userUploadFolder(ByVal strFolder)	

	Dim objFSO
	
	'Set error trapping
	On Error Resume Next
		
	'Creat an instance of the FSO object
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	'Check to see if an error has occurred
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("An error has occurred while uploading file/image.<br />Please check the File System Object (FSO) is installed on the server.", "create_FSO_object", "functions_upload.asp")

	'Disable error trapping
	On Error goto 0
	
	'If a folder doesn't exist for this user create one
	If objFSO.FolderExists(Server.MapPath(strFolder)) Then
		
		userUploadFolder = True
	Else
		userUploadFolder = False
		
	End If
	
	'Release the FSO object
	Set objFSO = Nothing
End Function




'**********************************************
'***  Check allocated space   ****
'**********************************************

Private Function folderSize(ByVal strFolder)	

	Dim objFSO
	
	'Set error trapping
	On Error Resume Next
		
	'Creat an instance of the FSO object
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	
	'Check to see if an error has occurred
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("An error has occurred while uploading file/image.<br />Please check the File System Object (FSO) is installed on the server.", "create_FSO_object", "functions_upload.asp")

	'Disable error trapping
	On Error goto 0
	
	'Now lets check the size of the folder (it's returned in bytes so converet to MB with 2 decimal places)
	folderSize = FormatNumber(CDbl(objFSO.GetFolder(Server.MapPath(strFolder)).Size / 1024 / 1024), 2)
	
	
	'Release the FSO object
	Set objFSO = Nothing
	
End Function




'******************************************
'***	AspJpeg supported image type   ****
'******************************************
Function AspJpegImage (ByVal strFileName)

	'Dimension varibles
	Dim intExtensionLoopCounter
	Dim saryAspJpegImageTypes(6)
	
	saryAspJpegImageTypes(0) = "jpeg"
	saryAspJpegImageTypes(1) = "gif" 
	saryAspJpegImageTypes(2) = "bmp" 
	saryAspJpegImageTypes(3) = "tiff"
	saryAspJpegImageTypes(4) = "png"
	saryAspJpegImageTypes(5) = "jpg"
	saryAspJpegImageTypes(6) = "tif"
	

	'Intilaise return value
	AspJpegImage = False

	'Loop through all the allowed extensions and see if the file has one
	For intExtensionLoopCounter = 0 To UBound(saryAspJpegImageTypes)

		If LCase(Right(strFileName, Len(saryAspJpegImageTypes(intExtensionLoopCounter))+1)) = "." & LCase(saryAspJpegImageTypes(intExtensionLoopCounter)) Then AspJpegImage = True
	Next

End Function
%>