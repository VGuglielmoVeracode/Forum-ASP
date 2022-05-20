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




'******************************************************
'***  Filters using 'HTML Secure (TM)' Technology *****
'******************************************************



'**********************************************
'***  Check HTML input for malicious code *****
'**********************************************

'Check input for tags and remove any that are not permitted for security reasons
Private Function HTMLsafe(ByVal strMessageInput)

	Dim strTempHTMLMessage		'Temporary message store
	Dim lngMessagePosition		'Holds the message position
	Dim intHTMLTagLength		'Holds the length of the HTML tags
	Dim strHTMLMessage		'Holds the HTML message
	Dim strTempMessageInput		'Temp store for the message input
	Dim lngLoopCounter		'Loop counter
	Dim strHyperlink		'Holds hyperlinks
	Dim strImageSrc			'Holds image src
	Dim strImageHeight		'Holds image height
	Dim strImageWidth		'Holds image Width
	Dim strImageBorder		'Holds image Border
	Dim strImageAlign		'Holds image Align
	Dim strImageAlt
	Dim strImageHSpace
	Dim strImageVSpace
	Dim strImageStyle
	Dim strImageTitle
	Dim intLoopCounter 	'Holds the loop counter


	'Include the array of unsafe HTML tags
	%><!--#include file="unsafe_HTML_tags_inc.asp" --><%


	'Strip scripting (this is just an extra check as these are stiped later (if in different format), but will give better formating of post if whole tag is striped now)
	strMessageInput = Replace(strMessageInput, "<script language=""javascript"">", "", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "<script language=javascript>", "", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "<script language=""vbscript"">", "", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "<script language=vbscript>", "", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "<script language=""jscript"">", "", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "<script language=jscript>", "", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "<script type=""text/javascript"">", "", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "<script type=text/javascript>", "", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "<script type=""text/vbscript"">", "", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "<script type=text/vbscript>", "", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "<script type=""text/jscript"">", "", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "<script type=text/jscript>", "", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "<script>", "", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "</script>", "", 1, -1, 1)


	'Strip dodgy styles (can be used to inject CSS into a page for XSS hacking exploit)
	strMessageInput = Replace(strMessageInput, "<style", "<", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "</style>", "", 1, -1, 1)
	
	'Strip dodgy video
	strMessageInput = Replace(strMessageInput, "<video", "<", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "</video>", "", 1, -1, 1)
	
	
	


	'Place the message input into a temp store
	strTempMessageInput = strMessageInput

	'Loop through each character in the post message looking for tags
	For lngMessagePosition = 1 to CLng(Len(strMessageInput))

		'If this is the end of the message then save some process time and jump out the loop
		If Mid(strMessageInput, lngMessagePosition, 1) = "" Then Exit For

		'If an HTML tag is found then move to the end of it so that we can strip the HTML tag and check it for malicious code
		If Mid(strMessageInput, lngMessagePosition, 1) = "<" Then
			

			'Get the length of the HTML tag
			intHTMLTagLength = (InStr(lngMessagePosition, strMessageInput, ">", 1) - lngMessagePosition)
			
			'Prevent errors by setting the HTML Tag Lenth end position to 0 if it is less than 0
			If intHTMLTagLength < 0 Then intHTMLTagLength = 0


			'Place the HTML tag back into the temporary message store
			strHTMLMessage = Mid(strMessageInput, lngMessagePosition, intHTMLTagLength + 1)

			'Place the HTML tag into a temporay variable store to be stripped of malcious code
			strTempHTMLMessage = strHTMLMessage
			
			'Response.write strTempHTMLMessage

			
			
			'Convert HTML encoding back into ASCII characters
			strTempHTMLMessage = removeHTMLencoding(strTempHTMLMessage)
			
			'If there is anymore HTML encoding left dump it
			strTempHTMLMessage = Replace(strTempHTMLMessage, "&#", "&amp;#", 1, -1, 1)
			
			
			
			'Remove ASCII non characters entities from 0 to 31
			For lngLoopCounter = 0 to 31
				strTempHTMLMessage = Replace(strTempHTMLMessage, CHR(lngLoopCounter), " ", 1, -1, 0)
			Next
			

			'***** Filter Hyperlinks *****

			'If this is an hyperlink tag then check it for malicious code
			If InStr(1, strTempHTMLMessage, "href", 1) <> 0 Then
				
				'Get just the href link
				strHyperlink = getHTMLProperty(strTempHTMLMessage, "href")
				
				'Call the format link function to strip malicious codes
				strHyperlink = formatLink(strHyperlink)
						
				'Rebuild the link
				If blnGroupURLs Then
					strTempHTMLMessage = "<a href=""" & strHyperlink & """"
					If blnLocalLinksSameWindow Then
						If InStr(strHyperlink, strForumPath) = 0 AND InStr(strHyperlink, strWebsiteURL) = 0 Then strTempHTMLMessage = strTempHTMLMessage & " target=""_blank"""
					Else
						strTempHTMLMessage = strTempHTMLMessage & " target=""_blank"""
					End If
					If blnNoFollowTagInLinks Then strTempHTMLMessage = strTempHTMLMessage & " rel=""nofollow"""
					strTempHTMLMessage = strTempHTMLMessage & ">"
				'Else build as text link
				Else
					strTempHTMLMessage = " " & strHyperlink & " - "
				End If
			End If	
			
			
			
			'***** Filter Image Tags *****
			
			'If this is an image then strip it of malicous code
			If InStr(1, strTempHTMLMessage, "img ", 1) <> 0 Then
			
				'Get the src image properties
				strImageSrc = getHTMLProperty(strTempHTMLMessage, "src")
				
				'If no image source then dump the img tag
				If strImageSrc = "" Then
					strTempHTMLMessage = ""
				
				'Filter the image and get the rest of it's properties
				Else
					'Call the check images function to strip malicious codes
					strImageSrc = checkImages(strImageSrc)
					
					'Get the rest of the image properties
					strImageHeight = getHTMLProperty(strTempHTMLMessage, "height")	
					strImageWidth = getHTMLProperty(strTempHTMLMessage, "width")
					strImageBorder = getHTMLProperty(strTempHTMLMessage, "border")
					strImageAlign = LCase(getHTMLProperty(strTempHTMLMessage, "align"))
					strImageAlt = getHTMLProperty(strTempHTMLMessage, "alt")
					strImageTitle = getHTMLProperty(strTempHTMLMessage, "title")
					strImageHSpace = getHTMLProperty(strTempHTMLMessage, "hspace")
					strImageVSpace = getHTMLProperty(strTempHTMLMessage, "vspace")
					'strImageStyle = getHTMLProperty(strTempHTMLMessage, "style") 'Styles can be used for XSS Hacking, it's best to leave these so they are removed
					
					'Filter alt, title and style input as no other checks can be done on these
					strImageAlt = removeAllTags(strImageAlt)
					strImageAlt = formatInput(strImageAlt)
					
					strImageTitle = removeAllTags(strImageTitle)
					strImageTitle = formatInput(strImageTitle)
					
					strImageStyle = removeAllTags(strImageStyle)
					strImageStyle = formatInput(strImageStyle)
					
		
					'Rebuild the image tag
					If blnGroupImages OR InStr(1, strImageSrc, "smileys/", 1) <> 0 Then
						strTempHTMLMessage = "<img src=""" & strImageSrc & """"
						If isNumeric(strImageHeight) Then strTempHTMLMessage = strTempHTMLMessage & " height=""" & strImageHeight & """"
						If isNumeric(strImageWidth) Then strTempHTMLMessage = strTempHTMLMessage & " width=""" & strImageWidth & """"
						If isNumeric(strImageHSpace) Then strTempHTMLMessage = strTempHTMLMessage & " hspace=""" & strImageHSpace & """"
						If isNumeric(strImageVSpace) Then strTempHTMLMessage = strTempHTMLMessage & " vspace=""" & strImageVSpace & """"
						If isNumeric(strImageBorder) Then strTempHTMLMessage = strTempHTMLMessage & " border=""" & strImageBorder & """" Else strTempHTMLMessage = strTempHTMLMessage & " border=""0"""
						If strImageAlign = "left" OR strImageAlign = "right" OR strImageAlign = "texttop" OR strImageAlign = "baseline" OR strImageAlign = "bottom" OR strImageAlign = "middle" OR strImageAlign = "top" Then strTempHTMLMessage = strTempHTMLMessage & " align=""" & strImageAlign & """"
						If strImageStyle <> "" Then strTempHTMLMessage = strTempHTMLMessage & " style=""" & strImageStyle & """"
						If strImageAlt <> "" Then 
							strTempHTMLMessage = strTempHTMLMessage & " alt=""" & strImageAlt & """"
							If strImageTitle = "" Then strTempHTMLMessage = strTempHTMLMessage & " title=""" & strImageAlt & """"
						End If
						If strImageTitle <> "" Then 
							strTempHTMLMessage = strTempHTMLMessage & " title=""" & strImageTitle & """"
							If strImageAlt = "" Then strTempHTMLMessage = strTempHTMLMessage & " alt=""" & strImageTitle & """"
						End If
						
						strTempHTMLMessage = strTempHTMLMessage & " />"	
						
					'Else the image is shown as a text link
					Else
						strTempHTMLMessage = " " & strImageSrc
						If strImageAlt <> "" Then strTempHTMLMessage = strTempHTMLMessage  & " - " & strImageAlt
					End If
					
					
                 
				End If
			End If


			'***** Filter Unwanted HTML Tags *****

			'If this is not an image or a link then cut all unwanted HTML out of the HTML tag
			If InStr(1, strTempHTMLMessage, "href", 1) = 0 AND InStr(1, strTempHTMLMessage, "img", 1) = 0 Then

				'Loop through the array of disallowed HTML tags
				For lngLoopCounter = LBound(saryUnSafeHTMLtags) To UBound(saryUnSafeHTMLtags)
					
					'If the disallowed HTML is found remove it and start over
					If Instr(1, strTempHTMLMessage,  saryUnSafeHTMLtags(lngLoopCounter), 1) Then
						
						'Remove the disallowed HTML
						strTempHTMLMessage = Replace(strTempHTMLMessage, saryUnSafeHTMLtags(lngLoopCounter), "", 1, -1, 1)
						
						'Start again as the hacker maybe placing maliciouse code around another disabllowed word to try and bypass the filter
						lngLoopCounter = 0
					End If
				Next
			End If



			'***** Format Unwanted HTML Tags *****

			'Extra check, Strip out malicious code from the HTML that may have not been stripped but trying to sneek through in a hyperlink or image src
			strTempHTMLMessage = formatInput(strTempHTMLMessage)
			
			
			'Remove any empty tags left after filtering
			strTempHTMLMessage = Replace(strTempHTMLMessage, "<>", "")
			strTempHTMLMessage = Replace(strTempHTMLMessage, "</>", "")


			'Place the new fromatted HTML tag back into the message post
			strTempMessageInput = Replace(strTempMessageInput, strHTMLMessage, strTempHTMLMessage, 1, -1, 1)

		End If
	Next

	'Return the function
	HTMLsafe = strTempMessageInput
End Function







'******************************************
'***  Get HTML tag single property    *****
'******************************************

'This function grabs a particular part of an HTML tag eg (href="get this part here")
Private Function getHTMLProperty(ByVal strHTMLtag, ByVal strHTMLproperty)
	
	Dim intPropertyStart
	Dim intPropertyEnd
	Dim strQuoteMarkChar1
	Dim strQuoteMarkChar2
	
	
	strHTMLtag = Replace(strHTMLtag, ">", " >")
	
	
	
	'First check to see if the part of the HTML tag we want to get actualy lives in the HTML tag
	If InStr(1, strHTMLtag, strHTMLproperty, 1) <> 0 Then
		
		
		'Find out what type of quote mark we are dealing with for this property eg. ' or "
		
		If InStr(InStr(1, strHTMLtag, strHTMLproperty, 1), strHTMLtag, strHTMLproperty & "=""", 1) <> 0 Then
			strQuoteMarkChar1 = """"
			strQuoteMarkChar2 = """"
		ElseIf InStr(InStr(1, strHTMLtag, strHTMLproperty, 1), strHTMLtag, strHTMLproperty & "='", 1) <> 0 Then 	
			strQuoteMarkChar1 = "'"
			strQuoteMarkChar2 = "'"
		ElseIf InStr(1, strHTMLtag, strHTMLproperty & "=", 1) <> 0 Then 	
			strQuoteMarkChar1 = ""
			strQuoteMarkChar2 = " "
		End If
		
		
		'Get where the part of the tag we want to look at starts
		intPropertyStart = InStr(InStr(1, strHTMLtag, strHTMLproperty, 1), strHTMLtag, strHTMLproperty & "=" & strQuoteMarkChar1, 1) + Len(strHTMLproperty & "=" & strQuoteMarkChar1)
		intPropertyEnd = InStr(intPropertyStart, strHTMLtag, strQuoteMarkChar2, 1)
		
						
		'If the start and end postions of the URL are correct then filter it
		If intPropertyEnd > intPropertyStart Then
					
			'Chop out everyting except the content of the property in question
			getHTMLProperty = Mid(strHTMLtag, intPropertyStart, intPropertyEnd-intPropertyStart)
			
			'Strip anymore quote marks and %0 (null) as they are not wanted in the return
			getHTMLProperty = Replace(getHTMLProperty, """", "", 1, -1, 1)
			getHTMLProperty = Replace(getHTMLProperty, "'", "", 1, -1, 1)
			getHTMLProperty = Replace(getHTMLProperty, "%22", "", 1, -1, 1)
			getHTMLProperty = Replace(getHTMLProperty, "%27", "", 1, -1, 1)
			getHTMLProperty = Replace(getHTMLProperty, "%0", "", 1, -1, 1)
							
		
		'This tag is not formatted correctly so return nothing
		Else
			getHTMLProperty = ""
		End If

	
	'Else the property is not in the tag so return nothing
	Else	
		getHTMLProperty = ""
	
	End If
	
End Function






'******************************************
'***  Check Images for malicious code *****
'******************************************

'Check images function
Private Function checkImages(ByVal strInputEntry)

	Dim strImageFileExtension	'Holds the file extension of the image
	Dim saryImageTypes		'Array holding allowed image types in the forum
	Dim intExtensionLoopCounter	'Holds the loop counter for the array
	Dim blnImageExtOK		'Set to true if the image extension is OK

	'If there is no . in the link then there is no extenison and so can't be an image
	If inStr(1, strInputEntry, ".", 1) = 0 Then

		strInputEntry = ""

	'Else remove malicious code and check the extension is an image extension
	Else

		'Initiliase variables
		blnImageExtOK = false

		'Get the file extension
		strImageFileExtension = LCase(Mid(strInputEntry, InStrRev(strInputEntry, "."), 4))

		'Get the image types allowed in the forum
		strImageTypes = strImageTypes & ";jpe;gif;jpg;bmp;png"

		'Place the image types into an array
		saryImageTypes = Split(Trim(strImageTypes), ";")

		'Loop through all the allowed extensions and see if the image has one
		For intExtensionLoopCounter = 0 To UBound(saryImageTypes)

			'Reformat extension to check
			saryImageTypes(intExtensionLoopCounter) = "." & Trim(Mid(saryImageTypes(intExtensionLoopCounter), 1, 3))

			'Check to see if the image extension is allowed
			If saryImageTypes(intExtensionLoopCounter) = strImageFileExtension Then blnImageExtOK = true
		Next

		'If the image extension is not OK then strip it from the image link
		If blnImageExtOK = false Then strInputEntry = Replace(strInputEntry, strImageFileExtension, "", 1, -1, 1)

		'Chop out any anything that is not normally found in an image URL
		strInputEntry = Replace(strInputEntry, "?", "", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, ";", "", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, "%3b", "", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, "{", "", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, "}", "", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, "%7b", "", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, "%7d", "", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, "%0", "", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, "^", "", 1, -1, 1)

		'URL Encode to prevent malicious code
		strInputEntry = Replace(strInputEntry, "(", "%28", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, ")", "%29", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, "[", "%5b", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, "]", "%5d", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, " ", "%20", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, "\", "%5C", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, Chr(9), "%09", 1, -1, 1) 'Tabs
		
		
		'Remove if the user is trying to use an FTP link
		strInputEntry = Replace(strInputEntry, "ftp://", "", 1, -1, 1)
	End If

	'Return
	checkImages = strInputEntry
End Function






'********************************************
'*** 		 Format Links 		*****
'********************************************

'Format links funtion
Private Function formatLink(ByVal strInputEntry)

	'URL Encode malisous characters from links and images
	strInputEntry = Replace(strInputEntry, """", "%22", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "'", "%27", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "(", "%28", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, ")", "%29", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "<", "%3c", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, ">", "%3e", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "[", "%5b", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "]", "%5d", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "{", "%7b", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "}", "%7d", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "\", "%5C", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, " ", "%20", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, Chr(9), "%09", 1, -1, 1) 'Tabs
	strInputEntry = Replace(strInputEntry, Chr(173), "%3c", 1, -1, 1) 'Vietmanise < tag
	
	'Remove a few bits
	strInputEntry = Replace(strInputEntry, "%0", "", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "^", "", 1, -1, 1)

	'Return
	formatLink = strInputEntry
End Function





'******************************************
'***  		Format user input     *****
'******************************************

'Format user input function
Private Function formatInput(ByVal strInputEntry)

	'Get rid of malicous code in the message
	strInputEntry = Replace(strInputEntry, Chr(9), "", 1, -1, 1) 'Remove Tabs
	strInputEntry = Replace(strInputEntry, "</script>", "", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "<script language=""javascript"">", "", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "<script language=javascript>", "", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "SCRIPT", "&#083;CRIPT", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Script", "&#083;cript", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "script", "&#115;cript", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "MOCHA", "&#077;OCHA", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Mocha", "&#077;ocha", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "mocha", "&#109;ocha", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "OBJECT", "&#079;BJECT", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Object", "&#079;bject", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "object", "&#111;bject", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "APPLET", "&#065;PPLET", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Applet", "&#065;pplet", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "applet", "&#097;pplet", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "ALERT", "&#065;LERT", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Alert", "&#065;lert", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "alert", "&#097;lert", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "EMBED", "&#069;MBED", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Embed", "&#069;mbed", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "embed", "&#101;mbed", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "EVENT", "&#069;VENT", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Event", "&#069;vent", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "event", "&#101;vent", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "DOCUMENT", "&#068;OCUMENT", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Document", "&#068;ocument", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "document", "&#100;ocument", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "COOKIE", "&#067;OOKIE", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Cookie", "&#067;ookie", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "cookie", "&#099;ookie", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "FORM", "&#070;ORM", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Form", "&#070;orm", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "form", "&#102;orm", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "IFRAME", "I&#070;RAME", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Iframe", "I&#102;rame", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "iframe", "i&#102;rame", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "TEXTAREA", "&#84;EXTAREA", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "Textarea", "&#84;extarea", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "textarea", "&#116;extarea", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "ON", "&#079;N", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "On", "&#079;n", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "on", "&#111;n", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "background-image", "background&ndash;image", 1, -1, 1)



	'Reformat a few bits
	strInputEntry = Replace(strInputEntry, "<STR&#079;NG>", "<strong>", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "<str&#111;ng>", "<strong>", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "</STR&#079;NG>", "</strong>", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "</str&#111;ng>", "</strong>", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "f&#111;nt", "font", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "F&#079;NT", "FONT", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "F&#111;nt", "Font", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "f&#079;nt", "font", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "f&#111;nt", "font", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "m&#111;no", "mono", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "M&#079;NO", "MONO", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "M&#111;no", "Mono", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "m&#079;no", "mono", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "m&#111;no", "mono", 1, -1, 1)

	'Return
	formatInput = strInputEntry
End Function






'********************************************
'*** 		 Format SQL input	*****
'********************************************

'Format SQL Query funtion
Private Function formatSQLInput(ByVal strInputEntry)

	'Remove malicous charcters from sql
	strInputEntry = Replace(strInputEntry, """", "", 1, -1, 1)
	
	'If this is mySQL need to get rid of the \ escape character and escape single quotes
	If strDatabaseType = "mySQL" Then
		strInputEntry = Replace(strInputEntry, "\", "\\", 1, -1, 1)
		strInputEntry = Replace(strInputEntry, "'", "\'", 1, -1, 1)
	'Else for Access and SQL server need to escape a single quote using two quotes
	Else
		strInputEntry = Replace(strInputEntry, "'", "''", 1, -1, 1)
	End If
	
	strInputEntry = Replace(strInputEntry, "[", "&#091;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "]", "&#093;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "<", "&lt;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, ">", "&gt;", 1, -1, 1)
	
	'Return
	formatSQLInput = strInputEntry
End Function





'*********************************************
'***  		Strip all tags		 *****
'*********************************************

'Remove all tags for text only display 
Private Function removeAllTags(ByVal strInputEntry)

	'Remove all HTML scripting tags etc. for plain text output
	strInputEntry = Replace(strInputEntry, "&", "&amp;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "<", "&lt;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, ">", "&gt;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "'", "&#039;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, """", "&quot;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "\", "&#092;", 1, -1, 1)
	'Fix for UK Sterling £
	strInputEntry = Replace(strInputEntry, Chr(194) & Chr(163), "&pound;", 1, -1, 1)  'Fix for  Â£
	strInputEntry = Replace(strInputEntry, Chr(163), "&pound;", 1, -1, 1)
	


	'Return
	removeAllTags = strInputEntry
End Function






'******************************************
'***  Non-Alphanumeric Character Strip ****
'******************************************

'Function to strip non alphanumeric characters
Private Function characterStrip(strTextInput)

	'Dimension variable
	Dim intLoopCounter 	'Holds the loop counter

	'Loop through the ASCII characters
	For intLoopCounter = 0 to 47
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0)
	Next

	'Loop through the ASCII characters numeric characters to lower-case characters
	For intLoopCounter = 91 to 96
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0)
	Next

	'Loop through the extended ASCII characters
	For intLoopCounter = 58 to 64
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0)
	Next

	'Loop through the extended ASCII characters
	For intLoopCounter = 123 to 255
		strTextInput = Replace(strTextInput, CHR(intLoopCounter), "", 1, -1, 0)
	Next


	'Return the string
	characterStrip = strTextInput

End Function




'********************************************
'***  Email Address Validation		 ****
'********************************************

'Function to validate emaiol address
Private Function emailAddressValidation(strEmailAddress)

	'Dimension variable
	Dim intLoopCounter 	'Holds the loop counter
	
	'Trim and change to lower case
	strEmailAddress = Trim(LCase(strEmailAddress))
	
	'Replace double dots
	strEmailAddress = Replace(strEmailAddress, "..", ".")
	
	'Remove dots from GMail accounts (dots are ignored by GMail, so spammers place random dots in GMail address to register mutiple times with same email)
	If Instr(strEmailAddress, "gmail.com") Then
		
		'Strip dots from GMail address
		strEmailAddress = Replace(strEmailAddress, ".", "")
		
		'Place back dot in to gmail.com
		strEmailAddress = Replace(strEmailAddress, "gmailcom", "gmail.com")
	End If
	
	'Loop through the ASCII characters
	For intLoopCounter = 0 to 37
		strEmailAddress = Replace(strEmailAddress, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Loop through the ASCII characters
	For intLoopCounter = 39 to 42
		strEmailAddress = Replace(strEmailAddress, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Remove single ASCII character
	strEmailAddress = Replace(strEmailAddress, CHR(44), "", 1, -1, 0)
	
	'Loop through the ASCII characters
	For intLoopCounter = 58 to 60
		strEmailAddress = Replace(strEmailAddress, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Remove single ASCII character
	strEmailAddress = Replace(strEmailAddress, CHR(62), "", 1, -1, 0)
	
	'Loop through the ASCII characters numeric characters to lower-case characters
	For intLoopCounter = 65 to 94
		strEmailAddress = Replace(strEmailAddress, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Remove single ASCII character
	strEmailAddress = Replace(strEmailAddress, CHR(96), "", 1, -1, 0)
	
	'Loop through the extended ASCII characters
	For intLoopCounter = 123 to 125
		strEmailAddress = Replace(strEmailAddress, CHR(intLoopCounter), "", 1, -1, 0) 
	Next
	
	'Loop through the extended ASCII characters
	For intLoopCounter = 127 to 255
		strEmailAddress = Replace(strEmailAddress, CHR(intLoopCounter), "", 1, -1, 0) 
	Next


	'Check whats left of the email address is valid
	If Len(strEmailAddress) < 5 OR NOT Instr(1, strEmailAddress, " ") = 0 OR InStr(1, strEmailAddress, "@", 1) < 2 OR InStrRev(strEmailAddress, ".") < InStr(1, strEmailAddress, "@", 1) Then strEmailAddress = ""


	'Return the string
	emailAddressValidation = strEmailAddress

End Function





'**********************************************
'*** 		 Strip HTML 		  *****
'**********************************************

'Remove HTML function
Private Function removeHTML(ByVal strMessageInput, ByVal lngReturnLength, ByVal blnRemoveBRtags)

	Dim objRegExp	'Holds regulare expresions object
	
	'Extra error handling
	If isNull(strMessageInput) Then strMessageInput = ""

	'Remove edit XML
	strMessageInput = Replace(strMessageInput, "<editID>", "<br /><br />" & strTxtEditBy & " ", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "</editID>", " ", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "<editDate>", " ", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "</editDate>", " ", 1, -1, 1)


	'If we want <br /> tags to remain the best thing to do is remove carridge returns,
	'then repleace <br /> tags with carridge returns that get changed back to <br /> tags once
	'the HTML has been striped

	'If leaving in <br /> tags
	If blnRemoveBRtags = false Then
		'Change <br /> tags into nulls so they are not striped and can be changed back later
		strMessageInput = Replace(strMessageInput, "<br />", "#cB2JcmbXp#6!ZZ$", 1, -1, 1)
		strMessageInput = Replace(strMessageInput, "<br>", "#cB2JcmbXp#6!ZZ$", 1, -1, 1)
	End If



	'Create regular experssions object
	Set objRegExp = New RegExp
	
	'Tell the regular experssions object to look for tags [HIDE][/HIDE] as tehse need to be striped
	With objRegExp
		.Pattern = "\[HIDE\][^\]]+\[/HIDE\]"
		.IgnoreCase = True
		.Global = True
	End With

	'Strip HTML
	strMessageInput = objRegExp.Replace(strMessageInput, "")

	'Tell the regular experssions object to look for tags <xxxx>
	With objRegExp
		.Pattern = "<[^>]+>"
		.IgnoreCase = True
		.Global = True
	End With

	'Strip HTML
	strMessageInput = objRegExp.Replace(strMessageInput, "")
	


	'Tell the regular experssions object to look for BB forum codes [xxxx]
	With objRegExp
		.Pattern = "\[[^\]]+\]"
		.IgnoreCase = True
		.Global = True
	End With

	'Strip BB forum codes
	strMessageInput = objRegExp.Replace(strMessageInput, "")

	'Distroy regular experssions object
	Set objRegExp = nothing


	'Replace a few characters in the remaining text
	strMessageInput = Replace(strMessageInput, "<", "&lt;", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, ">", "&gt;", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "'", "&#039;", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, """", "&#034;", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, "&nbsp;", "", 1, -1, 1)

	'If the length of the text is longer than the max
	strMessageInput = TrimString(strMessageInput, lngReturnLength)
	

	'Remove new lines as it's better for display as link titles
	strMessageInput = Replace(strMessageInput, vbCrLf, " ", 1, -1, 1)
	strMessageInput = Replace(strMessageInput, vbCr, " ", 1, -1, 1)

	'Place back in <br /> tags
	If blnRemoveBRtags = false Then strMessageInput = Replace(strMessageInput, "#cB2JcmbXp#6!ZZ$", vbCrLf & "       <br />", 1, -1, 1)

	'Return the function
	removeHTML = strMessageInput
End Function







'*********************************************
'*** Decode HTML encoding for plain text *****
'*********************************************

'Decode encoded strings
Private Function decodeString(ByVal strInputEntry)

	'Prevent errors
	If isNull(strInputEntry) Then strInputEntry = ""

	'Decode HTML character entities
	
	strInputEntry = Replace(strInputEntry, "&#065;", "A", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#066;", "B", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#067;", "C", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#068;", "D", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#069;", "E", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#070;", "F", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#071;", "G", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#072;", "H", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#073;", "I", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#074;", "J", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#075;", "K", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#076;", "L", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#077;", "M", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#078;", "N", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#079;", "O", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#080;", "P", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#081;", "Q", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#082;", "R", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#083;", "S", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#084;", "T", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#085;", "U", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#086;", "V", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#087;", "W", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#088;", "X", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#089;", "Y", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#090;", "Z", 1, -1, 0)

	strInputEntry = Replace(strInputEntry, "&#097;", "a", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#098;", "b", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#099;", "c", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#100;", "d", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#101;", "e", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#102;", "f", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#103;", "g", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#104;", "h", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#105;", "i", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#106;", "j", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#107;", "k", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#108;", "l", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#109;", "m", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#110;", "n", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#111;", "o", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#112;", "p", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#113;", "q", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#114;", "r", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#115;", "s", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#116;", "t", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#117;", "u", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#118;", "v", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#119;", "w", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#120;", "x", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#121;", "y", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#122;", "z", 1, -1, 0)


	strInputEntry = Replace(strInputEntry, "&#048;", "0", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#049;", "1", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#050;", "2", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#051;", "3", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#052;", "4", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#053;", "5", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#054;", "6", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#055;", "7", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#056;", "8", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#057;", "9", 1, -1, 0)


	'Non aplha numeric characters
	strInputEntry = Replace(strInputEntry, "&#039;", "'", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#39;", "'", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#061;", "=", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#61;", "=", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#091;", "[", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "&#91;", "[", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "&#092;", "\", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "&#92;", "\", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "&#093;", "]", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "&#93;", "]", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "&#146;", "'", 1, -1, 1)
	
	
	'Decode other entities
	strInputEntry = Replace(strInputEntry, "&amp;", "&", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "&lt;", "<", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "&gt;", ">", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "&#039;", "'", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "&quot;", """", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "&#092;", "\", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "&#091;", "[", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "&#093;", "]", 1, -1, 1)


	'Return
	decodeString = strInputEntry
End Function









'*********************************************
'***  	   Remove HTML encoding		 *****
'*********************************************

'Remove HTML encoding of ASCII characters A-z 0-10
Private Function removeHTMLencoding(ByVal strInputEntry)

	'Prevent errors
	If isNull(strInputEntry) Then strInputEntry = ""

	'Decode HTML character entities
	strInputEntry = Replace(strInputEntry, "&#065;", "A", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#066;", "B", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#067;", "C", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#068;", "D", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#069;", "E", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#070;", "F", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#071;", "G", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#072;", "H", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#073;", "I", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#074;", "J", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#075;", "K", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#076;", "L", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#077;", "M", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#078;", "N", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#079;", "O", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#080;", "P", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#081;", "Q", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#082;", "R", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#083;", "S", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#084;", "T", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#085;", "U", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#086;", "V", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#087;", "W", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#088;", "X", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#089;", "Y", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#090;", "Z", 1, -1, 0)

	strInputEntry = Replace(strInputEntry, "&#097;", "a", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#098;", "b", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#099;", "c", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#100;", "d", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#101;", "e", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#102;", "f", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#103;", "g", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#104;", "h", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#105;", "i", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#106;", "j", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#107;", "k", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#108;", "l", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#109;", "m", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#110;", "n", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#111;", "o", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#112;", "p", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#113;", "q", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#114;", "r", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#115;", "s", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#116;", "t", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#117;", "u", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#118;", "v", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#119;", "w", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#120;", "x", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#121;", "y", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#122;", "z", 1, -1, 0)


	strInputEntry = Replace(strInputEntry, "&#048;", "0", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#049;", "1", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#050;", "2", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#051;", "3", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#052;", "4", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#053;", "5", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#054;", "6", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#055;", "7", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#056;", "8", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#057;", "9", 1, -1, 0)
	
	'Repeat with 0
	strInputEntry = Replace(strInputEntry, "&#65;", "A", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#66;", "B", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#67;", "C", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#68;", "D", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#69;", "E", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#70;", "F", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#71;", "G", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#72;", "H", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#73;", "I", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#74;", "J", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#75;", "K", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#76;", "L", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#77;", "M", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#78;", "N", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#79;", "O", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#80;", "P", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#81;", "Q", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#82;", "R", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#83;", "S", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#84;", "T", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#85;", "U", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#86;", "V", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#87;", "W", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#88;", "X", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#89;", "Y", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#90;", "Z", 1, -1, 0)

	strInputEntry = Replace(strInputEntry, "&#97;", "a", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#98;", "b", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#99;", "c", 1, -1, 0)


	strInputEntry = Replace(strInputEntry, "&#48;", "0", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#49;", "1", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#50;", "2", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#51;", "3", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#52;", "4", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#53;", "5", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#54;", "6", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#55;", "7", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#56;", "8", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#57;", "9", 1, -1, 0)
	
	'Special characters
	strInputEntry = Replace(strInputEntry, "&#32;", " ", 1, -1, 0)
	strInputEntry = Replace(strInputEntry, "&#032;", " ", 1, -1, 0)

	'Return
	removeHTMLencoding = strInputEntry
End Function





'*********************************************
'***  	   SQL Injection Test		 *****
'*********************************************

'SQL Injection Test
Private Sub SqlInjectionTest(strData) 

	Dim blnSqlClear
	Dim sarySqlKeyword(32)
	Dim lngLoopCounter
	
	blnSqlClear = True
	
	'SQL Unsafe keywords
	sarySqlKeyword(0) = "select"
	sarySqlKeyword(1) = "drop"
	sarySqlKeyword(2) = "insert"
	sarySqlKeyword(3) = "set"
	sarySqlKeyword(4) = "alter"
	sarySqlKeyword(5) = "table"
	sarySqlKeyword(6) = "values"
	sarySqlKeyword(7) = "cast"
	sarySqlKeyword(8) = "declare"	
	sarySqlKeyword(9) = "char"
	sarySqlKeyword(10) = "order"
	sarySqlKeyword(11) = "create"
	sarySqlKeyword(12) = "rollback"
	sarySqlKeyword(13) = "savepoint"
	sarySqlKeyword(14) = "commit"
	sarySqlKeyword(15) = "limit"
	sarySqlKeyword(16) = "where"
	sarySqlKeyword(17) = "join"
	sarySqlKeyword(18) = "having"
	sarySqlKeyword(19) = "update"
	sarySqlKeyword(20) = "delete"
	sarySqlKeyword(21) = "begin"
	sarySqlKeyword(22) = "transaction"
	sarySqlKeyword(23) = "key"
	sarySqlKeyword(24) = "primary"
	sarySqlKeyword(25) = "grant"
	sarySqlKeyword(26) = "trigger"
	sarySqlKeyword(27) = "union"
	sarySqlKeyword(28) = "truncate"
	sarySqlKeyword(29) = "merge"
	sarySqlKeyword(30) = "cursor"
	sarySqlKeyword(31) = "index"
	sarySqlKeyword(32) = "exec"
	
	
	'Loop through the array of disallowed SQL Keywords
	For lngLoopCounter = LBound(sarySqlKeyword) To UBound(sarySqlKeyword)
						
		'If SQL keyword is found update 
		If Instr(1, strData,  sarySqlKeyword(lngLoopCounter), 1) Then
			blnSqlClear = False
		End If
	Next
	
	'If an error has occurred write an error to the page
	If blnSqlClear = False Then Call errorMsg("WARNING: SQL Injection attack detected.", "SqlInjectionTest()", "functions_common.asp")
End Sub




'*************************************
'*** SEO Friendly URL Titles   *****
'**************************************

'for URL rewrite search engine friendly page titles
Private Function SeoUrlTitle(ByVal strInputEntry, strPrefix)

	Dim intLoopCounter
	Dim objRegExp

	If blnSeoTitleQueryStrings = False Then Exit Function
	
	'Swap to lower case
	strInputEntry = strInputEntry
	
	'Remove any HTML encoding
	strInputEntry = decodeString(strInputEntry)

	strInputEntry = Replace(strInputEntry, "_", " ", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, ".", " ", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "/", " ", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "+", " ", 1, -1, 1)
	
	
	'Create regular experssions object
	Set objRegExp = New RegExp

	'Tell the regular experssions object to look for tags <xxxx>
	With objRegExp
		.Pattern = "[^\w\d\s]"  'same as [^a-zA-Z0-9 ]
		.IgnoreCase = True
		.Global = True
	End With

	'Strip HTML
	strInputEntry = objRegExp.Replace(strInputEntry, "")

	'Distroy regular experssions object
	Set objRegExp = nothing
	
	
	'Trim the final result
	strInputEntry = Trim(strInputEntry)
	
	'Replace spaces with hyphans
	strInputEntry = Replace(strInputEntry, " ", "-", 1, -1, 1)
	
	'Replace double hyphens
	strInputEntry = Replace(strInputEntry, "---", "-", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "--", "-", 1, -1, 1)
	
	
	'Return result (if any)
	If strInputEntry = "" Then 
		SeoUrlTitle = ""
	Else
		SeoUrlTitle = LCase(strPrefix & strInputEntry)
	End If

End Function


%>