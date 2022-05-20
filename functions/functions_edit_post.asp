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
'***   HTML to Forum Codes Function   *****
'******************************************

'Edit Post Function to convert posts back to forum codes
Private Function EditPostConvertion(ByVal strMessage)

	Dim strTempMessage	'Temporary word hold for e-mail and url words
	Dim strMessageLink	'Holds the new mesage link that needs converting back into code
	Dim lngStartPos		'Holds the start position for a link
	Dim lngEndPos		'Holds the end position for a word
	Dim intLoop		'Loop counter
	Dim intLoopBreakout     'Breakout of loop counter
	
	'Extra error handling
	If isNull(strMessage) Then strMessage = ""
	
	strMessage = Replace(strMessage, " target=""_blank""", "", 1, -1, 1)
	strMessage = Replace(strMessage, "", "", 1, -1, 1)
	strMessage = Replace(strMessage, " border=""0""", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<img src= """, "<img src=""", 1, -1, 1)
	strMessage = Replace(strMessage, "<a href= """, "<a href=""", 1, -1, 1)
	
	
	
	'Change the path to the emotion symbols back into the emotion codes
	For intLoop = 1 to UBound(saryEmoticons)
		strMessage = Replace(strMessage, "<img alt=""" & saryEmoticons(intLoop,1) & """ src=""" & saryEmoticons(intLoop,3) & """ align=""middle"">", saryEmoticons(intLoop,2), 1, -1, 1)
		strMessage = Replace(strMessage, "<img src=""" & saryEmoticons(intLoop,3) & """ align=""middle"">", saryEmoticons(intLoop,2), 1, -1, 1)
		strMessage = Replace(strMessage, "<img src=""" & saryEmoticons(intLoop,3) & """>", saryEmoticons(intLoop,2), 1, -1, 1)
	Next
	
	
	
	'If the message has been edited remove who edited the post
	If InStr(1, strMessage, "<edited>", 1) Then strMessage = removeEditorAuthor(strMessage)
	
	
	'Change the HTML codes back into my own codes for bold and italic
	strMessage = Replace(strMessage, "<b>", "[B]", 1, -1, 1)
	strMessage = Replace(strMessage, "</b>", "[/B]", 1, -1, 1)
	strMessage = Replace(strMessage, "<i>", "[I]", 1, -1, 1)
	strMessage = Replace(strMessage, "</i>", "[/I]", 1, -1, 1)
	strMessage = Replace(strMessage, "<u>", "[U]", 1, -1, 1)
	strMessage = Replace(strMessage, "</u>", "[/U]", 1, -1, 1)
	
	strMessage = Replace(strMessage, "<hr />", "[HR]", 1, -1, 1)
	strMessage = Replace(strMessage, "<hr>", "[HR]", 1, -1, 1)
	strMessage = Replace(strMessage, "<hr>", "[HR]", 1, -1, 1)
	strMessage = Replace(strMessage, "<ol>", "[LIST=1]", 1, -1, 1)
	strMessage = Replace(strMessage, "</ol>", "[/LIST=1]", 1, -1, 1)
	strMessage = Replace(strMessage, "<ul>", "[LIST]", 1, -1, 1)
	strMessage = Replace(strMessage, "</ul>", "[/LIST]", 1, -1, 1)
	strMessage = Replace(strMessage, "<li>", "[LI]", 1, -1, 1)
	strMessage = Replace(strMessage, "</li>", "[/LI]", 1, -1, 1)
	strMessage = Replace(strMessage, "<center>", "[CENTER]", 1, -1, 1)
	strMessage = Replace(strMessage, "</center>", "[/CENTER]", 1, -1, 1)
	
	strMessage = Replace(strMessage, "<strike>", "[STRIKE]", 1, -1, 1)
	strMessage = Replace(strMessage, "</strike>", "[/STRIKE]", 1, -1, 1)
	strMessage = Replace(strMessage, "<sub>", "[SUB]", 1, -1, 1)
	strMessage = Replace(strMessage, "</sub>", "[/SUB]", 1, -1, 1)
	strMessage = Replace(strMessage, "<sup>", "[SUP]", 1, -1, 1)
	strMessage = Replace(strMessage, "</sup>", "[/SUP]", 1, -1, 1)
	
	strMessage = Replace(strMessage, "<strong>", "[B]", 1, -1, 1)
	strMessage = Replace(strMessage, "</strong>", "[/B]", 1, -1, 1)
	strMessage = Replace(strMessage, "<em>", "[I]", 1, -1, 1)
	strMessage = Replace(strMessage, "</em>", "[/I]", 1, -1, 1)
	
	strMessage = Replace(strMessage, "<br />", vBCrLf, 1, -1, 1)
	strMessage = Replace(strMessage, "<br>", vBCrLf, 1, -1, 1)
	
	strMessage = Replace(strMessage, "<pre 100>", "[PRE]", 1, -1, 1)
	strMessage = Replace(strMessage, "<pre>", "[PRE]", 1, -1, 1)
	strMessage = Replace(strMessage, "</pre>", "[/PRE]", 1, -1, 1)
	
	strMessage = Replace(strMessage, "<P>", "[P]", 1, -1, 1)
	strMessage = Replace(strMessage, "</P>", "[/P]", 1, -1, 1)
	strMessage = Replace(strMessage, "<P align=center>", "[P ALIGN=CENTER]", 1, -1, 1)
	strMessage = Replace(strMessage, "<P align=justify>", "[P ALIGN=JUSTIFY]", 1, -1, 1)
	strMessage = Replace(strMessage, "<P align=left>", "[P ALIGN=LEFT]", 1, -1, 1)
	strMessage = Replace(strMessage, "<P align=right>", "[P ALIGN=RIGHT]", 1, -1, 1)
	
	strMessage = Replace(strMessage, "<div>", "[DIV]", 1, -1, 1)
	strMessage = Replace(strMessage, "</div>", "[/DIV]", 1, -1, 1)
	strMessage = Replace(strMessage, "<div align=""center"">", "[DIV ALIGN=CENTER]", 1, -1, 1)
	strMessage = Replace(strMessage, "<div align=""justify"">", "[DIV ALIGN=JUSTIFY]", 1, -1, 1)
	strMessage = Replace(strMessage, "<div align=""left"">", "[DIV ALIGN=LEFT]", 1, -1, 1)
	strMessage = Replace(strMessage, "<div align=""right"">", "[DIV ALIGN=RIGHT]", 1, -1, 1)
	strMessage = Replace(strMessage, "<div align=center>", "[DIV ALIGN=CENTER]", 1, -1, 1)
	strMessage = Replace(strMessage, "<div align=justify>", "[DIV ALIGN=JUSTIFY]", 1, -1, 1)
	strMessage = Replace(strMessage, "<div align=left>", "[DIV ALIGN=LEFT]", 1, -1, 1)
	strMessage = Replace(strMessage, "<div align=right>", "[DIV ALIGN=RIGHT]", 1, -1, 1)
	
	strMessage = Replace(strMessage, "<blockquote>", "[BLOCKQUOTE]", 1, -1, 1)
	strMessage = Replace(strMessage, "</blockquote>", "[/BLOCKQUOTE]", 1, -1, 1)
	
	strMessage = Replace(strMessage, "<font size=""1"">", "[SIZE=1]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font size=""2"">", "[SIZE=2]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font size=""3"">", "[SIZE=3]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font size=""4"">", "[SIZE=4]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font size=""5"">", "[SIZE=5]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font size=""6"">", "[SIZE=6]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font size=6>", "[SIZE=6]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font size=1>", "[SIZE=1]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font size=2>", "[SIZE=2]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font size=3>", "[SIZE=3]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font size=4>", "[SIZE=4]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font size=5>", "[SIZE=5]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font size=6>", "[SIZE=6]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font face=""Arial, Helvetica, sans-serif"">", "[FONT=Arial]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font face=""Courier New, Courier, mono"">", "[FONT=Courier]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font face=""Times New Roman, Times, serif"">", "[FONT=Times]", 1, -1, 1)
	strMessage = Replace(strMessage, "<font face=""Verdana, Arial, Helvetica, sans-serif"">", "[FONT=Verdana]", 1, -1, 1)
	
	
	'Reset veriable
	intLoopBreakout = 0
	
	'Loop through the message till all or any IMAGE links are converted back into BBcodes
	Do While InStr(1, strMessage, "<img ", 1) > 0 AND intLoopBreakout < 1000
						    	
		'Find the start position in the image tag
		lngStartPos = InStr(1, strMessage, "<img ", 1)
															
		'Find the position in the message for the image closing tag
		lngEndPos = InStr(lngStartPos, strMessage, "/>", 1) + 3
		
		'Make sure the end position is not in error
		If lngEndPos - lngStartPos =< 10 Then lngEndPos = lngStartPos + 10
						
		'Read in the code to be converted back into the forum codes
		strMessageLink = Trim(Mid(strMessage, lngStartPos, (lngEndPos - lngStartPos)))	
		
		'Place the image tag into the tempoary message variable
		strTempMessage = strMessageLink
		
		'Format the HTML image tag back into forum codes
		strTempMessage = Replace(strTempMessage, "src=""", "", 1, -1, 1)
		strTempMessage = Replace(strTempMessage, "<img ", "[IMG]", 1, -1, 1)
		strTempMessage = Replace(strTempMessage, """ />", "[/IMG]", 1, -1, 1)
		
		'Place the new fromatted codes into the message string body
		strMessage = Replace(strMessage, strMessageLink, strTempMessage, 1, -1, 1)	
		
		'increment loop,break out counter
		intLoopBreakout = intLoopBreakout + 1	
	Loop
	
	
	
	'Reset veriable
	intLoopBreakout = 0
	
	'Loop through the message till all or any HTML email links are converted back into codes
	Do While InStr(1, strMessage, "<a href=""mailto:", 1) > 0 AND InStr(1, strMessage, "</a>", 1) > 0 AND intLoopBreakout < 1000
						    	
		'Find the start position in the message of the HTML e-mail mailto tag
		lngStartPos = InStr(1, strMessage, "<a href=""mailto:", 1)
									
									
		'Find the position in the message for the </a> closing code
		lngEndPos = InStr(lngStartPos, strMessage, "</a>", 1) + 4
		
		'Make sure the end position is not in error
		If lngEndPos - lngStartPos =< 16 Then lngEndPos = lngStartPos + 16
						
		'Read in the code to be converted back into the forum codes
		strMessageLink = Trim(Mid(strMessage, lngStartPos, (lngEndPos - lngStartPos)))	
		
		'Place the message link into the tempoary message variable
		strTempMessage = strMessageLink
		
		'Format the HTML mailto link back into forum codes
		strTempMessage = Replace(strTempMessage, "<a href=""mailto:", "[EMAIL=", 1, -1, 1)
		strTempMessage = Replace(strTempMessage, "</a>", "[/EMAIL]", 1, -1, 1)
		strTempMessage = Replace(strTempMessage, """>", "]", 1, -1, 1)
		
		'Place the new fromatted codes into the message string body
		strMessage = Replace(strMessage, strMessageLink, strTempMessage, 1, -1, 1)	
		
		'increment loop,break out counter
		intLoopBreakout = intLoopBreakout + 1	
	Loop
	
	
	'Reset veriable
	intLoopBreakout = 0
	
	'Loop through the message till all or any hyperlinks are turned into back into froum codes
	Do While InStr(1, strMessage, "<a href=""", 1) > 0 AND InStr(1, strMessage, "</a>", 1) > 0 AND intLoopBreakout < 1000
						    	
		'Find the start position in the message of the HTML hyperlink
		lngStartPos = InStr(1, strMessage, "<a href=""", 1)
																	
		'Find the position in the message for the </a> closing code
		lngEndPos = InStr(lngStartPos, strMessage, "</a>", 1) + 4
		
		'Make sure the end position is not in error
		If lngEndPos - lngStartPos =< 9 Then lngEndPos = lngStartPos + 9
						
		'Read in the code to be converted back into forum codes from the message
		strMessageLink = Trim(Mid(strMessage, lngStartPos, (lngEndPos - lngStartPos)))	
		
		'Place the message link into the tempoary message variable
		strTempMessage = strMessageLink
		
		'Format the HTML hyperlink back into forum codes
		strTempMessage = Replace(strTempMessage, " rel=""nofollow""", "", 1, -1, 1)
		strTempMessage = Replace(strTempMessage, "<a href=""", "[URL=", 1, -1, 1)
		strTempMessage = Replace(strTempMessage, "</a>", "[/URL]", 1, -1, 1)
		strTempMessage = Replace(strTempMessage, """>", "]", 1, -1, 1)
		
		'Place the new fromatted codes into the message string body
		strMessage = Replace(strMessage, strMessageLink, strTempMessage, 1, -1, 1)	
		
		'increment loop,break out counter
		intLoopBreakout = intLoopBreakout + 1	
	Loop
	
	
	'Reset veriable
	intLoopBreakout = 0
	
	'Loop through the message till all font colour tags are converted back to forum codes
	Do While InStr(1, strMessage, "<font color=", 1) > 0 AND InStr(1, strMessage, "</font>", 1) > 0 AND intLoopBreakout < 1000
		
		
						    	
		'Find the start position in the message of the HTML colour tag
		lngStartPos = InStr(1, strMessage, "<font color=", 1)
									
									
		'Find the position in the message for the </font> closing code
		lngEndPos = InStr(lngStartPos, strMessage, "</font>", 1) + 8
		
		'Make sure the end position is not in error
		If lngEndPos - lngStartPos =< 12 Then lngEndPos = lngStartPos + 12
						
		'Read in the code to be converted back into the forum codes
		strMessageLink = Trim(Mid(strMessage, lngStartPos, (lngEndPos - lngStartPos)))	
		
		'Place the message link into the tempoary message variable
		strTempMessage = strMessageLink
		
		'Format the HTML colour tag back into forum codes
		strTempMessage = Replace(strTempMessage, "<font color=""", "[COLOR=", 1, -1, 1)
		strTempMessage = Replace(strTempMessage, "</font>", "[/COLOR]", 1, -1, 1)
		strTempMessage = Replace(strTempMessage, """>", "]", 1, -1, 1)
		
		'Place the new fromatted codes into the message string body
		strMessage = Replace(strMessage, strMessageLink, strTempMessage, 1, -1, 1)
		
		'increment loop,break out counter
		intLoopBreakout = intLoopBreakout + 1
	Loop
	
	'Turn any left over font tages to forum codes
	strMessage = Replace(strMessage, "</font>", "[/FONT]", 1, -1, 1)
	
	
	'Turn the HTML back into the charcaters entred by the user
	strMessage = Replace(strMessage, "<", "&lt;", 1, -1, 1)
	strMessage = Replace(strMessage, ">", "&gt;", 1, -1, 1)
	strMessage = Replace(strMessage, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", "       ", 1, -1, 1)
	strMessage = Replace(strMessage, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", "      ", 1, -1, 1)
	strMessage = Replace(strMessage, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", "     ", 1, -1, 1)
	strMessage = Replace(strMessage, "&nbsp;&nbsp;&nbsp;&nbsp;", "    ", 1, -1, 1)
	strMessage = Replace(strMessage, "&nbsp;&nbsp;&nbsp;", "   ", 1, -1, 1)
	strMessage = Replace(strMessage, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", vbTab, 1, -1, 1)
	strMessage = Replace(strMessage, Chr(10), "", 1, -1, 1)
	
	
	
	'Return function
	EditPostConvertion = strMessage
	
End Function






'******************************************
'*** Remove Post Editor Text Function *****
'******************************************

'Format Post Function to covert forum codes to HTML
Private Function removeEditorAuthor(ByVal strMessage)

	Dim lngStartPos
	Dim lngEndPos
		
	'Get the start and end position in the string of the XML to remove
	lngStartPos = InStr(1, strMessage, "<edited>", 1)
	lngEndPos = InStr(1, strMessage, "</edited>", 1) + 9
	If lngEndPos - lngStartPos =< 8 Then lngEndPos = lngStartPos + 9

	'If there is something returned strip the XML from the message
	removeEditorAuthor = Replace(strMessage, Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos)), "", 1, -1, 1)
		
End Function
%>