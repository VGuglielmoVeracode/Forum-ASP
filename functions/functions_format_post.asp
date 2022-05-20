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
'***    Strip entites from RTE posts   *****
'******************************************

Private Function WYSIWYGFormatPost(ByVal strMessage)

	'Format messages that use the WYSIWYG Editor
	strMessage = Replace(strMessage, " border=0>", ">", 1, -1, 1)
	strMessage = Replace(strMessage, " target=_blank>", ">", 1, -1, 1)
	strMessage = Replace(strMessage, " target=_top>", ">", 1, -1, 1)
	strMessage = Replace(strMessage, " target=_self>", ">", 1, -1, 1)
	strMessage = Replace(strMessage, " target=_parent>", ">", 1, -1, 1)
	strMessage = Replace(strMessage, " style=""CURSOR: hand""", "", 1, -1, 1)
	
	'Strip wordTidy tags
	strMessage = Replace(strMessage, "<wordTidy>", "", 1, -1, 1)
	strMessage = Replace(strMessage, "</wordTidy>", "", 1, -1, 1)
	
	
	'Strip out add blocking injection code
	
	'Strip MS Word 12 Bloat
	strMessage = Replace(strMessage, "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<meta name=""ProgId"" content=""Word.Document"">", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<meta name=""Generator"" content=""Microsoft Word 12"">", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<meta name=""Originator"" content=""Microsoft Word 12"">", "", 1, -1, 1)
	
	
	'Strip out code injected by Badly behaved Firefox plugins
	strMessage = Replace(strMessage, "<input type=""hidden"" id=""gwProxy"" />", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<input type=""hidden"" onclick=""jsCall();"" id=""jsProxy"" />", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<!--Session data-->", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<div id=""refHTML"">&nbsp;</div>", "", 1, -1, 1)
	
	'Strip out code injected by BitDefender
	strMessage = Replace(strMessage, "<script language=""javascript"" type=""text/javascript"" src=""/B1D671CF-E532-4481-99AA-19F420D90332/netdefender/hui/ndhui.js?0&amp;0&amp;0""></script>", "", 1, -1, 1)
	
	'Strip out Norton Internet Security pop up add blocking injected code
	strMessage = Replace(strMessage, "<SCRIPT> window.open=NS_ActualOpen; </SCRIPT>", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<SCRIPT language=javascript>postamble();</SCRIPT>", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<SCRIPT language=""javascript"">postamble();</SCRIPT>", "", 1, -1, 1)
	
	'Strip out Zone Alarm Pro's pop up add blocking injected code (bloody pain in the arse crap software)
	If Instr(1, strMessage, "<!-- ZoneLabs Popup Blocking Insertion -->", 1) Then
		strMessage = Replace(strMessage, "<!-- ZoneLabs Popup Blocking Insertion -->", "", 1, -1, 1)
		strMessage = Replace(strMessage, "<SCRIPT>" & vbCrLf & "window.open=NS_ActualOpen;" & vbCrLf & "orig_onload = window.onload;" & vbCrLf & "orig_onunload = window.onunload;" & vbCrLf & "window.onload = noopen_load;" & vbCrLf & "window.onunload = noopen_unload;" & vbCrLf & "</SCRIPT>", "", 1, -1, 1)
		strMessage = Replace(strMessage, "window.open=NS_ActualOpen; orig_onload = window.onload; orig_onunload = window.onunload; window.onload = noopen_load; window.onunload = noopen_unload;", "", 1, -1, 1)
	End If
	
	'Strip out Norton Personal Firewall 2003's pop up add blocking injected code
	strMessage = Replace(strMessage, "<!--" & vbCrLf & vbCrLf & "window.open = SymRealWinOpen;" & vbCrLf & vbCrLf & "//-->", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<!--" & vbCrLf & vbCrLf & "function SymError()" & vbCrLf & "{" & vbCrLf & "  return true;" & vbCrLf & "}" & vbCrLf & vbCrLf & "window.onerror = SymError;" & vbCrLf & vbCrLf & "//-->", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<!--" & vbCrLf & vbCrLf & "function SymError()" & vbCrLf & "{" & vbCrLf & "  return true;" & vbCrLf & "}" & vbCrLf & vbCrLf & "window.onerror = SymError;" & vbCrLf & vbCrLf & "var SymRealWinOpen = window.open;" & vbCrLf & vbCrLf & "function SymWinOpen(url, name, attributes)" & vbCrLf & "{" & vbCrLf & "  return (new Object());" & vbCrLf & "}" & vbCrLf & vbCrLf & "window.open = SymWinOpen;" & vbCrLf & vbCrLf & "//-->", "", 1, -1, 1)

	'Strip out Kerio Firewall pop up add blocking injected code (now Sunbelt)
	strMessage = Replace(strMessage, "<!-- Kerio Popup Killer - script has been appended by KPF -->", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<!-- Sunbelt Kerio Popup Killer -  has been appended by KPF -->", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<iframe id=""kpfLogFrame"" src=""http://127.0.0.1:44501/pl.html?START_LOG"" onload=""destroyIframe(this)"" style=""display:none;"">" & vbCrLf & vbCrLf & "</iframe>", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<iframe id=""kpfLogFrame"" src=""http://localhost:44501/pl.html?START_LOG"" onload=""destroyIframe(this)"" style=""display: none;"">" & vbCrLf & "</iframe>", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<iframe id=""kpfLogFrame"" src=""http://127.0.0.1:44501/pl.html?START_LOG"" onload=""destroyIframe(this)"" style=""display: none;"">" & vbCrLf & "</iframe>", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<script type=""text/javascript"">" & vbCrLf & "<!--" & vbCrLf & "	nopopups();" & vbCrLf & "//-->" & vbCrLf & "</script>", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<script type=""text/javascript"">" & vbCrLf & "<!--" & vbCrLf & "nopopups();" & vbCrLf & "//-->" & vbCrLf & "</script>", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<!-- Sunbelt Kerio Popup Killer - end of the  appended by KPF-->", "", 1, -1, 1)
	strMessage = Replace(strMessage, "<!-- Kerio Popup Killer - end of the script appended by KPF-->", "", 1, -1, 1)
	
	'Strip linux firewall for my LAN that injects this for ad blocking
	strMessage = Replace(strMessage, "<script>function PrivoxyWindowOpen(a, b, c){return(window.open(a, b, c));}</script>", "", 1, -1, 1)
	
	'Strip Foxcub translator
	strMessage = Replace(strMessage, "<iframe id=""foxcub_translator_popup_win"" style=""visibility: hidden; width: 0px; height: 0px; ""></iframe>", "", 1, -1, 1) 
	
	
	'Strip out CA Personal Firewall pop up add blocking injected code
	strMessage = Replace(strMessage, "<script type=""text/javascript"">_popupControl();</script>", "", 1, -1, 1)
	
	'Correct issue with UK Sterling Pound £ not display correctly
	strMessage = Replace(strMessage, Chr(194) & Chr(163), "&pound;", 1, -1, 1) 'Fix for  Â£
	strMessage = Replace(strMessage, Chr(163), "&pound;", 1, -1, 1)
	
	

	'Return the function
	WYSIWYGFormatPost = strMessage

End Function



'******************************************
'***        Format Post Function      *****
'******************************************

'Format Post Function to covert HTML tags into safe tags
Private Function FormatPost(ByVal strMessage)

	'Format spaces and HTML
	strMessage = Replace(strMessage, "<", "&lt;", 1, -1, 1)
	strMessage = Replace(strMessage, ">", "&gt;", 1, -1, 1)
	strMessage = Replace(strMessage, "       ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strMessage = Replace(strMessage, "      ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strMessage = Replace(strMessage, "     ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strMessage = Replace(strMessage, "    ", "&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strMessage = Replace(strMessage, "   ", "&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strMessage = Replace(strMessage, vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;", 1, -1, 1)
	strMessage = Replace(strMessage, Chr(13), "", 1, -1, 1)
	strMessage = Replace(strMessage, Chr(10), "<br />", 1, -1, 1)
	'Fix for UK Sterling £
	strMessage = Replace(strMessage, Chr(194) & Chr(163), "&pound;", 1, -1, 1)  'Fix for  Â£
	strMessage = Replace(strMessage, Chr(163), "&pound;", 1, -1, 1)
	
	
	'Return the function
	FormatPost = strMessage

End Function





'******************************************
'***   Format Forum Codes Function    *****
'******************************************

'Format Forum Codes Function to covert forum codes to HTML
Private Function FormatForumCodes(ByVal strMessage)


	Dim strTempMessage	'Temporary word hold for e-mail, fonts, and url words
	Dim strMessageLink	'Holds the new mesage link that needs converting back into code
	Dim lngStartPos		'Holds the start position for a link
	Dim lngEndPos		'Holds the end position for a word
	Dim intLoop		'Loop counter
	Dim objRegExp		'Holds regulare expresions object



	'If emoticons are on then change the emotion symbols for the path to the relative smiley icon
	If blnEmoticons = True Then
		'Loop through the emoticons array
		For intLoop = 1 to UBound(saryEmoticons)
			strMessage = Replace(strMessage, saryEmoticons(intLoop,2), "<img src=""" & saryEmoticons(intLoop,3) & """ align=""middle"">", 1, -1, 1)
		Next
	End If



	'Change forum codes for bold and italic HTML tags back to the normal satandard HTML tags
	strMessage = Replace(strMessage, "[B]", "<strong>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/B]", "</strong>", 1, -1, 1)
	strMessage = Replace(strMessage, "[STRONG]", "<strong>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/STRONG]", "</strong>", 1, -1, 1)
	strMessage = Replace(strMessage, "[I]", "<em>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/I]", "</em>", 1, -1, 1)
	strMessage = Replace(strMessage, "[EM]", "<em>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/EM]", "</em>", 1, -1, 1)
	strMessage = Replace(strMessage, "[U]", "<u>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/U]", "</u>", 1, -1, 1)
	
	strMessage = Replace(strMessage, "[HR]", "<hr />", 1, -1, 1)
	strMessage = Replace(strMessage, "[LIST=1]", "<ol>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/LIST=1]", "</ol>", 1, -1, 1)
	strMessage = Replace(strMessage, "[LIST]", "<ul>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/LIST]", "</ul>", 1, -1, 1)
	strMessage = Replace(strMessage, "[LI]", "<li>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/LI]", "</li>", 1, -1, 1)
	strMessage = Replace(strMessage, "[CENTER]", "<center>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/CENTER]", "</center>", 1, -1, 1)
	
	
	strMessage = Replace(strMessage, "[STRIKE]", "<strike>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/STRIKE]", "</strike>", 1, -1, 1)
	strMessage = Replace(strMessage, "[SUB]", "<sub>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/SUB]", "</sub>", 1, -1, 1)
	strMessage = Replace(strMessage, "[SUP]", "<sup>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/SUP]", "</sup>", 1, -1, 1)
	
	
	strMessage = Replace(strMessage, "[BR]", "<br />", 1, -1, 1)
	
	strMessage = Replace(strMessage, "[PRE]", "<pre 100>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/PRE]", "</pre>", 1, -1, 1)
	
	strMessage = Replace(strMessage, "[P]", "<p>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/P]", "</p>", 1, -1, 1)
	strMessage = Replace(strMessage, "[P ALIGN=CENTER]", "<p align=""center"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[P ALIGN=JUSTIFY]", "<p align=""justify"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[P ALIGN=LEFT]", "<p align=""left"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[P ALIGN=RIGHT]", "<p align=""right"">", 1, -1, 1)
	
	strMessage = Replace(strMessage, "[DIV]", "<div>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/DIV]", "</div>", 1, -1, 1)
	strMessage = Replace(strMessage, "[DIV ALIGN=CENTER]", "<div align=""center"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[DIV ALIGN=JUSTIFY]", "<div align=""justify"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[DIV ALIGN=LEFT]", "<div align=""left"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[DIV ALIGN=RIGHT]", "<div align=""right"">", 1, -1, 1)
	
	strMessage = Replace(strMessage, "[BLOCKQUOTE]", "<blockquote>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/BLOCKQUOTE]", "</blockquote>", 1, -1, 1)
	
	strMessage = Replace(strMessage, "[SIZE=1]", "<font size=""1"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[SIZE=2]", "<font size=""2"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[SIZE=3]", "<font size=""3"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[SIZE=4]", "<font size=""4"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[SIZE=5]", "<font size=""5"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[SIZE=6]", "<font size=""6"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[/SIZE]", "</font>", 1, -1, 1)
	
	strMessage = Replace(strMessage, "[FONT=Arial]", "<font face=""Arial, Helvetica, sans-serif"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[FONT=Courier]", "<font face=""Courier New, Courier, mono"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[FONT=Times]", "<font face=""Times New Roman, Times, serif"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[FONT=Verdana]", "<font face=""Verdana, Arial, Helvetica, sans-serif"">", 1, -1, 1)
	strMessage = Replace(strMessage, "[/FONT]", "</font>", 1, -1, 1)


	'Create regular experssions object
	Set objRegExp = New RegExp
	
	'Covert [IMG] BBcode in to HTML
	With objRegExp
	        .Pattern = "\[IMG\]\s*([^\]]+)\s*\[\/IMG\]"
	        .IgnoreCase = True
	        .Global = True
	End With
	strMessage = objRegExp.Replace(strMessage, "<img src=""$1"" />")
	
	
	'Covert [URL] BBcode in to HTML
	With objRegExp
		.Pattern = "\[URL\]\s*([^\]]+)\s*\[\/URL\]"
		.IgnoreCase = True
		.Global = True
	End With
	strMessage = objRegExp.Replace(strMessage, "<a href=""$1"">$1</a>")


	'Covert [URL=xxxx]xxx[/URL] BBcode in to HTML
	With objRegExp
		'.Pattern = "\[URL=\s*([^\]]+)\s*\]\s*([^\]]+)\s*\[\/URL\]"  'Orginal
		.Pattern = "\[URL=\s*\s*""?\s*([^\]]+)\s*""?\s*\s*\]\s*([^\]]+)\s*\[\/URL\]"  'Submitted by Jagoop
		.IgnoreCase = True
		.Global = True
	End With
	strMessage = objRegExp.Replace(strMessage, "<a href=""$1"">$2</a>")


	'Covert [EMAIL=xxxx]xxx[/EMAIL] BBcode in to HTML
	With objRegExp
		.Pattern = "\[EMAIL=\s*([^\]]+)\s*\]\s*([^\]]+)\s*\[\/EMAIL\]"
		.IgnoreCase = True
		.Global = True
	End With
	strMessage = objRegExp.Replace(strMessage, "<a href=""mailto:$1"">$2</a>")
	

	'Covert [FILE=xxxx]xxx[/FILE] BBcode in to HTML
	With objRegExp
		.Pattern = "\[FILE=\s*([^\]]+)\s*\]\s*([^\]]+)\s*\[\/FILE\]"
		.IgnoreCase = True
		.Global = True
	End With
	strMessage = objRegExp.Replace(strMessage, "<a href=""$1"">$2</a>")
	
	
	'Covert [COLOR=xxxx]xxx[/COLOR] BBcode in to HTML
	With objRegExp
		.Pattern = "\[COLOR=\s*([^\]]+)\s*\]\s*([^\]]+)\s*\[\/COLOR\]"
		.IgnoreCase = True
		.Global = True
	End With
	strMessage = objRegExp.Replace(strMessage, "<font color=""$1"">$2</font>")
	
	
	
	
	'Clean up
	Set objRegExp = nothing
	


	'Return the function
	FormatForumCodes = strMessage
End Function





'******************************************
'***   	   Format User Quote		***
'******************************************

'This function formats quotes that contain usernames
Function formatUserQuote(ByVal strMessage)


	'Declare variables
	Dim strQuotedAuthor 	'Holds the name of the author who is being quoted
	Dim strQuotedMessage	'Hold the quoted message
	Dim lngStartPos		'Holds search start postions
	Dim lngEndPos		'Holds end start postions
	Dim strBuildQuote	'Holds the built quoted message
	Dim strOriginalQuote	'Holds the quote in original format

	'Loop through all the quotes in the message and convert them to formated quotes
	Do While InStr(1, strMessage, "[QUOTE=", 1) > 0 AND InStr(1, strMessage, "[/QUOTE]", 1) > 0


		'Get the start and end in the message of the author who is being quoted
		lngStartPos = InStr(1, strMessage, "[QUOTE=", 1) + 7
		lngEndPos = InStr(lngStartPos, strMessage, "]", 1)

		'If there is something returned get the authors name
		If lngStartPos > 6 AND lngEndPos > 0 Then
			strQuotedAuthor = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))
		End If


		'Get the start and end in the message of the message to quote
		lngStartPos = lngStartPos + Len(strQuotedAuthor) + 1
		lngEndPos = InStr(lngStartPos, strMessage, "[/QUOTE]", 1)

		'Make sure the end position is not in error
		If lngEndPos - lngStartPos =< 0 Then lngEndPos = lngStartPos + Len(strQuotedAuthor)

		'If there is something returned get message to quote
		If lngEndPos > lngStartPos Then

			'Get the message to be quoted
			strQuotedMessage = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))

			'Srip out any perenetis for those that are use to BBcodes that are different
			strQuotedAuthor = Replace(strQuotedAuthor, """", "", 1, -1, 1)

			'Build the HTML for the displying of the quoted message
			strBuildQuote = vbCrLf & "<table width=""99%""><tr><td class=""BBquote""><img src=""" & strImagePath & "quote_box." & strForumImageType & """ title=""" & strTxtOriginallyPostedBy & " " & strQuotedAuthor & """ alt=""" & strTxtOriginallyPostedBy & " " & strQuotedAuthor & """ style=""vertical-align: text-bottom;"" /> <strong>" & strQuotedAuthor & " " & strTxtWrote & ":</strong><br /><br />" & strQuotedMessage & "</td></tr></table>"
		End If



		'Get the start and end position in the start and end position in the message of the quote
		lngStartPos = InStr(1, strMessage, "[QUOTE=", 1)
		lngEndPos = InStr(lngStartPos, strMessage, "[/QUOTE]", 1) + 8

		'Make sure the end position is not in error
		If lngEndPos - lngStartPos =< 7 Then lngEndPos = lngStartPos + Len(strQuotedAuthor) + 8

		'Get the original quote to be replaced in the message
		strOriginalQuote = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))

		'Replace the quote codes in the message with the new formated quote
		If strBuildQuote <> "" Then
			strMessage = Replace(strMessage, strOriginalQuote, strBuildQuote, 1, -1, 1)
		Else
			strMessage = Replace(strMessage, strOriginalQuote, Replace(strOriginalQuote, "[", "&#91;", 1, -1, 1), 1, -1, 1)
		End If
	Loop

	'Return the function
	formatUserQuote = strMessage

End Function




'******************************************
'***   	   Format Quote			***
'******************************************

'This function formats the quote
Function formatQuote(ByVal strMessage)


	'Declare variables
	Dim strQuotedMessage	'Hold the quoted message
	Dim lngStartPos		'Holds search start postions
	Dim lngEndPos		'Holds end start postions
	Dim strBuildQuote	'Holds the built quoted message
	Dim strOriginalQuote	'Holds the quote in original format

	'Loop through all the quotes in the message and convert them to formated quotes
	Do While InStr(1, strMessage, "[QUOTE]", 1) > 0 AND InStr(1, strMessage, "[/QUOTE]", 1) > 0

		'Get the start and end in the message of the author who is being quoted
		lngStartPos = InStr(1, strMessage, "[QUOTE]", 1) + 7
		lngEndPos = InStr(lngStartPos, strMessage, "[/QUOTE]", 1)

		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 7

		'If there is something returned get message to quote
		If lngEndPos > lngStartPos Then

			'Get the message to be quoted
			strQuotedMessage = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))

                        
			'Build the HTML for the displying of the quoted message
			strBuildQuote = vbCrLf & "<table width=""99%""><tr><td class=""BBquote""><img src=""" & strImagePath & "quote_box." & strForumImageType & """ title=""" & strTxtQuote & """ alt=""" & strTxtQuote & """ style=""vertical-align: text-bottom;"" /> " & strQuotedMessage & "</td></tr></table>"
		End If  
                        
                        
		'Get the start and end position in the start and end position in the message of the quote
		lngStartPos = InStr(1, strMessage, "[QUOTE]", 1)
		lngEndPos = InStr(lngStartPos, strMessage, "[/QUOTE]", 1) + 8

		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 7

		'Get the original quote to be replaced in the message
		strOriginalQuote = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))

		'Replace the quote codes in the message with the new formated quote
		If strBuildQuote <> "" Then
			strMessage = Replace(strMessage, strOriginalQuote, strBuildQuote, 1, -1, 1)
		Else
			strMessage = Replace(strMessage, strOriginalQuote, Replace(strOriginalQuote, "[", "&#91;", 1, -1, 1), 1, -1, 1)
		End If
	Loop

	'Return the function
	formatQuote = strMessage

End Function





'******************************************
'***   	   Format Code Block		***
'******************************************

'This function formats the code blocks
Function formatCode(ByVal strMessage)


	'Declare variables
	Dim strCodeMessage		'Hold the coded message
	Dim lngStartPos			'Holds search start postions
	Dim lngEndPos			'Holds end start postions
	Dim strBuildCodeBlock		'Holds the built coded message
	Dim strOriginalCodeBlock	'Holds the code block in original format

	'Loop through all the codes in the message and convert them to formated code block
	Do While InStr(1, strMessage, "[CODE]", 1) > 0 AND InStr(1, strMessage, "[/CODE]", 1) > 0
	
		'Get the start and end in the message of the author who is being coded
		lngStartPos = InStr(1, strMessage, "[CODE]", 1) + 6
		lngEndPos = InStr(lngStartPos, strMessage, "[/CODE]", 1)

		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 6

		'If there is something returned get message to code block
		If lngEndPos > lngStartPos Then

			'Get the message to be coded
			strCodeMessage = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))
			
			'Build the HTML for the displaying of the code
			strBuildCodeBlock = vbCrLf & "<table width=""99%""><tr><td><pre class=""BBcode"">" & strCodeMessage & "</pre></td></tr></table>"
		End If


		'Get the start and end position in the start and end position in the message of the code block
		lngStartPos = InStr(1, strMessage, "[CODE]", 1)
		lngEndPos = InStr(lngStartPos, strMessage, "[/CODE]", 1) + 7

		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 6

		'Get the original code to be replaced in the message
		strOriginalCodeBlock = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))

		'Replace the code codes in the message with the new formated code block
		If strBuildCodeBlock <> "" Then
			strMessage = Replace(strMessage, strOriginalCodeBlock, strBuildCodeBlock, 1, -1, 1)
		Else
			strMessage = Replace(strMessage, strOriginalCodeBlock, Replace(strOriginalCodeBlock, "[", "&#91;", 1, -1, 1), 1, -1, 1)
		End If
	Loop

	'Return the function
	formatCode = strMessage

End Function





'******************************************
'***   	   Format Spoiler Block		***
'******************************************

'This function formats the spoiler blocks
Function formatSpoiler(ByVal strMessage, blnRssFeed)


	'Declare variables
	Dim strSpoilerMessage		'Hold the spoiler message
	Dim lngStartPos			'Holds search start postions
	Dim lngEndPos			'Holds end start postions
	Dim strBuildSpoilerBlock	'Holds the built spoiler message
	Dim strOriginalSpoilerBlock	'Holds the spoiler block in original format

	'Loop through all the Spoilers in the message and convert them to formated Spoiler block
	Do While InStr(1, strMessage, "[SPOILER]", 1) > 0 AND InStr(1, strMessage, "[/SPOILER]", 1) > 0
	
		'Get the start and end in the message of the spoiler block
		lngStartPos = InStr(1, strMessage, "[SPOILER]", 1) + 9
		lngEndPos = InStr(lngStartPos, strMessage, "[/SPOILER]", 1)

		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 9

		'If there is something returned get message to Spoiler block
		If lngEndPos > lngStartPos Then

			'Get the message to be coded
			strSpoilerMessage = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))
			
			'Build the HTML for the displaying of the Spoiler in rss feeds
			If blnRssFeed Then
				strBuildSpoilerBlock = vbCrLf & "<table width=""99%""><tr><td><div class=""BBquote""><img src=""" & strImagePath & "spoiler." & strForumImageType & """ title=""" & strTxtSpoiler & """ alt=""" & strTxtSpoiler & """ style=""vertical-align: text-bottom;"" /> <strong>" & strTxtSpoiler & "</strong> <br />" & strSpoilerMessage & "</td></tr></table>"
	
			'Build the HTML for the displaying of the Spoiler
			Else
				'Build the HTML for the displaying of the Spoiler
				strBuildSpoilerBlock = vbCrLf & "<table width=""99%""><tr><td><div class=""BBquote""><img src=""" & strImagePath & "spoiler." & strForumImageType & """ title=""" & strTxtSpoiler & """ alt=""" & strTxtSpoiler & """ style=""vertical-align: text-bottom;"" /> <strong>" & strTxtSpoiler & "</strong> <input type=""button"" value=""" & strTxtShow & """ style=""text-align:center;width:100px;margin:0px;padding:0px;"" onclick=""if (this.parentNode.parentNode.getElementsByTagName('div')[1].getElementsByTagName('div')[0].style.display != '') { this.parentNode.parentNode.getElementsByTagName('div')[1].getElementsByTagName('div')[0].style.display = '';      this.innerText = ''; this.value = '" & strTxtHide & "'; } else { this.parentNode.parentNode.getElementsByTagName('div')[1].getElementsByTagName('div')[0].style.display = 'none'; this.innerText = ''; this.value = '" & strTxtShow & "'; }"" /><div><div style=""display: none;""><br />" & strSpoilerMessage & "</div></div></div></td></tr></table>"
			End If
		End If


		'Get the start and end position in the start and end position in the message of the Spoiler block
		lngStartPos = InStr(1, strMessage, "[SPOILER]", 1)
		lngEndPos = InStr(lngStartPos, strMessage, "[/SPOILER]", 1) + 10

		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 9

		'Get the original Spoiler to be replaced in the message
		strOriginalSpoilerBlock = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))

		'Replace the Spoiler Spoilers in the message with the new formated Spoiler block
		If strBuildSpoilerBlock <> "" Then
			strMessage = Replace(strMessage, strOriginalSpoilerBlock, strBuildSpoilerBlock, 1, -1, 1)
		Else
			strMessage = Replace(strMessage, strOriginalSpoilerBlock, Replace(strOriginalSpoilerBlock, "[", "&#91;", 1, -1, 1), 1, -1, 1)
		End If
	Loop

	'Return the function
	formatSpoiler = strMessage

End Function






'******************************************
'***   		Format Signature	***
'******************************************

'This function formats falsh codes
Function formatSignature(ByVal strSignature)

	Dim strSignatureArea
	
	'Create signature
	strSignatureArea = "" & _
	vbCrLf & "   <!-- Start Signature -->" & _
	vbCrLf & "    <div class=""msgSignature"">" & _
	vbCrLf & "     " & strSignature & _
	vbCrLf & "    </div>" & _
	vbCrLf & "   <!-- End Signature ""'' -->"

	'Return the function
	formatSignature = strSignatureArea

End Function




'******************************************
'***   	Format Flash File Support	***
'******************************************

'This function formats falsh codes
Function formatFlash(ByVal strMessage)


	'Declare variables
	Dim lngStartPos		'Holds search start postions
	Dim lngEndPos		'Holds end start postions
	Dim saryFlashAttributes 'Holds the features of the input flash file
	Dim intAttrbuteLoop	'Holds the attribute loop counter
	Dim strFlashWidth	'Holds the string value of the width of the Flash file
	Dim intFlashWidth	'Holds the interger value of the width of the flash file
	Dim strFlashHeight	'Holds the string value of the height of the Flash file
	Dim intFlashHeight	'Holds the interger value of the height of the flash file
	Dim strBuildFlashLink	'Holds the converted BBcode for the flash file
	Dim strTempFlashMsg	'Tempoary store for the BBcode
	Dim strFlashLink	'Holds the link to the flash file



	'Loop through all the codes in the message and convert them to formated flash links
	Do While InStr(1, strMessage, "[FLASH", 1) > 0 AND InStr(1, strMessage, "[/FLASH]", 1) > 0

		'Initiliase variables
		intFlashWidth = 250
		intFlashHeight = 250
		strFlashLink = ""
		strBuildFlashLink = ""
		strTempFlashMsg = ""

		'Get the Flash BBcode from the message
		lngStartPos = InStr(1, strMessage, "[FLASH", 1)
		lngEndPos = InStr(lngStartPos, strMessage, "[/FLASH]", 1) + 8

		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 6

		'Get the original Flash BBcode from the message
		strTempFlashMsg = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))




		'Get the start and end in the message of the attributes of the Flash file
		lngStartPos = InStr(1, strTempFlashMsg, "[FLASH", 1) + 6
		lngEndPos = InStr(lngStartPos, strTempFlashMsg, "]", 1)

		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos

		'If there is something returned get the details (eg. dimensions) of the flash file
		If strTempFlashMsg <> "" Then

			'Place any attributes for the flash file in an array
			saryFlashAttributes = Split(Trim(Mid(strTempFlashMsg, lngStartPos, lngEndPos-lngStartPos)), " ")

			'Get the dimensions of the Flash file
			'Loop through the array of atrributes that are for the falsh file to get the dimentions
			For intAttrbuteLoop = 0 To UBound(saryFlashAttributes)

				'If this is the width attribute then read in the width dimention
				If InStr(1, saryFlashAttributes(intAttrbuteLoop), "WIDTH=", 1) Then

					'Get the width dimention
					strFlashWidth = Replace(saryFlashAttributes(intAttrbuteLoop), "WIDTH=", "", 1, -1, 1)

					'Make sure we are left with a numeric number if so convert to an interger and place in an interger variable
					If isNumeric(strFlashWidth) Then intFlashWidth = CInt(strFlashWidth)
				End If

				'If this is the height attribute then read in the height dimention
				If InStr(1, saryFlashAttributes(intAttrbuteLoop), "HEIGHT=", 1) Then

					'Get the height dimention
					strFlashHeight = Replace(saryFlashAttributes(intAttrbuteLoop), "HEIGHT=", "", 1, -1, 1)

					'Make sure we are left with a numeric number if so convert to an interger and place in an interger variable
					If isNumeric(strFlashHeight) Then intFlashHeight = CInt(strFlashHeight)
				End If
			Next



			'Get the link to the flash file
			lngStartPos = InStr(1, strTempFlashMsg, "]", 1) + 1
			lngEndPos = InStr(lngStartPos, strTempFlashMsg, "[/FLASH]", 1)

			'Make sure the end position is not in error
			If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 8

			'Read in the code to be converted into a hyperlink from the message
			strFlashLink = Trim(Mid(strTempFlashMsg, lngStartPos, (lngEndPos - lngStartPos)))


			'Build the HTML for the displying of the flash file
			If strFlashLink <> "" Then
				strBuildFlashLink = "<embed src=""" & strFlashLink & """"
				strBuildFlashLink = strBuildFlashLink & " allowScriptAccess=""never"" allowNetworking=""internal"" quality=""high"" width=" & intFlashWidth & " height=" & intFlashHeight & " type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash""></embed>"
			End If
		End If



		'Replace the flash codes in the message with the new formated flash link
		If strBuildFlashLink <> "" Then
			strMessage = Replace(strMessage, strTempFlashMsg, strBuildFlashLink, 1, -1, 1)
		Else
			strMessage = Replace(strMessage, strTempFlashMsg, Replace(strTempFlashMsg, "[", "&#91;", 1, -1, 1), 1, -1, 1)
		End If
	Loop

	'Return the function
	formatFlash = strMessage

End Function






'******************************************
'***   	  YouTube Support		***
'******************************************

'This function formats YouTube
Function formatYouTube(ByVal strMessage)


	'Declare variables
	Dim strYouTubeLink		'Hold the You Tube Link
	Dim lngStartPos			'Holds search start postions
	Dim lngEndPos			'Holds end start postions
	Dim strBuildYouTube		'Holds the built coded message
	Dim strOriginalYouTube	'Holds the code block in original format

	'Loop through all the BB codes in the message and convert to a link to the YouTube movie
	Do While InStr(1, strMessage, "[TUBE]", 1) > 0 AND InStr(1, strMessage, "[/TUBE]", 1) > 0
	
		'Get the start and end of the YouTube BBcode
		lngStartPos = InStr(1, strMessage, "[TUBE]", 1) + 6
		lngEndPos = InStr(lngStartPos, strMessage, "[/TUBE]", 1)

		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 6

		'If there is a YouTube link then process
		If lngEndPos > lngStartPos Then

			'Get the YouTube link
			strYouTubeLink = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))
			
			'Replace watch?v= with v/ for those copy and pasting links
			strYouTubeLink = Replace(Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos)), "watch?v=", "v/")
			
			'YouTube have changed to HTTPS, if URL is old change to HTTPS.
			strYouTubeLink = Replace(strYouTubeLink, "http://", "https://")
			
			'Remove http://youtu.be from copy and paste share links 
			strYouTubeLink = Replace(strYouTubeLink, "https://youtu.be/", "")
			
			'See if the YouTube link contains the whole URL or just the file name
			If InStr(1, strYouTubeLink, "https://", 1) = 0 Then strYouTubeLink = "https://www.youtube.com/embed/" & strYouTubeLink
			
			'If EU Cookie Law enabled replace www.youtube.com with www.youtube-nocookie.com
			If blnEuCookieLaw Then strYouTubeLink = Replace(strYouTubeLink, "www.youtube.com", "www.youtube-nocookie.com")
			
			'Insert youTube movie
			strBuildYouTube = "<iframe width=""560"" height=""350"" src=""" & strYouTubeLink & """  frameborder=""0"" allowfullscreen></iframe>"
		End If

		
		'Get the start and end position in the start and end position in the message of the BBcode YouTube
		lngStartPos = InStr(1, strMessage, "[TUBE]", 1)
		lngEndPos = InStr(lngStartPos, strMessage, "[/TUBE]", 1) + 7

		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 6

		'Get the original code to be replaced in the message
		strOriginalYouTube = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))

		'Replace the code codes in the message with the new formated code block
		If strBuildYouTube <> "" Then
			strMessage = Replace(strMessage, strOriginalYouTube, strBuildYouTube, 1, -1, 1)
		Else
			strMessage = Replace(strMessage, strOriginalYouTube, Replace(strOriginalYouTube, "[", "&#91;", 1, -1, 1), 1, -1, 1)
		End If
	Loop

	'Return the function
	formatYouTube = strMessage

End Function







'******************************************
'***   	   Format Hide Block		***
'******************************************

'This function formats the code blocks
Function formatHide(ByVal strMessage)


	'Declare variables
	Dim strHideMessage		'Hold the coded message
	Dim lngStartPos			'Holds search start postions
	Dim lngEndPos			'Holds end start postions
	Dim strBuildHideMessage		'Holds the built coded message
	Dim strOriginalHiddenMessage	'Holds the code block in original format
	Dim blnViewHideMessage		'Set to true if the member can view the hidden message
	
	blnViewHideMessage = False

	'Loop through all the codes in the message and convert them to formated hide block
	Do While InStr(1, strMessage, "[HIDE]", 1) > 0 AND InStr(1, strMessage, "[/HIDE]", 1) > 0
		
		'Check to see if the members has replied in this topic
		If lngLoggedInUserID <> 2 AND blnModerator = false AND blnAdmin = false Then
			
			'SQL to get if the user has replied in this topic
			strSQL = "SELECT " & strDbTable & "Thread.Thread_ID, " & strDbTable & "Thread.Author_ID " & _
			"FROM " & strDbTable & "Thread" & strDBNoLock & " " & _
			"WHERE " & strDbTable & "Thread.Topic_ID = " & lngTopicID & " " & _
				"AND " & strDbTable & "Thread.Author_ID = " & lngLoggedInUserID & ";"
				
			'Set error trapping
			On Error Resume Next
					
			'Query the database
			rsCommon.Open strSQL, adoCon
				
			'If an error has occurred write an error to the page
			If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "BBHide_format", "functions_format_post.asp")
							
			'Disable error trapping
			On Error goto 0
			
			'If a record is returned then member can view this opst
			If NOT rsCommon.EOF Then blnViewHideMessage = True
				
			'Close recordset
			rsCommon.close
		End If
		
		'If an admin or moderator allow them to view the hidden message
		If blnModerator OR blnAdmin Then blnViewHideMessage = True
		
		
	
		'Get the start and end in the message of the author who is being coded
		lngStartPos = InStr(1, strMessage, "[HIDE]", 1) + 6
		lngEndPos = InStr(lngStartPos, strMessage, "[/HIDE]", 1)

		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 6

		'If there is something returned get message to display
		If lngEndPos > lngStartPos Then

			'Get the message to be display
			strHideMessage = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))
			
			'Build the HTML for the displaying of the hidden message
			strBuildHideMessage = vbCrLf &  strHideMessage
		End If


		'Get the start and end position in the start and end position in the message of the hide block
		lngStartPos = InStr(1, strMessage, "[HIDE]", 1)
		lngEndPos = InStr(lngStartPos, strMessage, "[/HIDE]", 1) + 7

		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 6

		'Get the original code to be replaced in the message
		strOriginalHiddenMessage = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))
		
		'If the user can not view the hidden message display a message telling them so
		If blnViewHideMessage = False Then strBuildHideMessage = "<br /><em>" & strTxtYouMustBeARegisteredMemberAndPostAReplyToViewMessage & "!</em><br v/>"

		'Replace the code codes in the message with the new formated code block
		If strBuildHideMessage <> "" Then
			strMessage = Replace(strMessage, strOriginalHiddenMessage, strBuildHideMessage, 1, -1, 1)
		Else
			strMessage = Replace(strMessage, strOriginalHiddenMessage, Replace(strOriginalHiddenMessage, "[", "&#91;", 1, -1, 1), 1, -1, 1)
		End If
	Loop

	'Return the function
	formatHide = strMessage

End Function






'******************************************
'***        Display edit author		***
'******************************************

'This function formats XML into the name of the author and edit date and time if a message has been edited
'XML is used so that the date can be stored as a double npresion number so that it can display the local edit time to the message reader
Function editedXMLParser(ByVal strMessage)

		'Declare variables
		Dim strEditedAuthor 	'Holds the name of the author who is editing the post
		Dim dtmEditedDate   	'Holds the date the post was edited
		Dim lngStartPos		'Holds search start postions
		Dim lngEndPos		'Holds end start postions


		'Get the start and end in the message of the author who edit the post
		lngStartPos = InStr(1, strMessage, "<editID>", 1) + 8
		lngEndPos = InStr(1, strMessage, "</editID>", 1)
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos
		

		'If there is something returned get the authors name
		strEditedAuthor = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))

		'Get the start and end in the message of the date the message was edited
		lngStartPos = InStr(1, strMessage, "<editDate>", 1) + 10
		lngEndPos = InStr(1, strMessage, "</editDate>", 1)
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos

		'If there is something returned get the date the message was edited
		dtmEditedDate = Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos))


		'Get the start and end position in the string of the XML to remove
		lngStartPos = InStr(1, strMessage, "<edited>", 1)
		lngEndPos = InStr(1, strMessage, "</edited>", 1) + 9
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos

		'If there is something returned strip the XML from the message
		strMessage = Replace(strMessage, Trim(Mid(strMessage, lngStartPos, lngEndPos-lngStartPos)), "", 1, -1, 1)


		'Place the name of the person who edited the post
		If strEditedAuthor <> "" Then
			'If there is a date and time display the author date and time the post was edited
			If IsDate(dtmEditedDate) Then
				dtmEditedDate = CDate(dtmEditedDate)
				If blnMobileBrowser Then
					editedXMLParser = strMessage & "<span style=""font-size:16px""><br /><br />" & strTxtEditBy & " " & strEditedAuthor & " - " & DateFormat(dtmEditedDate) & " " & strTxtAt & " " & TimeFormat(dtmEditedDate) & "</span>"
				Else	
					editedXMLParser = strMessage & "<span style=""font-size:10px""><br /><br />" & strTxtEditBy & " " & strEditedAuthor & " - " & DateFormat(dtmEditedDate) & " " & strTxtAt & " " & TimeFormat(dtmEditedDate) & "</span>"
				End If	
					
			'Just display the author name who edited the post
			Else
				If blnMobileBrowser Then
					editedXMLParser = strMessage & "<span style=""font-size:16px""><br /><br />" & strTxtEditBy & " " & strEditedAuthor & "</span>"
				Else	
					editedXMLParser = strMessage & "<span style=""font-size:10px""><br /><br />" & strTxtEditBy & " " & strEditedAuthor & "</span>"
				End If	
				
				
			End If
		End If
End Function





'******************************************
'***    Convert Post to Text Function	***
'******************************************

'Function to romove icons and colurs to just leave plain text
Function ConvertToText(ByVal strMessage)

	Dim strTempMessage	'Temporary word hold for e-mail and url words
	Dim strMessageLink	'Holds the new mesage link that needs converting back into code
	Dim lngStartPos		'Holds the start position for a link
	Dim lngEndPos		'Holds the end position for a word
	Dim intLoop		'Loop counter

	'Remove hyperlinks
	strMessage = Replace(strMessage, " target=""_blank""", "", 1, -1, 1)
	
	
	'Loop through the message till all or any hyperlinks are turned into back into froum codes
	Do While InStr(1, strMessage, "<a href=""", 1) > 0 AND InStr(1, strMessage, "</a>", 1) > 0
						    	
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
		If InStr(1, strTempMessage, "src=""", 1) Then
			strTempMessage = Replace(strTempMessage, "<a href=""", " ", 1, -1, 1)
			strTempMessage = Replace(strTempMessage, "</a>", " ", 1, -1, 1)
		Else
			strTempMessage = Replace(strTempMessage, "<a href=""", " <font color='#0000FF'>", 1, -1, 1)
			strTempMessage = Replace(strTempMessage, "</a>", " ", 1, -1, 1)
			strTempMessage = Replace(strTempMessage, """>", "</font> - ", 1, -1, 1)
		End If
		
		'Place the new fromatted codes into the message string body
		strMessage = Replace(strMessage, strMessageLink, strTempMessage, 1, -1, 1)		
	Loop
	
	'Get any that may slip through (don't look as good but still has the same effect)
	strMessage = Replace(strMessage, "<a href= """, "", 1, -1, 1)
	strMessage = Replace(strMessage, "<a href='", "", 1, -1, 1)
	strMessage = Replace(strMessage, "</a>", "", 1, -1, 1)

	'Return the message with the icons and text colours removed
	ConvertToText = strMessage

End Function





'******************************************
'***     Search Word Highlighter	***
'******************************************

'Function to highlight search words if coming from search page
Private Function searchHighlighter(ByVal strMessage, ByVal sarySearchWord)

	Dim intHighlightLoopCounter	'Loop counter to loop through words and hightlight them
	Dim strTempMessage		'Temporary message store
	Dim lngMessagePosition		'Holds the message position
	Dim intHTMLTagLength		'Holds the length of the HTML tags
	Dim intSearchWordLength		'Holds the length of teh search word
	Dim blnTempUpdate		'Set to true if the temp message variable is updated


	'Loop through each character in the post message
	For lngMessagePosition = 1 to Len(strMessage)

		'Initilise for each pass
		blnTempUpdate = False
		
		'If an HTML tag is found then move to the end of it so that no words in the HTML are highlighted
		If Mid(strMessage, lngMessagePosition, 1) = "<" Then
			

			'Get the length of the HTML tag
			intHTMLTagLength = (InStr(lngMessagePosition, strMessage, ">", 1) - lngMessagePosition)
			
			'Prevent errors
			If intHTMLTagLength < 0 Then intHTMLTagLength = lngMessagePosition

			'Place the HTML tag back into the tempary message store
			strTempMessage = strTempMessage & Mid(strMessage, lngMessagePosition, intHTMLTagLength)

			'Add the length of the HTML tag to the post message position variable
			lngMessagePosition = lngMessagePosition + intHTMLTagLength
		
		End If

		'Loop through the search words to see if they are in the message post
		For intHighlightLoopCounter = 0 to UBound(sarySearchWord)

			'If there is a search word in the array position check it
			If sarySearchWord(intHighlightLoopCounter) <> "" Then

				'Get the length of the search word
				intSearchWordLength = Len(sarySearchWord(intHighlightLoopCounter))

				'If the next XX characters are the same as the search word then highlight them
				If LCase(Mid(strMessage, lngMessagePosition, intSearchWordLength)) = LCase(sarySearchWord(intHighlightLoopCounter)) Then

					'Highlight the search word
					strTempMessage = strTempMessage & "<span class=""highlight"">" & Mid(strMessage, lngMessagePosition, intSearchWordLength) & "</span>"

					'Add the length of the replaced search word to the post message position variable
					lngMessagePosition = lngMessagePosition + intSearchWordLength - 1

					'Set the changed boolean to true
					blnTempUpdate = True
				End If
			End If
		Next

		'If a search word is not highlighted then add the character from the post message being checked to the temp variable
		If blnTempUpdate = False Then
			strTempMessage = strTempMessage & Mid(strMessage, lngMessagePosition, 1)
		End If
	Next

	'Return the function
	searchHighlighter = strTempMessage
End Function




'**********************************************
'***  convert text links to href links   ****
'**********************************************

Function InsertHyperlinks(ByVal strMessage)

	Dim objRegExp
	Dim objMatches, objMatch
	Dim lngStartPos, lngEndPos
	Dim strURL
	Dim strTempMessage
	

	'Initilise variables
  	strTempMessage = ""
  	lngStartPos = 1
  	lngEndPos = 1
  	
  	'Create object
  	Set objRegExp = New RegExp

	
	'Match URL's
	objRegExp.Pattern = "((www\.|(http|https|ftp|news|file)+\:\/\/)[&#95;.a-z0-9-]+\.[a-z0-9\/&#95;:@=.+?_,##%&~-]*[^.|\'|\# |!|\(|?|,| |>|<|;|\)])"
	objRegExp.IgnoreCase = True      
	objRegExp.Global = True 
	 
	'Excure regular expression
	Set objMatches = objRegExp.Execute(strMessage)
	  
	'loop through each text link found
	For Each objMatch in objMatches
	
	    	
	    	'Get end postion
	    	lngEndPos = objMatch.FirstIndex
	    	
	    	'Place in to temp message string
	    	strTempMessage = strTempMessage & Mid(strMessage, lngStartPos, lngEndPos - lngStartPos + 1)
	
	    	'Call Href function to build hyperlink
	    	strTempMessage = strTempMessage & GetHref(objMatch.Value)
	
	    	'Get the start position
	    	lngStartPos = lngEndPos + objMatch.Length + 1
	  Next
	  
	  'Replace the text URL in the string with the hyperlink
	  strMessage = strTempMessage & Mid(strMessage, lngStartPos)
	  

	  'Retun function
	  InsertHyperlinks = strMessage
End Function

'Used with InsertHyperlinks function above to build href hyperlink
Function GetHref(strURL)
	Dim strTempMessage

	strTempMessage = "<a href="""
	 
	If LCase(Left(strURL, 3)) = "www" Then
		strTempMessage = "<a href=""http://" & strURL & """>" & strURL & "</a>"
	Else
		strTempMessage = "<a href=""" & strURL & """>" & strURL & "</a>"
	End If
	GetHref = strTempMessage

End Function

%>