<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  Web Wiz Rich Text Editor(TM)
'**  http://www.richtexteditor.org
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





'***********************************
'****	non-RTE Toolbar 1	****
'***********************************


'Font type toolbar
'------------
If blnFontStyle OR blnFontType OR blnFontSize OR blnTextColour Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
'Font Type
If blnFontType  Then Response.Write(" <select name=""selectFont"" onchange=""FontCode(selectFont.options[selectFont.selectedIndex].value, \'FONT\');selectFont.options[0].selected = true;""><option selected>" & strTxtFontTypes & "</option><option value=\'FONT=Arial\'>Arial</option><option value=\'FONT=Courier\'>Courier New</option><option value=\'FONT=Times\'>Times</option><option value=\'FONT=Verdana\'>Verdana</option></select>")
'Font Size
If blnFontSize  Then Response.Write(" <select name=""selectSize"" onchange=""FontCode(selectSize.options[selectSize.selectedIndex].value, \'SIZE\');selectSize.options[0].selected = true;""><option selected>" & strTxtFontSizes & "</option><option value=\'SIZE=1\'>1</option><option value=\'SIZE=2\'>2</option><option value=\'SIZE=3\'>3</option><option value=\'SIZE=4\'>4</option><option value=\'SIZE=5\'>5</option><option value=\'SIZE=6\'>6</option></select>")
'Font colour
If blnTextColour Then Response.Write(" <select name=""selectColour"" onchange=""FontCode(selectColour.options[selectColour.selectedIndex].value, \'COLOR\');selectColour.options[0].selected = true;""><option value=""0"" selected>" & strTxtFontColour & "</option><option value=\'COLOR=black\'>" & strTxtBlack & "</option><option value=\'COLOR=white\'>" & strTxtWhite & "</option><option value=\'COLOR=blue\'>" & strTxtBlue & "</option><option value=\'COLOR=red\'>" & strTxtRed & "</option><option value=\'COLOR=green\'>" & strTxtGreen & "</option><option value=\'COLOR=yellow\'>" & strTxtYellow & "</option><option value=\'COLOR=orange\'>" & strTxtOrange & "</option><option value=\'COLOR=brown\'>" & strTxtBrown & "</option><option value=\'COLOR=magenta\'>" & strTxtMagenta & "</option><option value=\'COLOR=cyan\'>" & strTxtCyan & "</option><option value=\'COLOR=lime green\'>" & strTxtLimeGreen & "</option></select>")
If blnFontStyle OR blnFontType OR blnFontSize OR blnTextColour Then Response.Write("');")


Response.Write(vbCrLf & "	document.writeln('<br />');")



'***********************************
'****	non-RTE Toolbar 2	****
'***********************************

		
'Toolbar buttons

'Preview, print, spell toolbar
'------------
If blnPreview Then 
	Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
	'Button Pop up Preview
	Response.Write("<img src=""" & strImagePath & "post_button_preview.gif"" title=""" & strTxtPreview & """ onclick=""document.getElementById(\'pre\').value=document.getElementById(textArea).value; OpenPreviewWindow(document.' + formName + ');"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
	Response.Write("');")
End If

'Font style toolbar
'------------
If blnBold OR blnItalic OR blnUnderline Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
If blnBold Then Response.Write("<img src=""" & strImagePath & "post_button_bold.gif"" title=""" & strTxtBold & """ onclick=""AddMessageCode(\'B\',\'" & strTxtEnterBoldText & "\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnItalic Then Response.Write("<img src=""" & strImagePath & "post_button_italic.gif""  title=""" & strTxtItalic & """ onclick=""AddMessageCode(\'I\',\'" & strTxtEnterItalicText & "\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnUnderline Then Response.Write("<img src=""" & strImagePath & "post_button_underline.gif"" title=""" & strTxtUnderline & """ onclick=""AddMessageCode(\'U\',\'" & strTxtEnterUnderlineText & "\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")            
If blnBold OR blnItalic OR blnUnderline Then Response.Write("');")


'insert toolbar
'------------
If blnAdvAdddHyperlink OR blnAddHyperlink Or blnAttachments OR blnAdvAddImage OR blnAddImage OR blnImageUpload Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
'If hyperlink is enabled
If blnAdvAdddHyperlink OR blnAddHyperlink Then 
	Response.Write("<img src=""" & strImagePath & "post_button_hyperlink.gif"" onclick=""AddCode(\'URL\')"" title=""" & strTxtAddHyperlink & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
	Response.Write("<img src=""" & strImagePath & "post_button_email.gif"" onclick=""AddCode(\'EMAIL\')"" title=""" & strTxtAddEmailLink & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
End If
If blnAttachments Then Response.Write("<img src=""" & strImagePath & "post_button_file_upload.gif"" onclick=""winOpener(\'non_RTE_upload_files.asp?textArea=\'+textArea+\'" & strQsSID2 & "\',\'files\',0,1,400,163)"" title=""" & strTxtFileUpload & """ border=""0"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
'images
If blnAdvAddImage OR blnAddImage Then Response.Write("<img src=""" & strImagePath & "post_button_image.gif"" onclick=""AddCode(\'IMG\')"" title=""" & strTxtAddImage & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
'If image uploading is allowed have an image upload button
If blnImageUpload Then Response.Write("<img src=""" & strImagePath & "post_button_image_upload.gif"" onclick=""winOpener(\'non_RTE_upload_images.asp?textArea=\'+textArea+\'" & strQsSID2 & "\',\'images\',0,1,400,150)"" title=""" & strTxtImageUpload & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnAdvAdddHyperlink OR blnAddHyperlink Or blnAttachments OR blnAdvAddImage OR blnAddImage OR blnImageUpload Then Response.Write("');")


'List type and indent toolbar
'------------
If blnOrderList OR blnUnOrderList OR blnIndent Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
If blnOrderList OR blnUnOrderList Then Response.Write("<img src=""" & strImagePath & "post_button_or_list.gif"" onclick=""AddCode(\'LIST\', \'\')"" title=""" & strTxtList & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnIndent Then Response.Write("<img src=""" & strImagePath & "post_button_indent.gif"" onclick=""AddCode(\'INDENT\', \'\')"" title=""" & strTxtIndent & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />") 
If blnOrderList OR blnUnOrderList OR blnOutdent OR blnIndent Then Response.Write("');")


'Font block format toolbar
'------------
If blnCentre Then 
	Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
	Response.Write("<img src=""" & strImagePath & "post_button_centre.gif"" onclick=""AddMessageCode(\'center\',\'" & strTxtEnterCentredText & "\', \'\')"" title=""" & strTxtCentrejustify & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />") 
	Response.Write("');")
End If


'About toolbar
'------------
If blnLCode Then 
	Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
	Response.Write("<img src=""" & strImagePath & "post_button_about.gif"" onclick=""winOpener(\'RTE_popup_about.asp" & strQsSID1 & "\',\'about\',0,0,420,187)"" title=""" & strTxtAboutRichTextEditor & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
	Response.Write("');")
End If

'Mode
Response.Write(vbCrLf & "	document.writeln('&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<select name=""selectMode"" onchange=""PromptMode(this)""><option value=""1"" selected>" & strTxtPrompt & " " & strTxtMode & "</option><option value=""0"">" & strTxtBasic & " " & strTxtMode & "</option></select>');")
%>