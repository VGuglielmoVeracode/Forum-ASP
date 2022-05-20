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



Dim strMode

strMode = Request.QueryString("M")



'***************************
'****	RTE Toolbar 1	****
'***************************

'File toolbar
'------------
If blnNew Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
If blnNew Then Response.Write("<img src=""" & strImagePath & "post_button_new.gif"" title=""" & strTxtNewBlankDoc & """ onclick=""clearWebWizRTE()"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnNew Then Response.Write("');")



'Preview, print, spell toolbar
'------------
If blnPrint OR blnPreview OR (blnSpellCheck AND RTEenabled = "winIE") OR blnHTMLView Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
If blnPrint Then Response.Write("<img src=""" & strImagePath & "post_button_print.gif"" title=""" & strTxtPrint & """ onclick=""printEditor()"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
'Button Pop up Preview
If blnPreview Then Response.Write("<img src=""" & strImagePath & "post_button_preview.gif"" title=""" & strTxtPreview & """ onclick=""document.getElementById(\'pre\').value = document.getElementById(\'WebWizRTE\').contentWindow.document.body.innerHTML; OpenPreviewWindow(document.' + formName + ');"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
'Spell check button
If blnSpellCheck AND RTEenabled = "winIE" Then Response.Write("<img src=""" & strImagePath & "post_button_spell_check.gif"" onclick=""checkspell()"" title=""" & strTxtstrSpellCheck & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" id=""spellboundSC"" />")
If blnHTMLView Then Response.Write("<img src=""" & strImagePath & "post_button_html.gif"" title=""" & strTxtToggleHTMLView & """ onclick=""HTMLview()"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnPrint OR blnPreview OR (blnSpellCheck AND RTEenabled = "winIE") OR blnHTMLView Then Response.Write("');")


'The span tag needs to be put in to hide the other options in HTML view
Response.Write(vbCrLf & "	document.writeln('<span id=""ToolBar1"">');")



'cut, copy, paste toolbar
'------------
If blnCut OR blnCopy OR blnPaste OR blnWordPaste Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
If blnCut Then Response.Write("<img src=""" & strImagePath & "post_button_cut.gif"" onclick=""FormatText(\'cut\')"" title=""" & strTxtCut & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnCopy Then Response.Write("<img src=""" & strImagePath & "post_button_copy.gif"" onclick=""FormatText(\'copy\')"" title=""" & strTxtCopy & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnPaste Then Response.Write("<img src=""" & strImagePath & "post_button_paste.gif"" onclick=""FormatText(\'paste\')"" title=""" & strTxtPaste & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnWordPaste Then Response.Write("<img src=""" & strImagePath & "post_button_word.gif"" onclick=""winOpener(\'RTE_popup_word_paste.asp" & strQsSID1 & "\',\'save\',0,1,600,290)"" title=""" & strTxtPasteFromWord & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnCut OR blnCopy OR blnPaste OR blnWordPaste Then Response.Write("');")



'undo redo toolbar
'------------
If blnUndo OR blnRedo Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
If blnUndo Then Response.Write("<img src=""" & strImagePath & "post_button_undo.gif"" onclick=""FormatText(\'undo\')"" title=""" & strTxtUndo & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnRedo Then Response.Write("<img src=""" & strImagePath & "post_button_redo.gif"" onclick=""FormatText(\'redo\')"" title=""" & strTxtRedo & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnUndo OR blnRedo Then Response.Write("');")



'insert toolbar
'------------
If blnAdvAdddHyperlink OR blnAddHyperlink Or blnAttachments OR blnAdvAddImage OR blnAddImage OR blnInsertTable OR blnFlashFiles OR blnYouTube OR blnSpecialCharacters Or blnEmoticons Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
'If advanced hyperlink is enabled have a popup link
If blnAdvAdddHyperlink Then
	Response.Write("<img src=""" & strImagePath & "post_button_hyperlink.gif"" onclick=""winOpener(\'RTE_popup_link.asp" & strQsSID1 & "\',\'link\',0,1,490,337)"" title=""" & strTxtAddHyperlink & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
'Else have the basic hperlink adding feature
ElseIf blnAddHyperlink Then
	Response.Write("<img src=""" & strImagePath & "post_button_hyperlink.gif"" onclick=""FormatText(\'createLink\')"" title=""" & strTxtAddHyperlink & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
End If
If blnAttachments Then Response.Write("   <img src=""" & strImagePath & "post_button_file_upload.gif"" onclick=""winOpener(\'RTE_popup_file_atch.asp" & strQsSID1 & "\',\'files\',0,1,775,404)"" title=""" & strTxtFileUpload & """ border=""0"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
'Popup window for images
If blnAdvAddImage Then
	Response.Write("<img src=""" & strImagePath & "post_button_image.gif"" onclick=""winOpener(\'RTE_popup_adv_image.asp" & strQsSID1 & "\',\'insertImg\',0,1,775,402)"" title=""" & strTxtAddImage & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
ElseIf blnAddImage Then
	Response.Write("<img src=""" & strImagePath & "post_button_image.gif"" onclick=""winOpener(\'RTE_popup_image.asp" & strQsSID1 & "\',\'insertImg\',0,1,550,360)"" title=""" & strTxtAddImage & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
End If
'Popup window for Adobe Flash and YouTube
If blnFlashFiles OR blnYouTube Then Response.Write("   <img src=""" & strImagePath & "post_button_movie.gif"" onclick=""winOpener(\'RTE_popup_movie.asp" & strQsSID1 & "\',\'movie\',0,1,740,505)"" title=""" & strTxtInsertMovie & """ border=""0"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
'Popup window for tables
If blnInsertTable Then Response.Write("<img src=""" & strImagePath & "post_insert_table.gif"" onclick=""winOpener(\'RTE_popup_table.asp" & strQsSID1 & "\',\'insertTable\',0,1,400,198)"" title=""" & strTxtInsertTable & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
'Popup window for Special Characters
If blnSpecialCharacters Then Response.Write("<img src=""" & strImagePath & "post_button_sp_char.gif"" onclick=""winOpener(\'RTE_popup_special_characters.asp" & strQsSID1 & "\',\'insertTable\',0,1,550,304)"" title=""" & strTxtSpecialCharacters & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
'Button Pop up for emoticons
If blnEmoticons Then Response.Write("<img src=""" & strImagePath & "post_button_smiley.gif"" onclick=""winOpener(\'RTE_popup_emoticons.asp" & strQsSID1 & "\',\'emot\',0,1,650,350)"" title=""" & strTxtEmoticons & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnAdvAdddHyperlink OR blnAddHyperlink Or blnAttachments OR blnAdvAddImage OR blnAddImage OR blnInsertTable OR blnFlashFiles OR blnYouTube OR blnSpecialCharacters Or blnEmoticons Then Response.Write("');")



'List type and indent toolbar
'------------
If blnOrderList OR blnUnOrderList OR blnOutdent OR blnIndent Or blnHorizontalRule Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
If blnOrderList Then Response.Write("<img src=""" & strImagePath & "post_button_or_list.gif"" onclick=""FormatText(\'InsertOrderedList\', \'\')"" title=""" & strTxtstrTxtOrderedList & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnUnOrderList Then Response.Write("<img src=""" & strImagePath & "post_button_list.gif"" onclick=""FormatText(\'InsertUnorderedList\', \'\')"" title=""" & strTxtUnorderedList & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnOutdent Then Response.Write("<img src=""" & strImagePath & "post_button_outdent.gif"" onclick=""FormatText(\'Outdent\', \'\')"" title=""" & strTxtOutdent & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnIndent Then Response.Write("<img src=""" & strImagePath & "post_button_indent.gif"" onclick=""FormatText(\'indent\', \'\')"" title=""" & strTxtIndent & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnHorizontalRule Then Response.Write("<img src=""" & strImagePath & "post_button_h_rule.gif"" title=""" & strTxtHorizontalRule & """ onclick=""FormatText(\'inserthorizontalrule\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnOrderList OR blnUnOrderList OR blnOutdent OR blnIndent Or blnHorizontalRule Then Response.Write("');")


Response.Write(vbCrLf & "	document.writeln('</span><br />');")




'***************************
'****	RTE Toolbar 2	****
'***************************

Response.Write(vbCrLf & "	document.writeln('<span id=""ToolBar2"">');")


'Toolbar buttons


'Font type toolbar
'------------
If blnFontStyle OR blnFontType OR blnFontSize Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
'Font Style
If blnFontStyle Then Response.Write("<img id=""formatblock"" src=""" & strImagePath & "post_button_format.gif"" title=""" & strTxtFontStyle & """ onclick=""FormatText(\'formatblock\', \'\')"" class=""WebWizRTEbutton"" />")
'Font Type
If blnFontType  Then Response.Write("<img id=""fontname"" src=""" & strImagePath & "post_button_font.gif"" title=""" & strTxtFontTypes & """ onclick=""FormatText(\'fontname\', \'\')"" class=""WebWizRTEbutton"" />")
'Font Size
If blnFontSize  Then Response.Write("<img id=""fontsize"" src=""" & strImagePath & "post_button_size.gif"" title=""" & strTxtFontSizes & """ onclick=""FormatText(\'fontsize\', \'\')"" class=""WebWizRTEbutton"" />")
If blnFontStyle OR blnFontType OR blnFontSize Then Response.Write("');")



'Font style toolbar
'------------
If blnBold OR blnItalic OR blnUnderline Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
If blnBold Then Response.Write("<img src=""" & strImagePath & "post_button_bold.gif"" title=""" & strTxtBold & """ onclick=""FormatText(\'bold\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnItalic Then Response.Write("<img src=""" & strImagePath & "post_button_italic.gif""  title=""" & strTxtItalic & """ onclick=""FormatText(\'italic\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnUnderline Then Response.Write("<img src=""" & strImagePath & "post_button_underline.gif"" title=""" & strTxtUnderline & """ onclick=""FormatText(\'underline\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnBold OR blnItalic OR blnUnderline Then Response.Write("');")



'Stikethrough, super/sub script toolbar
'------------
If blnStrikeThrough OR blnSubscript OR blnSuperscript Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
If blnStrikeThrough Then Response.Write("<img src=""" & strImagePath & "post_button_strike.gif"" title=""" & strTxtStrikeThrough & """ onclick=""FormatText(\'strikethrough\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnSubscript Then Response.Write("<img src=""" & strImagePath & "post_button_subscript.gif"" title=""" & strTxtSubscript & """ onclick=""FormatText(\'subscript\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnSuperscript Then Response.Write("<img src=""" & strImagePath & "post_button_superscript.gif"" title=""" & strTxtSuperscript & """ onclick=""FormatText(\'superscript\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnStrikeThrough OR blnSubscript OR blnSuperscript Then Response.Write("');")



'Font block format toolbar
'------------
If blnLeftJustify OR blnCentre OR blnRightJustify OR blnFullJustify Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
If blnLeftJustify Then Response.Write("<img src=""" & strImagePath & "post_button_left_just.gif"" onclick=""FormatText(\'justifyleft\', \'\')"" title=""" & strTxtLeftJustify & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnCentre Then Response.Write("<img src=""" & strImagePath & "post_button_centre.gif"" onclick=""FormatText(\'justifycenter\', \'\')"" title=""" & strTxtCentrejustify & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnRightJustify Then Response.Write("<img src=""" & strImagePath & "post_button_right_just.gif"" onclick=""FormatText(\'justifyright\', \'\')"" title=""" & strTxtRightJustify & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnFullJustify Then Response.Write("<img src=""" & strImagePath & "post_button_justify.gif"" onclick=""FormatText(\'justifyfull\', \'\')"" title=""" & strTxtJustify & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnLeftJustify OR blnCentre OR blnRightJustify OR blnFullJustify Then Response.Write("');")



'Text colour toolbar
'------------
If blnTextColour OR blnTextBackgroundColour Then Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
If blnTextColour Then Response.Write("<img id=""forecolor"" src=""" & strImagePath & "post_button_colour_pallete.gif"" title=""" & strTxtTextColour & """ onclick=""FormatText(\'forecolor\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If RTEenabled = "winIE" AND blnTextBackgroundColour Then Response.Write("<img id=""backcolor"" src=""" & strImagePath & "post_button_fill.gif"" title=""" & strTxtBackgroundColour & """ onclick=""FormatText(\'backcolor\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If RTEenabled = "Gecko" AND blnTextBackgroundColour Then Response.Write("<img id=""hilitecolor"" src=""" & strImagePath & "post_button_fill.gif"" title=""" & strTxtBackgroundColour & """ onclick=""FormatText(\'hilitecolor\', \'\')"" class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
If blnTextColour OR blnTextBackgroundColour Then Response.Write("');")

Response.Write(vbCrLf & "	document.writeln('</span>');")


'About toolbar
'------------
If blnLCode Then
	Response.Write(vbCrLf & "	document.writeln('<img src=""" & strImagePath & "toolbar_separator.gif"" />")
	Response.Write("<img src=""" & strImagePath & "post_button_about.gif"" onclick=""winOpener(\'RTE_popup_about.asp" & strQsSID1 & "\',\'about\',0,0,420,177)"" title=""" & strTxtAboutRichTextEditor & """ class=""WebWizRTEbutton"" onmouseover=""overIcon(this)"" onmouseout=""outIcon(this)"" />")
	Response.Write("');")
End If
%>