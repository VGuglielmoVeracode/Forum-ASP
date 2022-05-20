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


Dim strPostPage



strPostPage = "new_post.asp"


'Turn off sintaures if disabled for group
If blnGroupSignatures = False Then blnSignatures = False


Response.Write("<br />" & _
vbCrLf & "<form method=""post"" name=""frmMessageForm"" id=""frmMessageForm"" action=""" & strPostPage & strQsSID1 & """ onSubmit=""return CheckForm();"" onReset=""return clearForm();"">" & _
vbCrLf & "<table cellspacing=""1"" cellpadding=""10"" class=""tableBorder"" align=""center"">" & _
vbCrLf & " <tr class=""tableLedger"">" & _
vbCrLf & "  <td>" & strTxtQuickReply & "</td>" & _
vbCrLf & " </tr>" & _
vbCrLf & " <tr class=""tableRow"">" & _
vbCrLf & "  <td>" & _
vbCrLf & "   <table cellspacing=""0"" cellpadding=""2"" align=""center"">")

        
'If the poster is in a guest then get them to enter a name
If lngLoggedInUserID = 2 AND (strMode <> "edit" AND strMode <> "editTopic" AND strMode <> "editPoll") Then
	   
	Response.Write(vbCrLf & "    <tr>" & _
	vbCrLf & "     <td align=""right"" width=""15%"">" & strTxtName & ":</td>" & _
	vbCrLf & "     <td width=""85%"">" & _
	vbCrLf & "      <input type=""text"" name=""Gname"" id=""Gname"" size=""20"" maxlength=""20"" tabindex=""1"" />" & _
	vbCrLf & "     </td>" & _
	vbCrLf & "    </tr>")
End If



'The message textarea
Response.Write(vbCrLf & "    <tr class=""tableRow"">" & _
vbCrLf & "     <td valign=""top"" align=""right"" width=""10%""><br />" & strTxtMessage & ":" & _
vbCrLf & "     <td width=""90%"" valign=""top"">")


'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Rich Text Editor(TM) " & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******



'Name of the HTML form the textarea is within
Const strFormName = "frmMessageForm"
	
'ID tag name of HTML textarea being replaced
Const strTextAreaName = "message"


'Load default CSS and Javascript
Response.Write(vbCrLf & vbCrLf & "     <script language=""JavaScript"" src=""RTE_javascript_common.asp"" type=""text/javascript""></script>")


'Javascript function to switch to full reply screen
Response.Write(vbCrLf & "     <script language=""JavaScript"">" & _
vbCrLf & "     function FullReply(formName){")

'If RTE enabled get the message content from the iframe
If RTEenabled() <> "false" AND blnRTEEditor AND blnWYSIWYGEditor Then
	Response.Write(vbCrLf & "     	document.getElementById('message').value=document.getElementById('WebWizRTE').contentWindow.document.body.innerHTML;")
Else
	Response.Write(vbCrLf & "     	document.getElementById('message').value=document.getElementById('message').value;")
End If


Response.Write(vbCrLf & "     	formName.action = 'new_reply_form.asp?TID=" & lngTopicID & "&PN=" & lngTotalRecordsPages & "&TR=" &  lngTotalRecords & strQsSID2 & "';" & _
vbCrLf & "     	formName.submit();" & _
vbCrLf & "     }" & _
vbCrLf & "     </script>")


'If this is an RTE enabled web browser load in the RTE content
If RTEenabled() <> "false" AND blnRTEEditor AND blnWYSIWYGEditor Then
	
	'Load in Javascript for RTE browsers
	Response.Write(vbCrLf & "     <script language=""JavaScript"" src=""RTE_javascript.asp?textArea=" & Server.URLEncode(strTextAreaName) & "&M=" & strMode & "&QR=true" & strQsSID2 & """ type=""text/javascript""></script>")

	'Create IFrame with message for those with javascript disabled
	Response.Write(vbCrLf & "     <script language=""javascript"">" & _
	vbCrLf & "      WebWizRTEtoolbar('" & strFormName & "');" & _
	vbCrLf & "      document.write ('<iframe id=""WebWizRTE"" src=""RTE_textarea.asp?mode=reply&ID=" & CInt(RND * 2000) & strQsSID2 & """ width=""650"" height=""150"" style=""border: #A5ACB2 1px solid"" onLoad=""initialiseWebWizRTE();"" tabindex=""2""></iframe>');" & _
	vbCrLf & "     </script>" & _
	vbCrLf & "     <noscript><strong><br /><br />" & strTxtJavaScriptEnabled & "</strong></noscript>")

	'Create hidden inputs
	Response.Write(vbCrLf & "     <input type=""hidden"" name=""message"" id=""message"" value="""" />")
	Response.Write(vbCrLf & "     <input type=""hidden"" name=""browser"" id=""browser"" value=""RTE"" />")

'If this is not an RTE enabled web browser load in the NON-RTE content
Else
	
	'Load in Javascript for non-RTE browsers
	Response.Write(vbCrLf & "     <script language=""JavaScript"" src=""non_RTE_javascript.asp?textArea=" & Server.URLEncode(strTextAreaName) & "&QR=true" & strQsSID2 & """ type=""text/javascript""></script>")

	'Create Toolbar and Iframes
	Response.Write(vbCrLf & "     <script language=""JavaScript"">WebWizRTEtoolbar('" & strFormName & "');</script>")

	'Create the text area
	Response.Write(vbCrLf & "     <textarea name=""message"" id=""message"" rows=""7"" wrap=""virtual"" tabindex=""2"" style=""width:")
	If blnMobileBrowser Then Response.Write("95%") Else Response.Write("596px")
	Response.Write(";"" />")
	Response.Write("</textarea>")
	
End If

'********************************************************************

Response.Write(vbCrLf & "     </td>" & _
vbCrLf & "    </tr>" & _
vbCrLf & "    <tr class=""tableRow"">" & _
vbCrLf & "     <td align=""right"">&nbsp;</td>" & _
vbCrLf & "     <td valign=""bottom"" class=""text"">")
	
'If rel=nofollow the display a message
If blnNoFollowTagInLinks AND blnMobileBrowser = False Then Response.Write("   	&nbsp;<span class=""smText"">" & strTxtNoFollowAppliedToAllLinks & "</span><br />")

Response.Write(vbCrLf & "     &nbsp;<input type=""checkbox"" name=""forumCodes"" id=""forumCodes"" value=""True"" checked tabindex=""3"" />" & strTxtEnable & " <a href=""javascript:winOpener('BBcodes.asp','codes',1,1,610,500)"">" & strTxtForumCodes & "</a> " & _
vbCrLf & "     </td></tr>")


'If signature of e-mail notify then display row to show
If (blnLoggedInUserEmail AND blnEmail) OR (blnLoggedInUserSignature AND ((blnSignatures AND strMode = "reply") )) Then

	Response.Write(vbCrLf & "    <tr class=""tableRow"">" & _
	vbCrLf & "    <td align=""right"" height=""7"">&nbsp;</td>" & _
	vbCrLf & "    <td height=""7"" valign=""bottom"" class=""text"">")

	'If the user has a signature offer them the chance to show it
	If blnLoggedInUserSignature AND blnSignatures Then

		Response.Write(vbCrLf & "      &nbsp;<input type=""checkbox"" name=""signature"" value=""True""")
		If blnAttachSignature = True Then Response.Write(" checked")
		Response.Write(" tabindex=""4"" />" & strTxtShowSignature & "&nbsp;")

	End If

	'Display e-mail notify of replies option
	If blnEmail AND blnLoggedInUserEmail Then

		Response.Write(vbCrLf & "     &nbsp;<input type=""checkbox"" name=""email"" value=""True""")
		If blnReplyNotify Then Response.Write(" checked")
		Response.Write(" tabindex=""5"" />" & strTxtEmailNotify & "&nbsp;")

	End If

		
Response.Write(vbCrLf & "    </td>" & _
vbCrLf & "    </tr>")


End If



'Display CAPTCHA images for Guest posting
If blnCAPTCHAsecurityImages AND lngLoggedInUserID = 2 Then 
	
	%>
 <tr>
  <td align="right" valign="top"><% = strTxtUniqueSecurityCode %>:</td>
  <td><!--#include file="CAPTCHA_form_inc.asp" --><span class="smText"><% = strTxtEnterCAPTCHAcode %></span></td>
 </tr><%


End If



Response.Write(vbCrLf & "     <tr class=""tableRow"">" & _
vbCrLf & "     <td>")

	'Ajax session ping (60,000 = 1 minute)
         Response.Write(vbCrLf & "      <input type=""hidden"" name=""mode"" id=""mode"" value=""" & strMode & """ />" & _
	 vbCrLf & "      <input type=""hidden"" name=""FID"" id=""FID"" value=""" & intForumID & """ />" & _
	 vbCrLf & "      <input type=""hidden"" name=""TID"" id=""TID"" value=""" & lngTopicID & """ />" & _
	 vbCrLf & "      <input type=""hidden"" name=""PN"" id=""PN"" value=""" & intRecordPositionPageNum & """ />" & _
         vbCrLf & "      <input type=""hidden"" name=""ThreadPos"" id=""ThreadPos"" value=""" & (lngTotalRecords + 1) & """ />"  & _
         vbCrLf & "      <input type=""hidden"" name=""QR"" id=""QR"" value=""1"" />" & _
	 vbCrLf & "      <input type=""hidden"" name=""pre"" id=""pre"" value="""" />"  & _
	 vbCrLf & "      <input type=""hidden"" name=""formID"" id=""formID"" value=""" & strFormID & """ />"  & _
         vbCrLf & "      <div id=""ajaxFormFields""><input type=""hidden"" name=""session"" id=""session"" value=""false"" /></div>"  & _
         vbCrLf & "      <script  language=""JavaScript"">getFormData(); function getFormData(){getAjaxData('ajax_session_alive.asp?XID=" & strFormID & strQsSID3 & "','ajaxFormFields'); setTimeout('getFormData()',120000);}</script>"  & _
         vbCrLf & "     </td>" & _
         vbCrLf & "     <td width=""70%"">")
         


Response.Write(vbCrLf & "       <input type=""submit"" id=""Submit"" name=""Submit"" value=""" & strTxtPostReply & """")





'If RTE enabled then use javascript to submit the RTE data
If RTEenabled() <> "false" AND blnRTEEditor AND blnWYSIWYGEditor Then
	Response.Write(" onclick=""document.getElementById('message').value=document.getElementById('WebWizRTE').contentWindow.document.body.innerHTML;""")
End If
Response.Write(" tabindex=""6"" />")  
   
                     
Response.Write(vbCrLf & "       <input type=""button"" name=""Preview"" value=""" & strTxtPreview & """ onclick=""")
'If RTE enabled get the preview content from the iframe
If RTEenabled() <> "false" AND blnRTEEditor AND blnWYSIWYGEditor Then
	Response.Write("document.getElementById('pre').value=document.getElementById('WebWizRTE').contentWindow.document.body.innerHTML;")
Else
	Response.Write("document.getElementById('pre').value=document.getElementById('message').value;")
End If
Response.Write(" OpenPreviewWindow(document.frmMessageForm);"" tabindex=""7"" />" & _

vbCrLf & "       <input type=""reset"" name=""Reset"" id=""Reset"" value=""" & strTxtClearForm & """ tabindex=""8"" />" & _
vbCrLf & "      </td>" & _
vbCrLf & "     </tr>" & _
vbCrLf & "    </table>" & _
vbCrLf & "   </td>" & _
vbCrLf & "  </tr>" & _
vbCrLf & " </table>" & _
vbCrLf & "</form>")
%>