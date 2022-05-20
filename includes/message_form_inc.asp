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





'If a private message go to pm post message page otherwise goto post message page
If strMode = "PM" Then
	strPostPage = "pm_new_message.asp" & strQsSID1

'If this is new post or poll post to new_post.asp page
ElseIf strMode = "new" OR  strMode = "poll" OR strMode = "reply" OR strMode = "QuickToFull" OR strMode = "quote" Then
	strPostPage = "new_post.asp?PN=" & Trim(Mid(Request.Querystring("PN"), 1, 3)) & strQsSID2

'Else this must be an edit
Else
	strPostPage = "edit_post.asp?PN=" & Trim(Mid(Request.Querystring("PN"), 1, 3)) & strQsSID2
End If



'Turn off sintaures if disabled for group
If blnGroupSignatures = False Then blnSignatures = False
	

%>
<form method="post" name="frmMessageForm" id="frmMessageForm" action="<% = strPostPage %>" onSubmit="return CheckForm();" onReset="return clearForm();">
 <table cellspacing="0" cellpadding="2" align="center"><%


'If the poster is in a guest then get them to enter a name
If lngLoggedInUserID = 2 AND (strMode <> "edit" AND strMode <> "editTopic" AND strMode <> "editPoll") Then
%>
  <tr>
   <td align="right" width="15%"><% = strTxtName %>:</td>
   <td align="left" width="60%">
    <input type="text" name="Gname" id="Gname" size="20" maxlength="20" tabindex="1" />
   </td>
  </tr><%

End If




'If this is a private message display the username box
If strMode = "PM" Then
%>
  <tr>
   <td align="right" width="15%"><% = strTxtToUsername %>:</td>
   <td align="left"  width="60%"><%


         'Get the users buddy list if they have one

	'Initlise the sql statement
	strSQL = "SELECT " & strDbTable & "Author.Username " & _
	"FROM " & strDbTable & "Author" & strDBNoLock & ", " & strDbTable & "BuddyList" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Author.Author_ID=" & strDbTable & "BuddyList.Buddy_ID " & _
		"AND " & strDbTable & "BuddyList.Author_ID=" & lngLoggedInUserID & " AND " & strDbTable & "BuddyList.Buddy_ID <> 2 " & _
	"ORDER BY " & strDbTable & "Author.Username ASC;"

	'Query the database
	rsCommon.Open strSQL, adoCon
%>
    <input type="text" name="member" id="member" size="15" maxlength="25" value="<% = strBuddyName %>"<% If NOT rsCommon.EOF Then Response.Write(" onchange=""document.frmMessageForm.selectMember.options[0].selected = true;""") %> tabindex="2" />
    <a href="javascript:winOpener('pop_up_member_search.asp<% = strQsSID1 %>','memSearch',0,1,580,355)"><img src="<% = strImagePath %>member_search.<% = strForumImageType %>" alt="<% = strTxtFindMember %>" title="<% = strTxtFindMember %>" border="0" align="absmiddle"></a><%

	'If there are records returned then display the users buddy list
        If NOT rsCommon.EOF Then

	Response.Write(vbCrLf & "    " & strSelectFormBuddyList & ":" & _
	vbCrLf & "    <select name=""selectMember"" onchange=""member.value=''"" tabindex=""3"">" & _
	vbCrLf & "     <option value="""">-- " & strTxtNoneSelected & " --</option>")

          	'Loop throuhgn and display the buddy list
          	Do While NOT rsCommon.EOF

          		Response.Write(vbCrLf & "     <option value=""" & rsCommon("Username") & """>" & rsCommon("Username") & "</option>")

           		'Move to next record in rs
           		rsCommon.MoveNext
           	Loop

	Response.Write(vbCrLf & "    </select>")

	Else
		Response.Write(vbCrLf & "    <input type=""hidden"" name=""selectMember"" id=""selectMember"" value="""" />")
        End If

        'Reset server variables
	rsCommon.Close
%>
   </td>
  </tr><%

End If




'If this is a new post or editing the first thread then display the subject text box
If strMode = "new" OR strMode="editTopic" OR strMode = "editPoll" OR strMode = "PM" OR strMode = "poll" Then
%>
  <tr>
   <td align="right"><% = strTxtSubjectFolder %>:</td>
   <td align="left" width="60%">
    <input type="text" name="subject" id="subject" size="30" maxlength="50"<% If strMode="editTopic" OR strMode = "editPoll" OR strMode="PM"  Then Response.Write(" value=""" & strTopicSubject & """") %> tabindex="2" /><%

        'If this is the forums moderator or forum admim then let them slect the priority level of the post
	If (blnAdmin OR blnPriority) AND (strMode = "new" or strMode="editTopic" OR strMode = "editPoll" OR strMode = "poll" or strMode = "editPoll") Then

		Response.Write("&nbsp;" & strTxtPinnedTopic & ":" & _
		vbCrLf & "    <select name=""priority"" id=""priority"" tabindex=""3""")
		If blnDemoMode Then Response.Write(" disabled=""disabled""")
		Response.Write(">" & _
		vbCrLf & "     <option value=""0""")
		If intTopicPriority = 0 Then Response.Write(" selected")
		Response.Write(">" & strTxtNormal & "</option>" & _
		vbCrLf & "     <option value=""1""")
		If intTopicPriority = 1 Then Response.Write(" selected")
		Response.Write(">" & strTxtPinnedTopic & "</option>")

         	'If this is the forum admin or moderator let them post an annoucment to this forum
         	If blnAdmin = True OR blnModerator Then

			Response.Write(vbCrLf & "     <option value=""2""")
			If intTopicPriority = 2 Then Response.Write(" selected")
			Response.Write(">" & strTopThisForum & "</option>")

        	End If

         	'If this is the forum admin let them post a priority post to all forums
         	If blnAdmin = True Then

			Response.Write(vbCrLf & "     <option value=""3""")
			If intTopicPriority = 3 Then Response.Write(" selected")
			Response.Write(">" & strTxtTopAllForums & "</option>")

		End If

		Response.Write("    </select>")

	End If
%>
   </td>
  </tr><%




	'Display message icon drop down
	If blnTopicIcon AND NOT strMode = "PM" Then

		'Get the topic icon array
		%><!--#include file="topic_icon_inc.asp" -->
  <tr>
   <td align="right" width="10%"><% = strTxtMessageIcon %>:</td>
   <td align="left" width="60%">
    <select name="icon" id="icon" onchange="(T_icon.src = icon.options[icon.selectedIndex].value)" tabindex="4">
     <option value="<% = strImagePath %>blank_smiley.gif"<% If strTopicIcon = "" Then Response.Write(" selected") %>><% = strTxtNoneSelected %></option><%

		'Loop through to display topic icons
		For intLoop = 1 TO Ubound(saryTopicIcon)

			Response.Write(vbCrLf & "     <option value=""" & saryTopicIcon(intLoop,2) & """")
			If strTopicIcon = saryTopicIcon(intLoop,2) Then Response.Write(" selected")
			Response.Write(">" & saryTopicIcon(intLoop,1) & "</option>")
		Next

		'If no topic Icon the get the default one
		If strTopicIcon = "" Then strTopicIcon = strImagePath & "blank_smiley.gif"
%>
    </select>
    &nbsp;&nbsp;<img src="<% = strTopicIcon %>" border="0" id="T_icon" alt"<% = strTxtMessageIcon %>" />
   </td>
  </tr><%

	End If


	'*************** Event Start *******************

	'Display Calendar event date input
	If blnCalendar AND blnEvents AND NOT strMode = "PM" Then
%>
  <tr>
   <td align="right" valign="top"><% = strTxtCalendarEvent %>:</td>
   <td align="left">
    <% = strTxtStartDate %>:
    <br />
    &nbsp;&nbsp;&nbsp;&nbsp;<% = strTxtDay %>
    <select name="eventDay" id="eventDay" tabindex="5">
     <option value="0"<% If intEventDay = 0 Then Response.Write(" selected") %>>----</option><%

		'Create lists day's for birthdays
		For intLoop = 1 to 31
			Response.Write(vbCrLf & "     <option value=""" & intLoop & """")
			If intEventDay = intLoop Then Response.Write(" selected")
			Response.Write(">" & intLoop & "</option>")
		Next

%>
    </select>
    <% = strTxtCMonth %>
    <select name="eventMonth" id="eventMonth" tabindex="6">
     <option value="0"<% If intEventMonth = 0 Then Response.Write(" selected") %>>---</option><%

		'Create lists of days of the month for birthdays
		For intLoop = 1 to 12
			Response.Write(vbCrLf & "     <option value=""" & intLoop & """")
			If intEventMonth = intLoop Then Response.Write(" selected")
			Response.Write(">" & intLoop & "</option>")
		Next

%>
    </select>
    <% = strTxtCYear %>
    <select name="eventYear" id="eventYear" tabindex="7">
     <option value="0"<% If intEventYear = 0 Then Response.Write(" selected") %>>-----</option><%

		'If this is an old event and the date is from a previous year, display that year
		If intEventYear <> 0 AND intEventYear < CInt(Year(Now())) Then Response.Write(VbCrLf & "     <option value=""" & intEventYear & """ selected>" & intEventYear & "</option>")

		'Create lists of years for birthdays
		For intLoop = CInt(Year(Now())) to CInt(Year(Now()))+1
			Response.Write(vbCrLf & "     <option value=""" & intLoop & """")
			If intEventYear = intLoop Then Response.Write(" selected")
			Response.Write(">" & intLoop & "</option>")
		Next

%>
    </select>
    <br />
    <% = strTxtEndDate %>: 
    <br />
    &nbsp;&nbsp;&nbsp;&nbsp;<% = strTxtDay %>
    <select name="eventDayEnd" id="eventDayEnd" tabindex="8"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
     <option value="0"<% If intEventDayEnd = 0 Then Response.Write(" selected") %>>----</option><%

		'Create lists day's for birthdays
		For intLoop = 1 to 31
			Response.Write(vbCrLf & "     <option value=""" & intLoop & """")
			If intEventDayEnd = intLoop Then Response.Write(" selected")
			Response.Write(">" & intLoop & "</option>")
		Next

%>
    </select>
    <% = strTxtCMonth %>
    <select name="eventMonthEnd" id="eventMonthEnd" tabindex="9"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
     <option value="0"<% If intEventMonthEnd = 0 Then Response.Write(" selected") %>>---</option><%

		'Create lists of days of the month for birthdays
		For intLoop = 1 to 12
			Response.Write(vbCrLf & "     <option value=""" & intLoop & """")
			If intEventMonthEnd = intLoop Then Response.Write(" selected")
			Response.Write(">" & intLoop & "</option>")
		Next

%>
    </select>
    <% = strTxtCYear %>
    <select name="eventYearEnd" id="eventYearEnd" tabindex="10"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
     <option value="0"<% If intEventYear = 0 Then Response.Write(" selected") %>>-----</option><%

		'If this is an old event and the date is from a previous year, display that year
		If intEventYearEnd <> 0 AND intEventYearEnd < CInt(Year(Now())) Then Response.Write(VbCrLf & "     <option value=""" & intEventYearEnd & """ selected>" & intEventYearEnd & "</option>")

		'Create lists of years for birthdays
		For intLoop = CInt(Year(Now())) to CInt(Year(Now()))+1
			Response.Write(vbCrLf & "     <option value=""" & intLoop & """")
			If intEventYearEnd = intLoop Then Response.Write(" selected")
			Response.Write(">" & intLoop & "</option>")
		Next

%>
    </select>
    <span class="smText">(<% = strTxtNotRequiredForSingleDateEvents %>)</span>
   </td>
  </tr><%

	End If
	
	'*************** Event End *******************
End If


	




'If this is a new poll then display space to enter the poll
If strMode = "poll" OR strMode = "editPoll" Then

	%><!--#include file="poll_form_inc.asp" --><%

End If


'The message textarea
%>
  <tr>
   <td valign="top" align="right" width="15%"><br /><br /><% If blnMobileBrowser = False Then Response.Write("<br /><br />") %><% = strTxtMessage %>:<%


'*************** Emoticons *******************

'If emoticons are enabled show them next to the post window
If blnEmoticons Then

%>
   <table border="0" cellspacing="0" cellpadding="4" align="center">
    <tr><td class="smText" colspan="3" align="center"><br /><% = strTxtEmoticons %></td></tr><%

	'Intilise the index position (we are starting at 1 instead of position 0 in the array for simpler calculations)
	intIndexPosition = 1

	'Calcultae the number of outer loops to do
	intNumberOfOuterLoops = 5

	'If there is a remainder add 1 to the number of loops
	If UBound(saryEmoticons) MOD 2 > 0 Then intNumberOfOuterLoops = intNumberOfOuterLoops + 1

	'Loop throgh th list of emoticons
	For intLoop = 1 to intNumberOfOuterLoops


	        Response.Write("<tr>")

		'Loop throgh the list of emoticons
		For intInnerLoop = 1 to 3

			'If there is nothing to display show an empty box
			If intIndexPosition > UBound(saryEmoticons) Then
				Response.Write(vbCrLf & "    <td>&nbsp;</td>")

			'Else show the emoticon
			Else
				If RTEenabled() <> "false" AND blnRTEEditor AND blnWYSIWYGEditor Then
					Response.Write(vbCrLf & "    <td><img src=""" & saryEmoticons(intIndexPosition,3) & """ border=""0"" title=""" & saryEmoticons(intIndexPosition,1) & """ alt=""" & saryEmoticons(intIndexPosition,1) & """ onclick=""AddEmoticon(this)"" id=""" & saryEmoticons(intIndexPosition,3) & """ style=""cursor: pointer;""></td>")
				Else
					Response.Write(vbCrLf & "    <td><img src=""" & saryEmoticons(intIndexPosition,3) & """ border=""0"" title=""" & saryEmoticons(intIndexPosition,1) & """ alt=""" & saryEmoticons(intIndexPosition,1) & """ onclick=""AddEmoticon('" & saryEmoticons(intIndexPosition,2) & "')"" id=""" & saryEmoticons(intIndexPosition,3) & """ style=""cursor: pointer;""></td>")
	              		End If
	              	End If

	              'Minus one form the index position
	              intIndexPosition = intIndexPosition + 1
		Next

		Response.Write("</tr>")

	Next

	If RTEenabled() <> "false" AND blnRTEEditor AND blnWYSIWYGEditor Then
		Response.Write(vbCrLf & "    <tr><td colspan=""3"" align=""center""><a href=""javascript:winOpener('RTE_popup_emoticons.asp" & strQsSID1 & "','emot',0,0,650,340)"" class=""smLink"" tabindex=""100"">" & strTxtMore & "</a></td></tr>")
	Else
		Response.Write(vbCrLf & "    <tr><td colspan=""3"" align=""center""><a href=""javascript:winOpener('non_RTE_popup_emoticons.asp" & strQsSID1 & "','emot',0,0,650,340)"" class=""smLink"" tabindex=""100"">" & strTxtMore & "</a></td></tr>")
	End If
%>
   </table>
  </td><%
End If

'******************************************

%>
   <td width="60%" valign="top"><%


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
Response.Write(vbCrLf & "    <script language=""JavaScript"" src=""RTE_javascript_common.asp"" type=""text/javascript""></script>")


'If this is an RTE enabled web browser load in the RTE content
If RTEenabled() <> "false" AND blnRTEEditor AND blnWYSIWYGEditor Then
	
	
	'If this is a quick reply to full reply then load up the session variable to pass across to the full RTE area
	If InStr(strMode, "QuickToFull") AND strMessage <> "" Then Session("Message") = strMessage


	'If this is Gecko based browser link to JS code for Gecko
	If (RTEenabled = "Gecko" OR RTEenabled = "opera") AND blnEmoticons Then Response.Write(vbCrLf & "    <script language=""JavaScript"" src=""RTE_javascript_gecko.asp"" type=""text/javascript""></script>")


	'Load in Javascript for RTE browsers (creating an IFrame for the RTE area)
%>
    <script language="JavaScript" src="RTE_javascript.asp?textArea=<% = Server.URLEncode(strTextAreaName) & "&M=" & strMode & strQsSID2 %>" type="text/javascript"></script>
    <script language="javascript" type="text/javascript">
       WebWizRTEtoolbar('<% = strFormName %>');
       document.write ('<iframe id="WebWizRTE" src="RTE_textarea.asp?mode=<% = strMode %>&ID=<% = lngMessageID %>&CACHE=<% = CInt(RND * 2000) & strQsSID2 %>" width="650" height="300" style="border: #A5ACB2 1px solid" onLoad="initialiseWebWizRTE();" tabindex="20"></iframe>');
    </script>
    <noscript><strong><br /><br /><% = strTxtJavaScriptEnabled %></strong></noscript>
    <input type="hidden" name="message" id="message" value="" />
    <input type="hidden" name="browser" id="browser" value="RTE" /><%

'If this is not an RTE enabled web browser load in the NON-RTE content
Else
%>
    <script language="JavaScript" src="non_RTE_javascript.asp?textArea=<% = Server.URLEncode(strTextAreaName) & strQsSID2 %>" type="text/javascript"></script>
    <script language="JavaScript" type="text/javascript">WebWizRTEtoolbar('<% = strFormName %>');</script>
     <textarea name="message" id="message" rows="<% If blnMobileBrowser Then Response.Write("10") Else Response.Write("18") %>" wrap="virtual" tabindex="20" style="width:<% If blnMobileBrowser Then Response.Write("95%") Else Response.Write("596px") %>;" /><%

	If strMode="edit" OR strMode = "quote" OR strMode = "new_comment_quote" OR strMode = "QuickToFull" or strMode="editTopic" or strMode="PM" Then Response.Write(vbCrLf & strMessage)

%></textarea><%

End If

'********************************************************************

%>
   </td>
  </tr>
  <tr>
   <td align="right">&nbsp;</td>
   <td align="left" valign="bottom"><%
   	
'If rel=nofollow the display a message
If blnNoFollowTagInLinks AND blnMobileBrowser = False Then Response.Write("&nbsp;<span class=""smText"">" & strTxtNoFollowAppliedToAllLinks & "</span><br />")

%>&nbsp;<input type="checkbox" name="forumCodes" value="True" checked tabindex="21" /><% = strTxtEnable %> <a href="javascript:winOpener('BBcodes.asp<% = strQsSID1 %>','codes',1,1,610,500)"><% = strTxtForumCodes %></a></td>
  </tr><%





'If not PM then display another row
If NOT strMode = "PM" Then

	'If signature of e-mail notify then display row to show
	If (blnLoggedInUserEmail AND blnEmail) OR (blnLoggedInUserSignature AND blnSignatures) Then

%>
  <tr>
   <td align="right">&nbsp;</td>
   <td align="left" valign="bottom"><%
   	
		'If the user has a signature offer them the chance to show it
		If blnLoggedInUserSignature AND blnSignatures Then

			Response.Write(vbCrLf & "    &nbsp;<input type=""checkbox"" name=""signature"" id=""signature"" value=""True""")
			If blnAttachSignature = True Then Response.Write(" checked")
			Response.Write(" tabindex=""22"" />" & strTxtShowSignature & "&nbsp;")

		End If

		'Display e-mail notify of replies option
		If blnEmail AND blnLoggedInUserEmail Then

			Response.Write(vbCrLf & "    &nbsp;<input type=""checkbox"" name=""email"" id=""email"" value=""True""")
			If blnReplyNotify Then Response.Write(" checked")
			Response.Write(" tabindex=""23"" />" & strTxtEmailNotify & "&nbsp;")

		End If

%>
   </td>
  </tr><%


	End If

'If this is a private e-mail and e-mail is on and the user gave an e-mail address let them choose to be notified when pm msg is read
ElseIf strMode = "PM" AND blnEmail AND blnLoggedInUserEmail Then

%>
   <tr>
   <td align="right" width="92">&nbsp;</td>
   <td valign="bottom">&nbsp;<input type="checkbox" name="email" id="email" value="True" tabindex="24"><span><% = strTxtEmailNotifyWhenPMIsRead %></span></td>
  </tr><%

End If



'Display CAPTCHA images for Guest posting
If blnCAPTCHAsecurityImages AND lngLoggedInUserID = 2 Then 
	
	%>
 <tr>
  <td align="right" valign="top"><% = strTxtUniqueSecurityCode %>:</td>
  <td><!--#include file="CAPTCHA_form_inc.asp" --><span class="smText"><% = strTxtEnterCAPTCHAcode %></span></td>
 </tr><%
     
End If

%>
 <tr>
  <td>
   <input type="hidden" name="mode" id="mode" value="<% = strMode %>" /><%

'Add hidden fields 
If NOT strMode = "PM" Then

    	Response.Write(vbCrLf & "    <input type=""hidden"" name=""FID"" id=""FID"" value=""" & intForumID & """ />")
	Response.Write(vbCrLf & "    <input type=""hidden"" name=""TID"" id=""TID"" value=""" & lngTopicID & """ />")
%>
    <input type="hidden" name="PID" id="PID" value="<% = lngMessageID %>" />
    <input type="hidden" name="PN" id="PN" value="<% = intRecordPositionPageNum %>" /><%
         'If reply get the thread position number in the topic
         If strMode = "reply" Then
         	Response.Write(vbCrLf & "    <input type=""hidden"" name=""ThreadPos"" id=""ThreadPos"" value=""" & (lngTotalRecords + 1) & """ />")
	End If
End If

	 'Ajax session ping (60,000 = 1 minute)
	 Response.Write("" & _
	 vbCrLf & "    <input type=""hidden"" name=""pre"" id=""pre"" value="""" />"  & _
	 vbCrLf & "      <input type=""hidden"" name=""formID"" id=""formID"" value=""" & strFormID & """ />"  & _
         vbCrLf & "      <div id=""ajaxFormFields""><input type=""hidden"" name=""session"" id=""session"" value=""false"" /></div>"  & _
         vbCrLf & "    <script  language=""JavaScript"">getFormData(); function getFormData(){getAjaxData('ajax_session_alive.asp?XID=" & strFormID & strQsSID3 & "','ajaxFormFields'); setTimeout('getFormData()',120000);}</script>")
%>
  </td>
  <td width="60%" align="left">
   <input type="submit" id="Submit" name="Submit" <%

'Select the correct button name for the page
If strMode="edit" OR strMode = "editTopic" OR strMode = "editPoll" Then

	Response.Write("value=""" & strTxtUpdatePost & """ ")

ElseIf strMode = "new_comment" OR strMode = "new_comment_QuickToFull" OR strMode = "new_comment_quote" Then
	
	Response.Write("value=""" & strTxtPostComment & """ ")
	
ElseIf strMode = "edit_comment"	Then
	
	Response.Write("value=""" & strTxtUpdateComment & """ ")
	
ElseIf strMode = "new" OR strMode = "poll" Then

	Response.Write("value=""" & strTxtPostNewTopic2 & """ ")

ElseIf strMode = "PM" Then

	Response.Write("value=""" & strTxtSendPM & """ ")
Else

	Response.Write("value=""" & strTxtPostReply & """ ")
End If


'If RTE enabled then use javascript to submit the RTE data
If RTEenabled() <> "false" AND blnRTEEditor AND blnWYSIWYGEditor Then
	Response.Write("onclick=""document.getElementById('message').value=document.getElementById('WebWizRTE').contentWindow.document.body.innerHTML;""")
End If

%> tabindex="26" />
   <input type="button" name="Preview" id="Preview" value="<% = strTxtPreview %>" onclick="<%
   
'If RTE enabled get the preview content from the iframe
If RTEenabled() <> "false" AND blnRTEEditor AND blnWYSIWYGEditor Then
	Response.Write("document.getElementById('pre').value=document.getElementById('WebWizRTE').contentWindow.document.body.innerHTML;")
Else
	Response.Write("document.getElementById('pre').value=document.getElementById('message').value;")
End If

%> OpenPreviewWindow(document.frmMessageForm);" tabindex="27" />
   <input type="reset" name="Reset" id="Reset" value="<% = strTxtClearForm %>" tabindex="28" />
  </td>
 </tr>
</table>
</form>