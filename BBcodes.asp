<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="includes/emoticons_inc.asp" -->
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




Response.Buffer = True

'Declare variables
Dim intIndexPosition		'Holds the idex poistion in the emiticon array
Dim intNumberOfOuterLoops	'Holds the outer loop number for rows
Dim intLoop			'Holds the loop index position
Dim intInnerLoop		'Holds the inner loop number for columns

'Reset Server Objects
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title><% = strTxtForumCodes %></title>
<meta name="generator" content="Web Wiz Forums" />

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>

<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body OnLoad="self.focus();">
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
  <tr>
    <td align="center"><h1><% = strTxtForumCodes %></h1></td>
  </tr>
   <tr>
    <td align="center"><% = strTxtYouCanUseForumCodesToFormatText %></td>
  </tr>
</table>
<br />
<table cellspacing="1" cellpadding="3" class="tableBorder" align="center" width="510">
    <tr class="tableLedger">
      <td width="65%"><% = strTxtTypedForumCode %></td>
      <td width="35%"><% = strTxtConvetedCode %></td>
    <tr />
    <tr class="tableSubLedger">
        <td colspan="2"><% = strTxtTextFormating %></td>
       </tr>
       <tr class="tableRow">
        <td>[B]<% = strTxtBold %>[/B]</td>
        <td><strong><% = strTxtBold %></strong></td>
       </tr>
       <tr class="tableRow">
        <td>[I]<% = strTxtItalic %>[/I]</td>
        <td><i><% = strTxtItalic %></i></td>
       </tr>
       <tr class="tableRow">
        <td>[U]<% = strTxtUnderline %>[/U]</td>
        <td><u><% = strTxtUnderline %></u></td>
       </tr>
       <tr class="tableRow">
        <td>[CENTER]<% = strTxtCentre %>[/CENTER]</td>
        <td align="center"><% = strTxtCentre %></td>
       </tr>
       <tr class="tableSubLedger">
        <td colspan="2"><% = strTxtImagesAndLinks %></td>
       </tr>
       <tr class="tableRow">
        <td>[IMG]http://example.com/smiley.gif[/IMG]</td>
        <td><img src="smileys/smiley4.gif" width="17" height="17"></td>
       </tr>
       <tr class="tableRow">
        <td>[URL=http://www.example.com]<% = strTxtMyLink %>[/URL]</td>
        <td><a href="#"><% = strTxtMyLink %></a></td>
       </tr>
       <tr class="tableRow">
        <td>[URL]http://www.example.com[/URL]</td>
        <td><a href="#">http://www.example.com</a></td>
       </tr>
       <tr class="tableRow">
        <td>[EMAIL=me@example.com]<% = strTxtMyEmail %>[/EMAIL]</td>
        <td><a href="#"><% = strTxtMyEmail %></a></td>
       </tr>
       <tr class="tableSubLedger">
        <td colspan="2"><% = strTxtFontTypes %></td>
       </tr>
       <tr class="tableRow">
        <td>[FONT=Arial]Arial[/FONT]</td>
        <td><font face="Arial, Helvetica, sans-serif" size="2">Arial</font></td>
       </tr>
       <tr class="tableRow">
        <td>[FONT=Courier]Courier[/FONT]</td>
        <td><font face="Courier New, Courier, mono">Courier</font></td>
       </tr>
       <tr class="tableRow">
        <td>[FONT=Times]Times[/FONT]</td>
        <td><font face="Times New Roman, Times, serif">Times</font></td>
       </tr>
       <tr class="tableRow">
        <td>[FONT=Verdana]Verdana[/FONT]</td>
        <td><font face="Verdana, Arial, Helvetica, sans-serif" size="2">Verdana</font></td>
       </tr>
       <tr class="tableSubLedger">
        <td colspan="2"><% = strTxtFontSizes %></td>
       </tr>
       <tr class="tableRow">
        <td>[SIZE=1]<% = strTxtFontSize %> 1[/SIZE]</td>
        <td><font size="1"><% = strTxtFontSize %> 1</font></td>
       </tr>
       <tr class="tableRow">
        <td>[SIZE=2]<% = strTxtFontSize %> 2[/SIZE]</td>
        <td><font size="2"><% = strTxtFontSize %> 2</font></td>
       </tr>
       <tr class="tableRow">
        <td>[SIZE=3]<% = strTxtFontSize %> 3[/SIZE]</td>
        <td><font size="3"><% = strTxtFontSize %> 3</font></td>
       </tr>
       <tr class="tableRow">
        <td>[SIZE=4]<% = strTxtFontSize %> 4[/SIZE]</td>
        <td><font size="4"><% = strTxtFontSize %> 4</font></td>
       </tr>
       <tr class="tableRow">
        <td>[SIZE=5]<% = strTxtFontSize %> 5[/SIZE]</td>
        <td><font size="5"><% = strTxtFontSize %> 5</font></td>
       </tr>
       <tr class="tableRow">
        <td>[SIZE=6]<% = strTxtFontSize %> 6[/SIZE]</td>
        <td><font size="6"><% = strTxtFontSize %> 6</font></td>
       </tr>
       <tr class="tableSubLedger">
        <td colspan="2"><% = strTxtFontColours %></td>
       </tr>
       <tr class="tableRow">
        <td>[COLOR=BLACK]<% = strTxtBlack & " " & strTxtFont %>[/COLOR]</td>
        <td><font color="black"><% = strTxtBlack & " " & strTxtFont %></font></td>
       </tr>
       <tr class="tableRow">
        <td>[COLOR=WHITE]<% = strTxtWhite & " " & strTxtFont %>[/COLOR]</td>
        <td><font color="white"><% = strTxtWhite & " " & strTxtFont %></font></td>
       </tr>
       <tr class="tableRow">
        <td>[COLOR=BLUE]<% = strTxtBlue & " " & strTxtFont %>[/COLOR]</td>
        <td><font color="blue"><% = strTxtBlue & " " & strTxtFont %></font></td>
       </tr>
       <tr class="tableRow">
        <td>[COLOR=RED]<% = strTxtRed  & " " & strTxtFont %>[/COLOR]</td>
        <td><font color="red"><% = strTxtRed & " " & strTxtFont %></font></td>
       </tr>
       <tr class="tableRow">
        <td>[COLOR=GREEN]<% = strTxtGreen & " " & strTxtFont %>[/COLOR]</td>
        <td><font color="green"><% = strTxtGreen & " " & strTxtFont %></font></td>
       </tr>
       <tr class="tableRow">
        <td>[COLOR=YELLOW]<% = strTxtYellow & " " & strTxtFont %>[/COLOR]</td>
        <td><font color="yellow"><% = strTxtYellow & " " & strTxtFont %></font></td>
       </tr>
       <tr class="tableRow">
        <td>[COLOR=ORANGE]<% = strTxtOrange & " " & strTxtFont %>[/COLOR]</td>
        <td><font color="orange"><% = strTxtOrange & " " & strTxtFont %></font></td>
       </tr>
       <tr class="tableRow">
        <td>[COLOR=BROWN]<% = strTxtBrown & " " & strTxtFont %>[/COLOR]</td>
        <td><font color="brown"><% = strTxtBrown & " " & strTxtFont %></font></td>
       </tr>
       <tr class="tableRow">
        <td>[COLOR=MAGENTA]<% = strTxtMagenta & " " & strTxtFont %>[/COLOR]</td>
        <td><font color="magenta"><% = strTxtMagenta & " " & strTxtFont %></font></td>
       </tr>
       <tr class="tableRow">
        <td>[COLOR=CYAN]<% = strTxtCyan & " " & strTxtFont %>[/COLOR]</td>
        <td><font color="cyan"><% = strTxtCyan & " " & strTxtFont %></font></td>
       </tr>
       <tr class="tableRow">
        <td>[COLOR=LIME GREEN]<% = strTxtLimeGreen & " " & strTxtFont %>[/COLOR]</td>
        <td><font color="limegreen"><% = strTxtLimeGreen & " " & strTxtFont %></font></td>
       </tr>
       <tr class="tableSubLedger">
        <td colspan="2"><% = strTxtQuoting %></td>
       </tr>
       <tr class="tableRow">
        <td colspan="2">[Quote=Username]<% = strTxtQuotedMessage & " " & strTxtWithUsername %>[/QUOTE]</td>
       </tr>
       <tr class="tableRow">
        <td colspan="2">[Quote]<% = strTxtQuotedMessage %>[/QUOTE]</td>
       </tr>
       <tr class="tableSubLedger">
        <td colspan="2"><% = strTxtCodeandFixedWidthData %></td>
       </tr>
       <tr class="tableRow">
        <td colspan="2">[CODE]<% = strTxtMyCodeData %>[/CODE]</td>
       </tr>
       <tr class="tableSubLedger">
        <td colspan="2"><% = strTxtHideContent %></td>
       </tr>
       <tr class="tableRow">
        <td colspan="2">[HIDE]<% = strTxtPostContentHiddenUntilReply %>[/HIDE]</td>
       </tr>
       <tr class="tableSubLedger">
        <td colspan="2"><% = strTxtSpoilerContent %></td>
       </tr>
       <tr class="tableRow">
        <td colspan="2">[SPOILER]<% = strTxtHiddenSpolierContentShowenWhehnButtonClicked %>[/SPOILER]</td>
       </tr><%
'If Falsh
If blnFlashFiles Then 
	%>
       <tr class="tableSubLedger">
        <td colspan="2"><% = strTxtFlashFilesImages %></td>
       </tr>
       <tr class="tableRow">
        <td colspan="2">[FLASH WIDTH=50 HEIGHT=50]http://www.example.com/flash.swf[/FLASH]</td>
       </tr><%
End If

'If YouTube
If blnYouTube Then 
	%>
       <tr class="tableSubLedger">
        <td colspan="2"><% = strTxtYouTube %></td>
       </tr>
       <tr class="tableRow">
        <td colspan="2">[TUBE]<% = strTxtFileName %>[/TUBE]</td>
       </tr><%
End If


If blnEmoticons Then 
	%>
       <tr class="tableSubLedger">
        <td colspan="2"><% = strTxtEmoticons %></td>
       </tr>
       <tr class="tableRow">
        <td colspan="2">
         <table width="100%" border="0" cellspacing="0" cellpadding="4"><%

'Intilise the index position (we are starting at 1 instead of position 0 in the array for simpler calculations)
intIndexPosition = 1

'Calcultae the number of outer loops to do
intNumberOfOuterLoops = UBound(saryEmoticons) / 2

'If there is a remainder add 1 to the number of loops
If UBound(saryEmoticons) MOD 2 > 0 Then intNumberOfOuterLoops = intNumberOfOuterLoops + 1

'Loop throgh th list of emoticons
For intLoop = 1 to intNumberOfOuterLoops

        Response.Write("<tr  class=""tableRow"">")

	'Loop throgh th list of emoticons
	For intInnerLoop = 1 to 2

		'If there is nothing to display show an empty box
		If intIndexPosition > UBound(saryEmoticons) Then
			Response.Write("<td width=""5"" class=""text"">&nbsp;</td>")
			Response.Write("<td width=""92"" class=""text"">&nbsp;</td>")
			Response.Write("<td width=""44"" class=""text"">&nbsp;</td>")
		'Else show the emoticon
		Else
			Response.Write("<td width=""5"" class=""text""><img src=""" & saryEmoticons(intIndexPosition,3) & """ border=""0"" alt=""" & saryEmoticons(intIndexPosition,2) & """></td>")
                	Response.Write("<td width=""92"" class=""text"" nowrap>" & saryEmoticons(intIndexPosition,1) & "</td>")
                	Response.Write("<td width=""44"" class=""text"">" & saryEmoticons(intIndexPosition,2) & "</td>")
              	End If

              'Minus one form the index position
              intIndexPosition = intIndexPosition + 1
	Next

	Response.Write("</tr>")
Next
%></table>
        </td>
       </tr><%
End If
      %>
</table>
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
  <tr>
    <td align="center"><br /><input type="button" name="closeWin" onclick="javascript:window.close();" value="<% = strTxtCloseWindow %>"><br /></td>
  </tr>
</table>
<div align="center"><%

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	If blnTextLinks = True Then
		Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
	Else
  		Response.Write("<a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank"" ><img src=""webwizforums_image.asp"" border=""0"" title=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ alt=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ /></a>")
	End If

	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")
End If
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
</div>
</body>
</html>