<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="RTE_configuration/RTE_setup.asp" -->
<!--#include file="common.asp" -->
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




Response.Buffer = True


'Clean up
Call closeDatabase()


'If the page is opening in SSL then keep it running in SSL
If Request.ServerVariables("HTTPS") = "on" Then blnSslEnabledPage = True



'Declare variables
Dim intIndexPosition		'Holds the idex poistion in the emiticon array
Dim intNumberOfOuterLoops	'Holds the outer loop number for rows
Dim intLoop			'Holds the loop index position
Dim intInnerLoop		'Holds the inner loop number for columns
Dim intASCIINo			'Holds the ascii number of the char
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="copyright" content="Copyright (C) 2001-2010 Web Wiz" />
<title>Special Characters</title>

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Rich Text Editor(TM) ver. " & strRTEversion & "" & _
vbCrLf & "Info: http://www.richtexteditor.org" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******


'If this is Gecko based browser or Opera link to JS code for Gecko
If RTEenabled = "Gecko" OR RTEenabled = "opera" Then Response.Write(vbCrLf & "<script language=""JavaScript"" src=""RTE_javascript_gecko.asp"" type=""text/javascript""></script>")
	
%>	
<script  language="JavaScript">
//Function add special character
function AddSpecialChar(spChar) {	

	editor = window.opener.document.getElementById('WebWizRTE');<%
	
'If this is Mozilla or Opera then we need to call insertElementPosition to find where to place the image
If RTEenabled = "Gecko" OR RTEenabled = "opera" Then 
		
	Response.Write(vbCrLf & vbCrLf & "	try{" & _
				vbCrLf & "		insertElementPosition(editor.contentWindow, editor.contentWindow.document.createTextNode(spChar.id));" & _
				vbCrLf & "	}catch(exception){" & _
				vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
				vbCrLf & "		editor.contentWindow.focus();" & _
				vbCrLf & "	}")
	
'Else this is IE so placing the image is simpler
Else
	Response.Write(vbCrLf & vbCrLf & "	try{" & _
				vbCrLf & "		editor.contentWindow.focus();" & _
				vbCrLf & "		editor.contentWindow.document.selection.createRange().pasteHTML(spChar.id);" & _
				vbCrLf & "	}catch(exception){" & _
				vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
				vbCrLf & "		editor.contentWindow.focus();" & _
				vbCrLf & "	}")
End If

'Set focus
'If Opera change the focus method
If RTEenabled = "opera" Then
		
	Response.Write(vbCrLf & "	editor.focus();")
Else
	Response.Write(vbCrLf & "	editor.contentWindow.focus();")
End If
%>
	window.close();
}
 <%


'Use the following for non IE5 users
If RTEenabled <> "winIE5" Then
%>
//Function to hover special char
function overSpChar(specialCharacter) {
	
	specialCharacter.className='RTEmouseOver';
	document.getElementById("spCharName").value = specialCharacter.id;
}

//Function to moving off special char
function outSpChar(specialCharacter) {
	
	specialCharacter.className='';
	document.getElementById("spCharName").value = '';
}
<%


End If 

%>
</script>
<style type="text/css">
input.display {
	border: 0px;	
	font-family: Arial, Helvetica, sans-serif;
	font-size: 36px;
	color: #000000;
	text-align: center;
	height: 45px;
	width: 45px;
}
</style>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body style="margin:0px;" OnLoad="self.focus();">
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="RTEtableTopRow">
  <form method="post" name="frmImageInsrt">
    <tr class="RTEtableTopRow">
      <td colspan="2"><h1><% = strTxtSpecialCharacters %></h1></td>
    </tr>
    <tr>
      <td colspan="2" width="85%" class="RTEtableRow">
      <table width="100%" border="0" cellspacing="1" cellpadding="0"><%

'Initiliase ascii no
intASCIINo = 32

'Intilise the index position (we are starting at 1 instead of position 0 in the array for simpler calculations)
intIndexPosition = 1

'Calcultae the number of outer loops to do
intNumberOfOuterLoops = 215 / 20

'If there is a remainder add 1 to the number of loops
If 215 MOD 20 > 0 Then intNumberOfOuterLoops = intNumberOfOuterLoops + 1

'Loop throgh th list of characters
For intLoop = 1 to intNumberOfOuterLoops
      

	Response.Write(vbCrLf & "	<tr class=""text"">")

	'Loop throgh the list of characters
	For intInnerLoop = 1 to 20  
	
		'Calculate ascii no
		intASCIINo = intASCIINo + 1
		
		'Miss out some chars that don't display correctly or are spaces or delete
		If intASCIINo = 127 Then intASCIINo = 128
		If intASCIINo = 129 Then intASCIINo = 130
		If intASCIINo = 141 Then intASCIINo = 142
		If intASCIINo = 143 Then intASCIINo = 145
		If intASCIINo = 157 Then intASCIINo = 158
		If intASCIINo = 160 Then intASCIINo = 161
		If intASCIINo = 173 Then intASCIINo = 174
	
		'If there is nothing to display show an empty box
		If intIndexPosition > 215 Then 
			Response.Write(vbCrLf & "          <td width=""20"" height=""20"" class=""RTEbutton""><img width=""1"" height=""1""></td>") 

		'Else show the character
		Else 
			Response.Write(vbCrLf & "          <td width=""20"" height=""20"" class=""RTEbutton"" id=""&#" & intASCIINo & ";"" title=""&#" & intASCIINo & ";""  onMouseover=""overSpChar(this)"" onMouseout=""outSpChar(this)"" OnClick=""AddSpecialChar(this)"" align=""center"" style=""cursor: default;"">&#" & intASCIINo & ";</td>")
              	End If
              
              'Minus one form the index position
              intIndexPosition = intIndexPosition + 1 
	Next    
	        
	Response.Write(vbCrLf & "	</tr>")
	
Next             
%></table>
      </td>
      <td width="15%" align="center" valign="top" class="RTEtableRow"><input name="spCharName" type="text" class="RTEtableRow" style="border: 0px; text-align: center; font-size: 36px;" id="spCharName" value="" size="1" maxlength="1" readonly="readonly">
      </td>
    </tr>
    <tr>
     <td class="RTEtableBottomRow" valign="top">&nbsp;<%

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnAbout Then
	Response.Write("<span class=""text"" style=""font-size:10px""><a href=""http://www.richtexteditor.org"" target=""_blank"" style=""font-size:10px"">Web Wiz Rich Text Editor</a> version " & strRTEversion & "</span>")
End If 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******      
      
      %></td>
      <td colspan="2" align="right" class="RTEtableBottomRow">&nbsp;<input type="button" name="cancel" value=" <% = strTxtCancel %> " onclick="window.close()"><br />
      <br /><br />
      </td></tr>
  </form>
</table>
</body>
</html>