<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
<!--#include file="RTE_configuration/RTE_setup.asp" -->
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





'Clean up
Call closeDatabase()


'If the page is opening in SSL then keep it running in SSL
If Request.ServerVariables("HTTPS") = "on" Then blnSslEnabledPage = True


'Dimension veriables
Dim lngRows
Dim lngCols
Dim lngWidth
Dim strWidthValue
Dim strAlign
Dim lngBorder
Dim lngPad
Dim lngSpace
Dim lngRowsLoopCounter
Dim lngColsLoopCounter

'Intalise varibales
lngWidth = 100
lngCols = 1
lngWidth = 1


'If this a post back read in the form elements
If isNumeric(Request.Form("rows")) AND isNumeric(Request.Form("cols")) AND Request.Form("postBack") Then
	
	'Get form elements
	If isNumeric(Request.Form("rows")) Then lngRows = LngC(Request.Form("rows"))
	If isNumeric(Request.Form("cols")) Then lngCols = LngC(Request.Form("cols"))
	If isNumeric(Request.Form("width")) Then lngWidth = LngC(Request.Form("width"))
	strWidthValue = Request.Form("range")
	strAlign = Request.Form("align")
	If isNumeric(Request.Form("border")) Then lngBorder = LngC(Request.Form("border"))
	If isNumeric(Request.Form("pad")) Then lngPad = LngC(Request.Form("pad"))
	If isNumeric(Request.Form("space")) Then lngSpace = LngC(Request.Form("space"))	
End If


%>	
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Table Properties</title>

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


'If this a post back write javascript
If isNumeric(Request.Form("rows")) AND isNumeric(Request.Form("cols")) AND Request.Form("postBack") Then
	
	'If this is Gecko based browser link to JS code for Gecko
	If RTEenabled = "Gecko" Then Response.Write(vbCrLf & "<script language=""JavaScript"" src=""RTE_javascript_gecko.asp"" type=""text/javascript""></script>")
	
	
	Response.Write(vbCrLf & "<script  language=""JavaScript"">")

%>	
    	   	
    	editor = window.opener.document.getElementById('WebWizRTE');
    	
    	rows = <% = lngRows %>;
    	cols = <% = lngCols %>;
   
    	if ((rows > 0) && (cols > 0)) {
      
	      	table = editor.contentWindow.document.createElement("table");
	      	
	      	table.setAttribute("border", "<% = lngBorder %>");
	      	table.setAttribute("cellpadding", "<% = lngPad %>");
	      	table.setAttribute("cellspacing", "<% = lngSpace %>");
	      	table.setAttribute("align", "<% = strAlign %>");
	      	table.setAttribute("width", "<% = lngWidth & strWidthValue %>");
	      
	      	tbody = editor.contentWindow.document.createElement("tbody");
      
      		for (var rowNo=0; rowNo < rows; rowNo++) {
        
        		tr = editor.contentWindow.document.createElement("tr");
        
        		for (var colNo=0; colNo < cols; colNo++) {
          
		          	td = editor.contentWindow.document.createElement("td");
		          	tr.appendChild(td);<%      
      
      		'If this is Mozilla then we need to place a <br> tag in the table cells
      		If RTEenabled = "Gecko" Then %>
		          	br = editor.contentWindow.document.createElement("br");
		          	td.appendChild(br);<%
		End If
		
		%>        
        		}
        
        	tbody.appendChild(tr);
      		}
      
      		table.appendChild(tbody);<%      
      
      		'If this is Mozilla or Opera then we need to call insertElementPosition to find where to place the table
      		If RTEenabled = "Gecko" OR RTEenabled = "opera" Then 
			
			Response.Write(vbCrLf & vbCrLf & "	try{" & _
						vbCrLf & "		insertElementPosition(editor.contentWindow, table);" & _
						vbCrLf & "	}catch(exception){" & _
						vbCrLf & "		alert('" & strTxtErrorInsertingObject & "');" & _
						vbCrLf & "		editor.contentWindow.focus();" & _
						vbCrLf & "	}")
		
		'Else this is IE so it's simpler to place in the table
		Else
			Response.Write(vbCrLf & vbCrLf & "	try{" & _
					vbCrLf & "		editor.contentWindow.focus();" & _
					vbCrLf & "		editor.contentWindow.document.selection.createRange().pasteHTML(table.outerHTML);" & _
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
	
	
	Response.Write("</script>")

End If

%>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body style="margin:0px;" OnLoad="self.focus();">
<table width="100%"  border="0" cellpadding="3" cellspacing="0" class="RTEtableTopRow">
<form method="post">
  <tr class="RTEtableTopRow">
    <td><h1><% = strTxtTableProperties %></h1></td>
  </tr>
  <tr>
    <td class="RTEtableRow"><table width="100%" border="0" align="center" cellpadding="1" cellspacing="0">
        <tr>
          <td colspan="2">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td width="28%" align="right" class="text"><% = strTxtRows %>:</td>
                <td width="4%"><input name="rows" type="text" id="rows" value="2" size="2" maxlength="2" autocomplete="off" /></td>
                <td width="23%" align="right">&nbsp;</td>
                <td width="45%">&nbsp;</td>
              </tr>
              <tr>
                <td align="right" class="text"><% = strTxtColumns %>:</td>
                <td><input name="cols" type="text" id="cols" value="2" size="2" maxlength="2" autocomplete="off" /></td>
                <td align="right" class="text"><% = strTxtWidth %>:</td>
                <td><input name="width" type="text" id="width" value="100" size="3" maxlength="3" autocomplete="off" />
                    <select name="range" id="range">
                      <option value="%" selected>%</option>
                      <option><% = strTxtpixels %></option>
                    </select>
                </td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td width="50%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td class="text"><% = strTxtLayout %></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td align="right" class="text"><% = strTxtAlignment %>:</td>
                <td><select size="1" name="align" id="align">
                    <option value="" selected>Default</option>
                    <option value="left">Left</option>
                    <option value="center">Center</option>
                    <option value="right">Right</option>
                </select></td>
              </tr>
              <tr>
                <td align="right" class="text"><% = strTxtBorder %> :</td>
                <td><input name="border" type="text" id="border" value="1" size="2" maxlength="2" autocomplete="off" /></td>
              </tr>
          </table></td>
          <td width="50%">
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td class="text"><% = strTxtSpacing %></td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td width="50%" align="right" class="text"><% = strTxtCellPad %>
                  :</td>
                <td width="50%"><input name="pad" type="text" id="pad" value="1" size="2" maxlength="2" autocomplete="off" /></td>
              </tr>
              <tr>
                <td align="right" class="text"><% = strTxtCellSpace %>
                  :</td>
                <td>
                  <input name="space" type="text" id="space" value="1" size="2" maxlength="2" autocomplete="off" /></td>
              </tr>
          </table></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td align="right" class="RTEtableBottomRow"><input type="hidden" name="postBack" value="true">
      <input type="submit" name="Submit" value="     <% = strTxtOK %>     ">&nbsp;<input type="button" name="cancel" value=" <% = strTxtCancel %> " onclick="window.close()">
<br /><br />
</td>
  </tr>
  </form>
</table>
</body>
</html>
