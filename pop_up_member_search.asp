<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="common.asp" -->
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


Dim strUsername		'Holds the users username
Dim saryMembers		'Holds the recordset array of members
Dim intCurrentRecord	'Holds the current records for the posts
Dim lngTotalRecords	'Holds the total number of therads in this topic


blnSslEnabledPage = True


'If this is a post back then search the member list
If Request.Form("name") <> "" Then

	'Read in the username
	strUsername = Request.Form("name")
	
	'Get rid of milisous code
	strUsername = formatSQLInput(strUsername)
	
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Author.Username " & _
	"FROM " & strDbTable & "Author" & strDBNoLock & " " & _
	"WHERE " & strDbTable & "Author.Username Like '" & strUsername & "%' " & _
	"ORDER BY " & strDbTable & "Author.Username ASC;"
		
	'Query the database
	rsCommon.Open strSQL, adoCon
	
	'Place into an array
	If NOT rsCommon.EOF Then 
		saryMembers = rsCommon.GetRows
		
		'Count the number of records
		lngTotalRecords = Ubound(saryMembers,2) + 1
	End If
	
	'Clean up
	rsCommon.Close
End If


'Clean up
Call closeDatabase()
%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title><% = strTxtFindMember %></title>
<%

'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write(vbCrLf & vbCrLf & "<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->")
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******

%>
 	
<script  language="JavaScript">

//Function to check form is filled in correctly before submitting
function CheckForm () {

	var errorMsg = "";
	
	//Check for a Username
	if (document.getElementById('frmMemSearch').name.value==""){
	
		msg = "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine1 %>\n";
		msg += "<% = strTxtErrorDisplayLine2 %>\n";
		msg += "<% = strTxtErrorDisplayLine %>\n\n";
		msg += "<% = strTxtErrorDisplayLine3 %>\n";
	
		alert(msg + "\n<% = strTxtErrorUsername %>");
		document.getElementById('frmMemSearch').name.focus();
		return false;
	}
	
	return true;
}


//Function to place the username in the text box of the opening frame
function getUserName(selectedName)
{
	
	window.opener.document.<% 
	
	If Request.QueryString("RP") = "BUD" Then 
		Response.Write("frmBuddy.username") 
	ElseIf Request.QueryString("RP") = "SEARCH" Then 
		Response.Write("frmSearch.USR") 
	Else 
		Response.Write("frmMessageForm.member")
	End If 
	
	%>.focus();
	window.opener.document.<% 
	
	If Request.QueryString("RP") = "BUD" Then 
		Response.Write("frmBuddy.username") 
	ElseIf Request.QueryString("RP") = "SEARCH" Then 
		Response.Write("frmSearch.USR") 
	Else 
	  	Response.Write("frmMessageForm.member") 
	End If 
	
	%>.value = selectedName;
	window.close();
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body style="margin:0px;" OnLoad="self.focus();">
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
  <tr>
    <td align="center"><h1><% = strTxtFindMember %></h1></td>
  </tr>
</table>
<br />
<form method="post" name="frmMemSearch" id="frmMemSearch" action="pop_up_member_search.asp?<% 

'Set the return page
If Request.QueryString("RP") = "BUD" Then 
	Response.Write("RP=BUD")
ElseIf Request.QueryString("RP") = "SEARCH" Then
	Response.Write("RP=SEARCH")
End If
	
	%><% = strQsSID2 %>" onSubmit="return CheckForm();">
  <table cellspacing="1" cellpadding="3" class="tableBorder" align="center" style="width: 550px;">
    <tr class="tableLedger">
      <td colspan="2"><% = strTxtFindMember %></td>
    </tr>
    <tr class="tableRow">
     <td colspan="2" class="smText"><% = strTxtTypeTheNameOfMemberInBoxBelow %> </td>
    </tr>
    <tr class="tableRow">
      <td width="32%" height="30" align="right"><% = strTxtMemberSearch %>: </td>
      <td width="68%"><input type="text" name="name" id="name" size="15" maxlength="15" value="<% = strUsername %>" /> <input type="submit" name="Submit" id="Submit" value="<% = strTxtSearch %>" /></td>
    </tr><%
    
'If this is a post back then display the results
If Request.Form("name") <> "" Then

	If lngTotalRecords > 0 then
%>        
    <tr class="tableRow">
     <td colspan="2" class="smText"><% = strTxtSelectNameOfMemberFromDropDownBelow %> </td>
    </tr><%

	End If

%>
    <tr class="tableRow"> 
      <td height="30" align="right"><% = strTxtSelectMember %>: </td>
      <td><select name="userName" id="userName"><%

	'If there are no records found then display an error message
	If lngTotalRecords = 0 then
		
		Response.Write(vbCrLf & "        <option value="""" selected>" & strTxtNoMatchesFound & "</option>") 
	
	'Else there are matches found so display the result
	Else   
	
		'Do....While Loop to loop through the recorset to display the topic posts
		Do While intCurrentRecord < lngTotalRecords
		
			'Disply the usernames found
			Response.Write(vbCrLf & "        <option value=""" & saryMembers(0,intCurrentRecord) & """>" & saryMembers(0,intCurrentRecord) & "</option>")       
           		
           		'Jump to the next record
           		intCurrentRecord = intCurrentRecord + 1
           
           	Loop
           
        End If        
        %>
       </select> <input type="button" name="Button" id="Button" value="<% = strTxtSelect %>" onclick="getUserName(frmMemSearch.userName.options[frmMemSearch.userName.selectedIndex].value);" />
      </td>
   </tr><% 
        
        
        
End If        

%>
 </table>
</form>
<br />
<table class="basicTable" cellspacing="0" cellpadding="3" align="center">
  <tr>
    <td align="center"><input type="button" name="ok" onclick="javascript:window.close();" value="<% = strTxtCloseWindow %>"><br />
      <br /><% 
    
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
If blnLCode = True Then
	If blnTextLinks = True Then 
		Response.Write("<span class=""text"" style=""font-size:10px"">Forum Software by <a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank""  style=""font-size:10px"">Web Wiz Forums&reg;</a> version " & strVersion & "</span>")
	Else
  		Response.Write("<a href=""http://www.webwizforums.com"" rel=""nofollow"" target=""_blank"" ><img src=""webwizforums_image.asp"" border=""0"" title=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ alt=""Forum Software by Web Wiz Forums&reg; version " & strVersion& """ /></a>")
	End If
	
	Response.Write("<br /><span class=""text"" style=""font-size:10px"">Copyright &copy;2001-2019 Web Wiz Ltd.</span>")

'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
End If
%>
    </td>
  </tr>
</table>
</body>