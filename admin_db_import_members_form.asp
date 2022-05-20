<% @ Language=VBScript %>
<% Option Explicit %>
<!--#include file="admin_common.asp" -->
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




'Set the response buffer to true as we maybe redirecting
Response.Buffer = True 



%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Import Subscribers from External Database</title>
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
<script language="JavaScript" type="text/javascript">

//Function to check form is filled in correctly before submitting
function CheckForm () {
	
	if (document.frmImport.tableName.value==""){
		alert("Please enter the name of the table to import members from");
		document.frmImport.tableName.focus();
		return false;
	}
	if (document.frmImport.usernameField.value==""){
		alert("Please enter the database field name that contains the members Username");
		document.frmImport.nameField.focus();
		return false;
	}
	return true
}
</script>
<link href="<% = strCSSfile %>default_style.css" rel="stylesheet" type="text/css" />
</head>
<body>
<!-- #include file="includes/admin_header_inc.asp" -->
   <h1>Import Members from External Database </h1>
   <a href="admin_menu.asp" target="_self">Return to the the Admin Control Panel Menu</a><br />
   <br />
    <span class="text">
   <%
  
'If member API enabled then display a message to the user
If blnMemberAPI Then 
	Response.Write("   This option is not available if the Member API is enabled.<br /><br />After logging in through your own login system, members will be added to the forum using the Member API when they enter the forum.")
Else
	Response.Write("   This tool allows you to import members into Web Wiz Forums  from an external database.")
End If
%>
    </span><br />
   <br /><%

'If member API enabled do not show the form
If blnMemberAPI = False  Then 
	
%>
   <form action="admin_import_members.asp" method="post" name="frmImport" id="frmImport" onsubmit="return CheckForm();">
    <table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" class="tableBorder">
     <tr>
      <td colspan="2" class="tableLedger">Database Details</td>
     </tr>
     <tr class="text">
      <td class="tableRow"><strong>Database Type</strong>:&nbsp;&nbsp;</td>
      <td class="tableRow"><select name="dbType" id="dbType"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %>>
        <option value="access" selected="selected">MS Access 2000 or above</option>
        <option value="access97">MS Access 97</option>
        <option value="SQLServer">MS SQL Server</option>
        <option value="mySQL">mySQL</option>
       </select></td>
     </tr>
     <tr class="text">
      <td class="tableRow"><strong>Name of Database</strong>: <br />
       <span class="smText">This is the Name of your database or database file name if connecting to an Access database. </span></td>
      <td valign="top" class="tableRow"><input name="dbName" type="text" id="dbName" size="20" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr>
      <td valign="top" class="tableRow"><strong>Database Path </strong> (Access Only): <br />
       <span class="smText">This is the path to your database. Must be located on the same server or network mapped drive.&nbsp; </span></td>
      <td class="tableRow"><input name="location" type="text" id="location" size="40" maxlength="100"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
       <br />
       <input name="locType" type="radio" value="physical"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
       Physical Server Path to Database <span class="smText">(not URL) </span><br />
       <input name="locType" type="radio" value="virtual" checked="checked"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> />
       Path from this application to database </td>
     </tr>
     <tr class="text">
      <td class="tableRow"><strong>Database Server Name or IP </strong>(SQL Server, mySQL only):<span class="smText"><br />
       This is the name or IP address to your MS SQL Server </span></td>
      <td valign="top" class="tableRow"><input name="dbServerIP" type="text" id="dbServerIP" size="20" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr class="text">
      <td class="tableRow"><p>Database  Username:<br />
        <span class="smText">This is the Username you use to login to your database.</span></p></td>
      <td valign="top" class="tableRow"><input name="username" type="text" id="username" size="20" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr class="text">
      <td class="tableRow">Database  Password: <br />
       <span class="smText">This is the Password you use to login to your database. </span></td>
      <td valign="top" class="tableRow"><input name="password" type="text" id="password" size="20" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr>
      <td colspan="2" class="tableLedger">Database Table and Field Details </td>
     </tr>
     <tr class="text">
      <td width="46%" class="tableRow" >Name of Table <span class="smText">(required)</span>:&nbsp;&nbsp;<br />
       <span class="smText">This is the name of the table in your database where the details of the members you wish to import into your Forum are stored. </span></td>
      <td width="54%" valign="top" class="tableRow" ><input name="tableName" type="text" id="tableName" size="20" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr class="text">
      <td class="tableRow" >Username Source Field <span class="smText">(required</span>):<br />
        <span class="smText">This is the name of the field in your database that stores the Usernames of the members you wish to import.</span></td>
      <td valign="top" class="tableRow" ><input name="usernameField" type="text" id="usernameField" size="20" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr class="text">
      <td class="tableRow" >Password Source Field:<br />
        <span class="smText">This is the name of the field in your database that stores the Passwords of your members you are importing. If you don't specify a field for Passwords, a password will be created
         for your member. </span></td>
      <td valign="top" class="tableRow" ><input name="passwordField" type="text" id="passwordField" size="20" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr class="text">
      <td class="tableRow" >Email Address Source Field:<br />
       <span class="smText">This is the name of the field in your database that stores the Email Addresses you wish to import.</span></td>
      <td valign="top" class="tableRow" ><input name="emailField" type="text" id="emailField" size="20" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
      <tr class="text">
      <td class="tableRow">Real Name<br />
       <span class="smText">This is an optional Real Name import and is the name of the field in your database holding the persons real name.</span></td>
      <td valign="top" class="tableRow" ><input name="realNameFirst" type="text" id="realNameFirst" size="20" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
      <tr class="text">
      <td class="tableRow">Real Name (Last Name)<br />
       <span class="smText">If you are using the Real Name import above this is used if the Real Name of the person is stored in two fields, for example when you have a First Name and Last Name fields.</span></td>
      <td valign="top" class="tableRow" ><input name="realNameLast" type="text" id="realNameLast" size="20" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr class="text">
      <td class="tableRow" >Location</td>
      <td valign="top" class="tableRow" ><input name="where" type="text" id="where" size="20" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr class="text">
      <td class="tableRow" >Signature</td>
      <td valign="top" class="tableRow" ><input name="signature" type="text" id="signature" size="20" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr class="text">
      <td class="tableRow" >No of Posts </td>
      <td valign="top" class="tableRow" ><input name="Posts" type="text" id="Posts" size="20" maxlength="50"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
     </tr>
     <tr>
      <td colspan="2" class="tableLedger">Forum Group</td>
     </tr>
     <tr class="tableRow">
      <td colspan="2" class="text">
       <table width="100%"  border="0" cellspacing="1" cellpadding="1">
        <tr class="tableRow">
         <td width="100%" class="tableTopRow" colspan="2">Please select which Web Wiz Forums Memebr Group you would like the Imported Members placed in.</td>
        </tr><%
          
	'Initalise the strSQL variable with an SQL statement to query the database
	strSQL = "SELECT " & strDbTable & "Group.* FROM " & strDbTable & "Group ORDER BY " & strDbTable & "Group.Group_ID ASC;"
		
	'Query the database
	rsCommon.Open strSQL, adoCon

	
	'Loop through cats
	Do While NOT rsCommon.EOF
		
		Response.Write(vbCrLf & "   <tr class=""tableRow""> " & _
		vbCrLf & "    <td width=""1%"" align=""right""><input type=""checkbox"" name=""GID"" id=""GID"" value=""" & rsCommon("Group_ID") & """")
		If blnDemoMode Then Response.Write(" disabled=""disabled""")
		Response.Write(" /></td>" & _
		vbCrLf & "    <td width=""99%"">" & rsCommon("Name") & "</td>" & _
		vbCrLf & "   </tr>")
   
		'Move to next record in rs
		rsCommon.MoveNext
	Loop
	
	'Close RS
	rsCommon.close	

 %>
       </table>      </td>
     </tr>
     <tr class="text">
      <td colspan="2" align="center" class="tableBottomRow"><input name="M" type="hidden" id="M" value="db" />
      <input name="Submit" type="submit" value="Import" /></td>
     </tr>
    </table>
    <br />
   </form><%
End If

'Close DB
Call closeDatabase() 
%>
   <br />
   <!-- #include file="includes/admin_footer_inc.asp" -->
