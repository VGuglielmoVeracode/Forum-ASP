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




'Set the response buffer to true
Response.Buffer = False 


'Clean up
Call closeDatabase()




'Check to see if the user is using an access database
If strDatabaseType <> "Access" Then
	
	'Display message to sql server usres
	Response.Write "This page only works on Access but your Database Type is set to SQL Server"
	Response.End
End If

'Dimension variables
Dim objJetEngine		'Holds the jet database engine object
Dim objFSO			'Holds the FSO object
Dim strCompactDB		'Holds the destination of the compacted database


%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Compact and Repair Access Database</title>

<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Response.Write("<!--//" & _
vbCrLf & "/* *******************************************************" & _
vbCrLf & "Software: Web Wiz Forums(TM) ver. " & strVersion & "" & _
vbCrLf & "Info: http://www.webwizforums.com" & _
vbCrLf & "Copyright: (C)2001-2019 Web Wiz Ltd. All rights reserved" & _
vbCrLf & "******************************************************* */" & _
vbCrLf & "//-->" & vbCrLf & vbCrLf)
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
%>
    	
<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center"> 
 <h1>Compact and Repair Access Database</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br />
  From here you can Compact and Repair the Access database used for the Forum making the database size smaller and the forum perform faster.<br>
  This feature can also repair a damaged or corrupted database.<br>
 </p><%

'If this is a post back run the compact and repair
If Request.Form("postBack") Then 
	
	
	'Check the form ID to prevent XCSRF
	Call checkFormID(Request.Form("formID"))
	
%>
<table width="80%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
   <td class="text"><ol><%
   	

 
 	'Create an intence of the FSO object
 	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
 	
 	'Back up the database
 	objFSO.CopyFile strDbPathAndName, Replace(strDbPathAndName, ".mdb", "-backup.mdb", 1, -1, 1)
 	
 	Response.Write("	<li>Database backed up to:-<br/><span class=""smText"">" & Replace(strDbPathAndName, ".mdb", "-backup.mdb", 1, -1, 1) & "</span><br /><br /></li>")


	'Create an instance of the JET engine object
	Set objJetEngine = Server.CreateObject("JRO.JetEngine")

	'Get the destination and name of the compacted database
	strCompactDB = Replace(strDbPathAndName, ".mdb", "-tmp.mdb", 1, -1, 1)

	'Compact database
	objJetEngine.CompactDatabase strCon, "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & strCompactDB
	
	'Display text that new compact db is created
	Response.Write("	<li>New compacted database:-<br/><span class=""smText"">" & strCompactDB & "</span><br /><br /></li>")
	
	'Release Jet object
	Set objJetEngine = Nothing
	
	
	
	
	'Delete old database
	objFSO.DeleteFile strDbPathAndName
	
	'Display text that that old db is deleted
	Response.Write("	<li>Old uncompacted database deleted:-<br/><span class=""smText"">" & strDbPathAndName & "</span><br /><br /></li>")
	
	
	
	'Rename temporary database to old name
	objFSO.MoveFile strCompactDB, strDbPathAndName
	
	'Display text that that old db is deleted
	Response.Write("	<li>Rename compacted database from:-<br/><span class=""smText"">" & strCompactDB & "</span><br />To:-<br /><span class=""smText"">" & strDbPathAndName & "</span><br /><br /></li>")
	

	'Release FSO object
	Set objFSO = Nothing
	
	
	Response.Write("	The Forums Access database is now compacted and repaired")

%></ol></td>
  </tr>
 </table>
<%
Else

%>
 <p class="text"> Please note: If the 'Compact and Repair' procedure fails a backup of your database will be created ending with '-backup.mdb'.<br />
 </p>
</div>
<form action="admin_compact_access_db.asp<% = strQsSID1 %>" method="post" name="frmCompact" id="frmCompact">
 <div align="center"><br />
  <input name="postBack" type="hidden" id="postBack" value="true">
  <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
  <input type="submit" name="Submit" value="Compact and Repair Database">
 </div>
</form><%

End If

%>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->
