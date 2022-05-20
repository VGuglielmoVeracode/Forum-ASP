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





'Set the timeout higher so that it doesn't timeout half way through
Session.Timeout =  1000

'Set the response buffer to true as we maybe redirecting
Response.Buffer = True


'If in demo mode redirect
If blnDemoMode Then
	Call closeDatabase()
	Response.Redirect("admin_web_wiz_forums_premium.asp" & strQsSID1)
End If




'Dimension variables
Dim lngDelAuthorID		'Holds the authors ID to be deleted
Dim intNoOfDays			'Holds the number of days to delete posts from
Dim lngNumberOfMembers		'Holds the number of members that are deleted
Dim rsThread			'Holds the threads recordset
Dim blnUnActive			'Set to true if deleting non active accounts only

'Initilise variables
lngNumberOfMembers = 0



'Check the form ID to prevent XCSRF
Call checkFormID(Request.Form("formID"))


'get the number of days to delete from
intNoOfDays = IntC(Request.Form("days"))
blnUnActive = BoolC(Request.Form("unactive"))



'Get all the Topics from the database to be deleted

'Initalise the strSQL variable with an SQL statement to get the topic from the database
strSQL = "SELECT " & strDbTable & "Author.Author_ID " & _
"FROM " & strDbTable & "Author " & _
"WHERE (" & strDbTable & "Author.Join_date < " & strDatabaseDateFunction & " - " & intNoOfDays  & " " & _
	"AND " & strDbTable & "Author.No_of_posts = 0) " & _
	"AND " & strDbTable & "Author.Author_ID > 2 "
If blnUnActive = True Then 
	strSQL = strSQL & "AND " & strDbTable & "Author.Active = " & strDBFalse & " "
End If
strSQL = strSQL & ";"


'Query the database
rsCommon.Open strSQL, adoCon


'Create a record set object to the Threads held in the database
Set rsThread = Server.CreateObject("ADODB.Recordset")


'Loop through all all the members to delete
Do While NOT rsCommon.EOF

	'Get the author ID
	lngDelAuthorID = CLng(rsCommon("Author_ID"))


	'Check to make sure that there isn't any posts by the member
	strSQL = "SELECT" & strDBTop1 & " " & strDbTable & "Thread.Thread_ID " & _
		"FROM " & strDbTable & "Thread " & _
		"WHERE " & strDbTable & "Thread.Author_ID = " & lngDelAuthorID & strDBLimit1 & ";"

	'Query the database
	rsThread.Open strSQL, adoCon
	
	'If there are no posts start deleting
	If rsThread.EOF Then
		
		'Delete the members buddy list
		'Initalise the strSQL variable with an SQL statement
		strSQL = "DELETE FROM " & strDbTable & "BuddyList WHERE (Author_ID = "  & lngDelAuthorID & ") OR (Buddy_ID ="  & lngDelAuthorID & ");"
		
		'Execute SQL
		adoCon.Execute(strSQL)	
		
		
		'Delete the members private msg's
		strSQL = "DELETE FROM " & strDbTable & "PMMessage WHERE (Author_ID = "  & lngDelAuthorID & ");"
			
		'Execute SQL
		adoCon.Execute(strSQL)	
		
		
		'Delete the members private msg's
		strSQL = "DELETE FROM " & strDbTable & "PMMessage WHERE (From_ID = "  & lngDelAuthorID & ");"
			
		'Execute SQL
		adoCon.Execute(strSQL)
		
		
		'Set all the users private messages to Guest account
		strSQL = "UPDATE " & strDbTable & "PMMessage SET From_ID = 2 WHERE (From_ID = "  & lngDelAuthorID & ");"
			
		'Execute SQL
		adoCon.Execute(strSQL)
		
		
		'Set all the users posts to the Guest account
		strSQL = "UPDATE " & strDbTable & "Thread SET Author_ID = 2 WHERE (Author_ID = "  & lngDelAuthorID & ");"
			
		'Execute SQL
		adoCon.Execute(strSQL)
				
		
		'Delete the user from the email notify table
		strSQL = "DELETE FROM " & strDbTable & "EmailNotify WHERE (Author_ID = "  & lngDelAuthorID & ");"
			
		'Execute SQL
		adoCon.Execute(strSQL)
		
		
		'Delete the user from forum permissions table
		strSQL = "DELETE FROM " & strDbTable & "Permissions WHERE (Author_ID = "  & lngDelAuthorID & ");"
			
		'Execute SQL
		adoCon.Execute(strSQL)
		
		
		'Delete the user from forum author table
		strSQL = "DELETE FROM " & strDbTable & "Author WHERE (Author_ID = "  & lngDelAuthorID & ");"
			
		'Execute SQL
		adoCon.Execute(strSQL)
		
		
		'Total up the number of members deleted
		lngNumberOfMembers = lngNumberOfMembers + 1
	End If
	
	'Close the recordset
	rsThread.Close

	'Move to the next record
	rsCommon.MoveNext
Loop



'Reset Server Objects
Set rsThread = Nothing
rsCommon.Close
Call closeDatabase()

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<meta name="generator" content="Web Wiz Forums" />
<title>Batch Delete Members</title>

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
 <h1>Batch Delete Members </h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a><br />
  <br>
  <br>
  <br>
  <br>
  <br />
  <% = lngNumberOfMembers %> Members have been Deleted.<br />
 </p>
</div>
<!-- #include file="includes/admin_footer_inc.asp" -->
