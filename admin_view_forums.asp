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






'Set the response buffer to true
Response.Buffer = True

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Administer Forums</title>
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

<!-- #include file="includes/admin_header_inc.asp" -->
<div align="center"><h1>Administer Forums</h1><br />
  <a href="admin_menu.asp<% = strQsSID1 %>">Control Panel Menu</a></div>
<form action="admin_update_forum_order.asp<% = strQsSID1 %>" method="post" name="form1" id="form1">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td align="center" class="text">From here you can add, delete, edit, or lock Categories and Forums.<br />
        <br />
        Click on Forum name or Category to Amend Details.<br />
        <br />
        Select the order you would like the forums to be in from the Order drop down list and click on the Update Order button.<br />
        <span class="smText">If forums are not correctly set in an order you may find stability issues within your forums. </span> </td>
    </tr>
  </table>
  <br />
  <%

'Dimension variables
Dim strCategory			'Holds the categories
Dim intCatID			'Holds the category ID number
Dim strForumName		'Holds the forum name
Dim strForumDiscription		'Holds the forum description
Dim blnForumLocked		'Set to true if the forum is locked
Dim intLoop			'Holds the number of times round in the Loop Counter
Dim intNumOfForums		'Holds the number of forums
Dim intForumOrder		'Holds the order number of the forum
Dim intNumOfCategories		'Holds the number of categories
Dim intCatOrder			'Holds the order number of the category
Dim intSubForumID
Dim sarryForums			'Holds the getrows db call for forums
Dim intCurrentRecord		'Holds the current record
Dim intCurrentRecord2		'holds the second current record
Dim saryCategories		'Holds the categories
Dim intCatCurrentRecord		'Holds the currnet record for the cats
Dim strForumURL


'Read the various categories from the database
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT " & strDbTable & "Category.Cat_name, " & strDbTable & "Category.Cat_ID, " & strDbTable & "Category.Cat_order FROM " & strDbTable & "Category ORDER BY " & strDbTable & "Category.Cat_order ASC, " & strDbTable & "Category.Cat_ID ASC;"


'Query the database
rsCommon.Open strSQL, adoCon

'Place the rs into an array
If NOT rsCommon.EOF Then
	saryCategories = rsCommon.GetRows()
End If

'Close rs
rsCommon.close

'Check there are categories to display
If isArray(saryCategories) = false Then


	'If there are no categories to display then display the appropriate error message
	Response.Write vbCrLf & "<br /><br /><br /><span class=""text""><strong>There are no Categories to display. <a href=""admin_category_details.asp" & strQsSID2 & """>Click here to create a Forum Category</a></strong></span><br /><br /><br />"

'Else there the are categories so write the HTML to display categories and the forum names and a discription
Else

	'Get the number of categories
	intNumOfCategories = Ubound(saryCategories,2) + 1

	'Loop round to read in all the categories in the database
	Do While NOT intCatCurrentRecord > Ubound(saryCategories,2)


		'Get the category name from the database
		strCategory = saryCategories(0,intCatCurrentRecord)
		intCatID = CInt(saryCategories(1,intCatCurrentRecord))
		intCatOrder = CInt(saryCategories(2,intCatCurrentRecord))


		'Display the category name
		
		%>
  <table border="0" align="center" cellpadding="2" cellspacing="1" class="tableBorder">
    <tr>
      <td width="77%" height="26" class="tableLedger">Forums</td>
      <td width="6%" align="center" class="tableLedger">Sub&nbsp;Forum</td>
      <td width="6%" height="26" align="center" class="tableLedger">Lock</td>
      <td width="6%" height="26" align="center" class="tableLedger">Delete</td>
      <td width="5%" height="26" align="center" class="tableLedger">Order</td>
    </tr>
    <tr>
      <td colspan="3" class="tableSubLedger"><a href="admin_category_details.asp?CatID=<% = intCatID & strQsSID2 %>"><% = strCategory %></a></td>
      <td align="center" class="tableSubLedger"><a href="admin_delete_category.asp?CatID=<% = intCatID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 %>" onclick="return confirm('Are you sure you want to Delete this Category?\n\nWARNING: Deleting this category will permanently  remove all Forum(s) in this Category and all the Posts!')"><img src="<% = strImagePath %>delete.png" border="0" title="Delete Category" /></a></td>
      <td align="center" class="tableSubLedger"><select name="catOrder<% = intCatID %>">
          <%
          Response.Write(intNumOfCategories)
           'loop round to display the number of forums for the order select list
           For intLoop = 1 to intNumOfCategories
		Response.Write("<option value=""" & intLoop & """ ")

			'If the loop number is the same as the order number make this one selected
			If intCatOrder = intLoop Then
				Response.Write("selected")
			End If

		Response.Write(">" & intLoop & "</option>")
           Next
           %>
        </select>
      </td>
    </tr>
    <%
		'Read the various forums from the database
		'******************************************
		'Initalise the strSQL variable with an SQL statement to query the database
		strSQL = "SELECT " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Sub_ID," & strDbTable & "Forum.Forum_name," & strDbTable & "Forum.Forum_description," & strDbTable & "Forum.Forum_Order," & strDbTable & "Forum.Locked," & strDbTable & "Forum.Forum_URL " & _
		"FROM " & strDbTable & "Forum " & _
		"WHERE " & strDbTable & "Forum.Cat_ID = " & intCatID  & " " & _
		"ORDER BY " & strDbTable & "Forum.Forum_Order ASC;"

		'Query the database
		rsCommon.Open strSQL, adoCon

		'Check there are forum's to display
		If rsCommon.EOF Then

			'If there are no forum's to display then display the appropriate error message
			Response.Write vbCrLf & "<td bgcolor=""#FFFFFF"" colspan=""5"" class=""tableBottomRow"">There are no Forum's to display. <a href=""admin_forum_details.asp" & strQsSID2 & """>Click here to create a Forum</a></td>"
			
			'Close the recordset as it is no longer needed
			rsCommon.Close
			
		'Else there the are forum's to write the HTML to display it the forum names and a discription
		Else

			'Initilise current record
			intCurrentRecord = 0
			
			'Read in the row from the db using getrows for better performance
			sarryForums = rsCommon.GetRows()
			
			'Get the number of forums
			rsCommon.Close
			
			'Get the number of forums
			intNumOfForums = Ubound(sarryForums,2) + 1
			

			'Loop round to read in all the forums in the database
			Do While NOT intCurrentRecord > Ubound(sarryForums,2)
			
				'If this is a subforum jump to the next record, unless we have run out of forums
				Do While CInt(sarryForums(1,intCurrentRecord)) > 0 
					
					'Go to next record
					intCurrentRecord = intCurrentRecord + 1
					
					'If we have run out of records jump out of loop
					If intCurrentRecord > Ubound(sarryForums,2) Then Exit Do
				Loop
				
				'If we have run out of records jump out of loop
				If intCurrentRecord > Ubound(sarryForums,2) Then Exit Do

				'Read in forum details from the database
				intForumID = CInt(sarryForums(0,intCurrentRecord))
				strForumName = sarryForums(2,intCurrentRecord)
				strForumDiscription = sarryForums(3,intCurrentRecord)
				intForumOrder = CInt(sarryForums(4,intCurrentRecord))
				blnForumLocked = CBool(sarryForums(5,intCurrentRecord))
				strForumURL = sarryForums(6,intCurrentRecord)
				
				'If no link then clear this part
				If strForumURL = "http://" Then strForumURL = ""

				'Write the HTML of the forum descriptions and hyperlinks to the forums
				
				%>
    <tr>
      <td class="tableRow"><%
      	
      				'If a forum link
      				If strForumURL <> "" Then

      					Response.Write("Link: <a href=""admin_forum_link.asp?FID=" & intForumID & strQsSID2 & """>" & strForumName & "</a>")
      					Response.Write(" <a href=""" & strForumURL & """ target=""_blank""><img src=""" & strImagePath & "new_window.png"" alt=""Open link in new window"" /></a>")
      				
      				'Else a normal forum
      				Else
      					Response.Write("<a href=""admin_forum_details.asp?FID=" & intForumID & strQsSID2 & """>" & strForumName & "</a>")
      				
      				End If
      	%> <br />
        <span class="smText"><% = strForumDiscription %></span></td>
      <td align="center" class="tableRow"><% 

      				'If not a forum URL the allow it to be turned in to a sub forum
      				If strForumURL = "" OR IsNull(strForumURL) Then 
      					%><a href="admin_change_to_subforum.asp?FID=<% = intForumID %><% = strQsSID2 %>"><img src="<% = strImagePath %>yes.png" border="0" title="Make this a Sub Forum" /></a><% 
      					
      				End If 
      				%></td>
      <td width="6%" align="center" class="tableRow"><%

		            	'If the forum is locked and the user is admin let them unlock it
				If blnForumLocked = True Then
				  	Response.Write ("	<a href=""admin_lock_forum.asp?mode=UnLock&FID=" & intForumID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """ OnClick=""return confirm('Are you sure you want to Un-Lock this Forum?')""><img src=""" & strImagePath & "locked.png"" border=""0"" align=""baseline"" title=""Un-Lock Forum""></a>")
				'If the forum is not locked and this is the admin then let them lock it
				ElseIf blnForumLocked = False Then
				  	Response.Write ("	<a href=""admin_lock_forum.asp?mode=Lock&FID=" & intForumID  & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """ OnClick=""return confirm('Are you sure you want to Lock this Forum?')""><img src=""" & strImagePath & "unlocked.png"" border=""0"" align=""baseline"" title=""Lock Forum""></a>")
				End If

               %>
      </td>
      <td width="6%" align="center" class="tableRow"><a href="admin_delete_forum.asp?FID=<% = intForumID & "&amp;XID=" & getSessionItem("KEY")  & strQsSID2 %>" onclick="return confirm('Are you sure you want to Delete this Forum?\n\nWARNING: Deleting this forum will permanently  remove all Posts in this Forum!')"><img src="<% = strImagePath %>delete.png" border="0" title="Delete Forum" /></a></td>
      <td width="5%" class="tableRow"  align="center"><select name="forumOrder<% = intForumID %>">
          <%

			    	'loop round to display the number of forums for the order select list
			 	For intLoop = 1 to intNumOfForums
			
					Response.Write("<option value=""" & intLoop & """ ")
			
						'If the loop number is the same as the order number make this one selected
						If intForumOrder = intLoop Then
							Response.Write("selected")
						End If
			
					Response.Write(">" & intLoop & "</option>")
			   	Next
           %>
        </select>
      </td>
    </tr>
    <%					
				
				'See if there are any subforums to this forum
				'*********************************************
				
				'Intilise record 2
				intCurrentRecord2 = 0
				
				'Loop round to read in any sub forums in the stored array recordset
				Do While NOT intCurrentRecord2 > Ubound(sarryForums,2)
				
					'If this is a subforum of the main forum show it
					If CInt(sarryForums(1,intCurrentRecord2)) = intForumID Then
						
						'Read in forum details from the database
						intSubForumID = CInt(sarryForums(0,intCurrentRecord2))
						strForumName = sarryForums(2,intCurrentRecord2)
						strForumDiscription = sarryForums(3,intCurrentRecord2)
						intForumOrder = CInt(sarryForums(4,intCurrentRecord2))
						blnForumLocked = CBool(sarryForums(5,intCurrentRecord2))
						strForumURL = sarryForums(6,intCurrentRecord)
						
						'If no link then clear this part
						If strForumURL = "http://" Then strForumURL = ""

						'Write the HTML of the forum descriptions and hyperlinks to the forums
				
				%>
    <tr>
      <td class="tableRow">&nbsp;&nbsp;<img src="<% = strImagePath %>arrow.gif" />&nbsp;<%
      	
      				'If a forum link
      				If strForumURL <> "" Then

      					Response.Write("Link: <a href=""admin_forum_link.asp?FID=" & intSubForumID & strQsSID2 & """>" & strForumName & "</a>")
      					Response.Write(" <a href=""" & strForumURL & """ target=""_blank""><img src=""" & strImagePath & "new_window.png"" alt=""Open link in new window"" /></a>")
      				
      				'Else a normal forum
      				Else
      					Response.Write("<a href=""admin_forum_details.asp?sub=true&amp;FID=" & intSubForumID & strQsSID2 & """>" & strForumName & "</a>")
      				
      				End If
      	%> <br />
        &nbsp;&nbsp;&nbsp;&nbsp;<span class="smText"><% = strForumDiscription %></span></td>
      <td align="center" class="tableRow"><a href="admin_remove_sub_forum.asp?FID=<% = intSubForumID & "&amp;XID=" & getSessionItem("KEY")  & strQsSID2 %>" onclick="return confirm('Are you sure you want to change this Sub Forum into a Main Forum?')"><img src="<% = strImagePath %>no.png" border="0" title="Change to Main Forum" /></a></td>
      <td width="6%" align="center" class="tableRow"><%

				            	'If the forum is locked and the user is admin let them unlock it
						If blnForumLocked = True Then
						  	Response.Write ("	<a href=""admin_lock_forum.asp?mode=UnLock&FID=" & intSubForumID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """ OnClick=""return confirm('Are you sure you want to Un-Lock this Forum?')""><img src=""" & strImagePath & "locked.png"" border=""0"" align=""baseline"" title=""Un-Lock Forum""></a>")
						'If the forum is not lovked and this is the admin then let them lock it
						ElseIf blnForumLocked = False Then
						  	Response.Write ("	<a href=""admin_lock_forum.asp?mode=Lock&FID=" & intSubForumID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 & """ OnClick=""return confirm('Are you sure you want to Lock this Forum?')""><img src=""" & strImagePath & "unlocked.png"" border=""0"" align=""baseline"" title=""Lock Forum""></a>")
						End If

               %>
      </td>
      <td width="6%" align="center" class="tableRow"><a href="admin_delete_forum.asp?FID=<% = intSubForumID & "&amp;XID=" & getSessionItem("KEY") & strQsSID2 %>" onclick="return confirm('Are you sure you want to Delete this Forum?\n\nWARNING: Deleting this forum will permanently  remove all Posts in this Forum!')"><img src="<% = strImagePath %>delete.png" border="0" title="Delete Sub Forum" /></a></td>
      <td width="5%" class="tableRow"  align="center"><select name="forumOrder<% = intSubForumID %>">
          <%

					    	'loop round to display the number of forums for the order select list
					 	For intLoop = 1 to intNumOfForums
					
							Response.Write("<option value=""" & intLoop & """ ")
					
								'If the loop number is the same as the order number make this one selected
								If intForumOrder = intLoop Then
									Response.Write("selected")
								End If
					
							Response.Write(">" & intLoop & "</option>")
					   	Next
           %>
        </select>
      </td>
    </tr>
    <%		
						
					End If
					
				
					'Move to next record
					intCurrentRecord2 = intCurrentRecord2 + 1
				Loop
				
				
				'Move to next record
				intCurrentRecord = intCurrentRecord + 1
			
			'Loop back round for next forum
			Loop
		End If
		
		'Move to the next database record
		intCatCurrentRecord = intCatCurrentRecord + 1
		
		%>
  </table>
  <br />
  <%
		
	'Loop back round for next category
	Loop
	
	'Clean up
	Set rsCommon = Nothing
End If

'Clean up
Call closeDatabase()
%>
  <table width="98%" border="0" cellspacing="0" cellpadding="3" align="center">
    <tr align="right">
     <td width="100%">
      <input type="hidden" name="formID" id="formID" value="<% = getSessionItem("KEY") %>" />
      <input type="submit" name="Submit" value="Update Order"<% If blnDemoMode Then Response.Write(" disabled=""disabled""") %> /></td>
    </tr>
  </table>
  <br />
</form>
<table border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
  <tr>
    <td colspan="2" class="tableLedger">Create Options </td>
  </tr>
  <tr class="tableRow">
    <td colspan="2" align="center" class="text">Use the buttons below to create new, Categories, Forums, or Sub Forums <br />
      <table width="100%" border="0" cellspacing="0" cellpadding="1" align="center">
        <tr align="center">
          <td><form action="admin_category_details.asp<% = strQsSID1 %>" method="post" name="form2" id="form2">
              <input type="submit" name="Submit2" value="Create New Category" />
            </form></td>
          <td><form action="admin_forum_details.asp<% = strQsSID1 %>" method="post" name="form3" id="form3">
              <input type="submit" name="Submit3" value="Create New Forum" />
            </form></td>
          <td><form action="admin_forum_details.asp?sub=true<% = strQsSID2 %>" method="post" name="form4" id="form4">
              <input type="submit" name="Submit4" value="Create New Sub Forum" />
            </form></td>
           <td><form action="admin_forum_link.asp<% = strQsSID1 %>" method="post" name="form5" id="form5">
              <input type="submit" name="Submit5" value="Create New External Link" />
            </form></td>
        </tr>
      </table></td>
  </tr>
</table>
<br />
<!-- #include file="includes/admin_footer_inc.asp" -->