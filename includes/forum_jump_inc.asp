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



'Declare variables
Dim sarryForumJump		'Holds the array with all the forums
Dim strJumpCatName		'Holds the name of the category
Dim intJumpCatID		'Holds the ID number of the category
Dim strJumpForumName		'Holds the name of the forum to jump to
Dim lngJumpFID			'Holds the forum id to jump to
Dim intJumpSubFID		'Holds if the forum is a sub forum
Dim intJumpCurrentRecord	'Holds the current location in the array
Dim intJumpTempRecord		'Temporay record store
Dim blnJumpHideForum		'Holds if the jump forum is hidden or not
Dim blnJumpRead			'Holds if the jump forum if user has access


Response.Write(strTxtForumJump & _
vbCrLf & "   <select onchange=""linkURL(this)"" name=""SelectJumpForum"">" & _
vbCrLf & "    <option value="""" disabled=""disabled"" selected=""selected"">-- " & strTxtSelectForum & " --</option>" & _
vbCrLf & "    <optgroup label=""" & strTxtForums & """>")




'Read the various categories, forums, and permissions from the database in one hit for extra performance
'Initalise the strSQL variable with an SQL statement to query the database
strSQL = "" & _
"SELECT " & strDbTable & "Category.Cat_ID, " & strDbTable & "Category.Cat_name, " & strDbTable & "Forum.Forum_ID, " & strDbTable & "Forum.Sub_ID, " & strDbTable & "Forum.Forum_name, " & strDbTable & "Forum.Hide, " & strDbTable & "Permissions.View_Forum " & _
"FROM " & strDbTable & "Category" & strDBNoLock & ", " & strDbTable & "Forum" & strDBNoLock & ", " & strDbTable & "Permissions" & strDBNoLock & " " & _
"WHERE " & strDbTable & "Category.Cat_ID = " & strDbTable & "Forum.Cat_ID " & _
	"AND " & strDbTable & "Forum.Forum_ID = " & strDbTable & "Permissions.Forum_ID " & _
	"AND (" & strDbTable & "Permissions.Author_ID = " & lngLoggedInUserID & " OR " & strDbTable & "Permissions.Group_ID = " & intGroupID & ") " & _
"ORDER BY " & strDbTable & "Category.Cat_order, " & strDbTable & "Forum.Forum_Order, " & strDbTable & "Permissions.Author_ID DESC;"
	

'Set error trapping
On Error Resume Next
	
'Query the database
rsCommon.Open strSQL, adoCon

'If an error has occurred write an error to the page
If Err.Number <> 0 Then	Call errorMsg("An error has occurred while executing SQL query on database.", "get_forum_jump_data", "forum_jump_inc.asp")
			
'Disable error trapping
On Error goto 0


'Place the recordset into an array
If NOT rsCommon.EOF Then 
	
	'Place the recordset into an array
	sarryForumJump = rsCommon.GetRows()


	'Close the recordset
	rsCommon.Close


	'SQL Query Array Look Up table
	'0 = Cat_ID
	'1 = Cat_name
	'2 = Forum_ID
	'3 = Sub_ID
	'4 = Forum_name
	'5 = Hide
	'6 = Read 
	
	'Loop round to show all the categories and forums
	Do While intJumpCurrentRecord <= Ubound(sarryForumJump,2)
		
		'Loop through the array looking for forums that are to be shown
		'if a forum is found to be displayed then show the category and the forum, if not the category is not displayed as there are no forums the user can access
		Do While intJumpCurrentRecord <= Ubound(sarryForumJump,2)
		
			'Read in details
			blnJumpHideForum = CBool(sarryForumJump(5,intJumpCurrentRecord))
			blnJumpRead = CBool(sarryForumJump(6,intJumpCurrentRecord))
					
			'If this forum is to be shown then leave the loop and display the cat and the forums
			If blnJumpHideForum = False OR blnJumpRead Then Exit Do
			
			'Move to next record
			intJumpCurrentRecord = intJumpCurrentRecord + 1
		Loop
				
		'If we have run out of records jump out of loop
		If intJumpCurrentRecord > Ubound(sarryForumJump,2) Then Exit Do
	 
		
		
		'Read in the deatils for the category
		intJumpCatID = CInt(sarryForumJump(0,intJumpCurrentRecord))
		strJumpCatName = sarryForumJump(1,intJumpCurrentRecord)		
		
		
		'Display category
		Response.Write vbCrLf & "      <optgroup label=""&nbsp;&nbsp;" & strJumpCatName & """>"
		
		
		
		'Loop round to display all the forums for this category
		Do While intJumpCurrentRecord <= Ubound(sarryForumJump,2)
		
			'Read in the forum details from the recordset
			lngJumpFID = CInt(sarryForumJump(2,intJumpCurrentRecord))
			intJumpSubFID = CInt(sarryForumJump(3,intJumpCurrentRecord))
			strJumpForumName = sarryForumJump(4,intJumpCurrentRecord)
			blnJumpHideForum = CBool(sarryForumJump(5,intJumpCurrentRecord))
			blnJumpRead = CBool(sarryForumJump(6,intJumpCurrentRecord))
			
			
			'If this forum is to be hidden but the user is allowed access to it set the hidden boolen back to false
			If blnJumpHideForum AND blnJumpRead Then blnJumpHideForum = False

			'If the forum is not a hidden forum to this user, display it
			If blnJumpHideForum = False AND intJumpSubFID = 0 Then
				'Display a link in the link list to the forum
				Response.Write (vbCrLf & "       <option value=""forum_topics.asp?FID=" & lngJumpFID & strQsSID2 & SeoUrlTitle(strJumpForumName, "&amp;title=") & """>&nbsp;" & strJumpForumName & "</option>")	
			End If
			
			
			
			'See if this forum has any sub forums
			'Initilise variables
			intJumpTempRecord = 0
					
			'Loop round to read in any sub forums in the stored array recordset
			Do While intJumpTempRecord <= Ubound(sarryForumJump,2)
			
				'Becuase the member may have an individual permission entry in the permissions table for this forum, 
				'it maybe listed twice in the array, so we need to make sure we don't display the same forum twice
				If intJumpSubFID = CInt(sarryForumJump(2,intJumpTempRecord)) Then intJumpTempRecord = intJumpTempRecord + 1
				
				'If there are no records left exit loop
				If intJumpTempRecord > Ubound(sarryForumJump,2) Then Exit Do
				
				'If this is a subforum of the main forum then get the details
				If CInt(sarryForumJump(3,intJumpTempRecord)) = lngJumpFID Then
				
					'Read in the forum details from the recordset
					intJumpSubFID = CInt(sarryForumJump(2,intJumpTempRecord))
					strJumpForumName = sarryForumJump(4,intJumpTempRecord)
					blnJumpHideForum = CBool(sarryForumJump(5,intJumpTempRecord))
					blnJumpRead = CBool(sarryForumJump(6,intJumpTempRecord))
					
					
					'If this forum is to be hidden but the user is allowed access to it set the hidden boolen back to false
					If blnJumpHideForum = True AND blnJumpRead = True Then blnJumpHideForum = False
		
					'If the forum is not a hidden forum to this user, display it
					If blnJumpHideForum = False Then
						'Display a link in the link list to the forum
						Response.Write (vbCrLf & "       <option value=""forum_topics.asp?FID=" & intJumpSubFID & strQsSID2 & SeoUrlTitle(strJumpForumName, "&amp;title=") & """>&nbsp;&nbsp;-&nbsp;" & strJumpForumName & "</option>")	
					End If
				End If
				
				'Move to next record 
				intJumpTempRecord = intJumpTempRecord + 1
				
			Loop
			
			
					
			'Move to the next record in the array
			intJumpCurrentRecord = intJumpCurrentRecord + 1
			
			
			'If there are more records in the array to display then run some test to see what record to display next and where				
			If intJumpCurrentRecord <= Ubound(sarryForumJump,2) Then

				'Becuase the member may have an individual permission entry in the permissions table for this forum, 
				'it maybe listed twice in the array, so we need to make sure we don't display the same forum twice
				If lngJumpFID = CInt(sarryForumJump(2,intJumpCurrentRecord)) Then intJumpCurrentRecord = intJumpCurrentRecord + 1
				
				'If there are no records left exit loop
				If intJumpCurrentRecord > Ubound(sarryForumJump,2) Then Exit Do
				
				'See if the next forum is in a new category, if so jump out of this loop to display the next category
				If intJumpCatID <> CInt(sarryForumJump(0,intJumpCurrentRecord)) Then Exit Do
			End If
		Loop
		
		
		Response.Write(vbCrLf & "     </optgroup>")
	Loop
Else

	'Close the recordset
	rsCommon.Close

End If

Response.Write(vbCrLf & "    </optgroup>" & _
vbCrLf & "   </select>")

%>