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




'Clean up
Call closeDatabase()

%>
<!-- #include file="includes/browser_page_encoding_inc.asp" -->
<title>Forum Control Panel Menu</title>
<meta name="generator" content="Web Wiz Forums" />
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
<h1>Web Wiz Forums Admin Control Panel  <% If blnDemoMode Then  Response.write("<br /> ***** DEMO MODE ****") %></h1>
<br />
   <table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
    <tr>
     <td class="tableLedger">Control Panel Menu</td>
    </tr>
    <tr>
     <td class="tableRow"><table width="100%" border="0" cellpadding="15" cellspacing="4">
       <tr>
        <td width="33%" align="center"><a href="admin_menu.asp?C=admin<% = strQsSID2 %>"><img src="forum_images/forum_setup.png" alt="Forum Setup and Admin" border="0" /></a><br />
         <a href="admin_menu.asp?C=admin<% = strQsSID2 %>"><strong>Forum Setup and Admin</strong></a></td>
        <td width="33%" align="center"><a href="admin_menu.asp?C=setup<% = strQsSID2 %>"><img src="forum_images/toolbox.png" alt="Setup and Configuration" border="0" /></a> <br />
         <a href="admin_menu.asp?C=setup<% = strQsSID2 %>"><strong>Setup and Configuration</strong></a> </td>
        <td width="33%" align="center"><a href="admin_menu.asp?C=members<% = strQsSID2 %>"><img src="forum_images/gorups_members.png" alt="Group and Members Tools" border="0" /></a><br />
         <a href="admin_menu.asp?C=members<% = strQsSID2 %>"><strong>Group and Members Tools</strong></a></td>
       </tr>
      </table>
       <table width="100%" border="0" cellpadding="15" cellspacing="4">
        <tr>
         <td width="33%" align="center"><a href="admin_menu.asp?C=security<% = strQsSID2 %>"><img src="forum_images/security.png" alt="Security Settings" border="0" /></a><br />
          <a href="admin_menu.asp?C=security<% = strQsSID2 %>"><strong>Security Settings</strong></a> </td>
         <td width="33%" align="center"><a href="admin_menu.asp?C=tools<% = strQsSID2 %>"><img src="forum_images/tools.png" alt="Forum Tools" border="0" /></a><br />
          <a href="admin_menu.asp?C=tools<% = strQsSID2 %>"><strong>Forum Tools</strong></a></td>
         
         <td width="33%" align="center"><a href="admin_menu.asp?C=upgrades<% = strQsSID2 %>""><img src="forum_images/webwizforums_box_sm.png" alt="Upgrades (Premium Edition)" border="0" /></a><br />
          <a href="admin_menu.asp?C=upgrades<% = strQsSID2 %>"><strong>Upgrades<br /></strong></a></td>
        
        </tr>
      </table></td>
    </tr>
   </table><%
   
'If the database is not moved tell the user there forum is not secure
If strDatabaseType = "Access" AND strDbPathAndName = Server.MapPath("database/wwForum.mdb") Then   
   
%>  
 <br />
 <table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
   <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong>SECURITY ALERT!!</strong>
    <br /><br />
    Your Forums Access database has not been secured.
    <br /><br />
    <a href="http://www.webwizforums.comkb/" target="_blank">View information on how to secure your Forums's Access database.</a>
  </tr>
</table><%

End If


'If the database is not moved tell the user there forum is not secure
If LCase(strLoggedInUsername) = "administrator" AND lngLoggedInUserID = 1 AND blnDemoMode = False Then
   
 %>  
 <br />
 <table class="errorTable" cellspacing="1" cellpadding="3" align="center">
  <tr>
   <td><img src="<% = strImagePath %>error.png" alt="<% = strTxtError %>" /> <strong>SECURITY ALERT!!</strong>
    <br /><br />
    Your Admin Login Username and Password are not secure.
    <br /><br />
    <a href="admin_change_admin_username.asp">Update your Admin Login.</a>
  </tr>
</table><%  
	
End If
   


'If they want forum admin menu
If Request.QueryString("C") = "admin" Then

%>
   <br />
   <table width="100%" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Forum Administration</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_view_forums.asp<% = strQsSID1 %>">Forum Administration</a><br />
   Create, Amend, Delete any forum's and forum categories, alter forum details, set forum permissions, lock forums, password protect forums, etc.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_statistics.asp<% = strQsSID1 %>">Forum Statistics</a><br />
Displays a list of forum statistics.</td>
 </tr>
</table>
   <%

End If

'If they want memebrs and group menu
If Request.QueryString("C") = "members" Then

%>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Group Administration</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_view_groups.asp<% = strQsSID1 %>">Group Administration</a><br />
   Create, Amend, Delete, change the details, set permissions etc. of  Member Groups.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_group_permissions_form.asp<% = strQsSID1 %>">Set Group Permissions</a><br />
Setup or change Group Permissions </td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_group_points.asp<% = strQsSID1 %>">Set Group Points</a><br />
Setup or change The Group Point System for the number of Points Members get for various actions within the forums.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_view_ladder_groups.asp<% = strQsSID1 %>">Ladder Group Administration</a><br />
From this option you can create, delete or edit Ladder Groups.</td>
 </tr>
</table>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Member Administration</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_select_members.asp<% = strQsSID1 %>">Member Administration</a><br />
   Administer members accounts, make them moderators, change status, delete members, suspend accounts, etc.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_register.asp<% = strQsSID1 %>">Add New Member </a><br />
   From this option you can Add a New Forum Member. </td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_find_user.asp<% = strQsSID1 %>">Set  Member Permissions</a><br />
From this option you can configure permissions  for individual Members.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_suspend_registration.asp<% = strQsSID1 %>">Suspend New Registrations</a><br />
   From this option you can Suspend New Users from Registering to use the forum.</td>
 </tr>
</table><%

End If

'If they want setup menu
If Request.QueryString("C") = "setup" Then

%>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Settings and Configuration Tools </td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_web_wiz_forums_settings.asp<% = strQsSID1 %>">Web Wiz Forums General Settings</a><br />
   Configure general settings for your Forum.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_user_settings.asp<% = strQsSID1 %>">User Settings </a><br />
  Configure settings for your users.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_post_topic_configure.asp<% = strQsSID1 %>">Post and Topic Settings</a><br />
  Configure the way Topics and Post look and feel.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_registration_settings.asp<% = strQsSID1 %>">Registration and Profile Settings</a><br />
  Configure what items are compulsory for registration, custom registration items, and how member profiles are display.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_pm_configure.asp<% = strQsSID1 %>">Private Messaging Settings</a><br />
  Configure the Private Messenger.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_calendar_configure.asp<% = strQsSID1 %>">Events Calendar Settings</a><br />
  Configure the Events Calendar.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_chat_room_settings.asp<% = strQsSID1 %>">Chat Room Settings</a><br />
   Configure the settings for the Chat Room.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_mobile_settings.asp<% = strQsSID1 %>">Mobile Device Settings</a><br />
   Configure SmartPhone and Tablet Mobile Device settings for your forum.</td>
 </tr>
  <tr class="tableRow">
  <td><a href="admin_seo_settings.asp<% = strQsSID1 %>">Search Engine Optimisation (SEO) Settings</a><br />
   Configure Search Engine Optimisations (SEO) to your forum and Analytics Tracking.</td>
 </tr>
  <tr class="tableRow">
  <td><a href="admin_rss_configure.asp<% = strQsSID1 %>">RSS Feed Settings</a><br />
  RSS Feeds allow you to syndicate content from your forum. Use this option to configure how your RSS Feeds will work.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_newspad_configure.asp<% = strQsSID1 %>">NewsPad Settings (Mass Email Tool)</a><br />
  Web Wiz NewsPad allows you to send eNewsletters and mass email your Forum Members.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_upload_configure.asp<% = strQsSID1 %>">Upload Settings</a><br />
Allow users to upload images and files in their posts, also setup Avatar uploading.</td>
</tr>
<tr class="tableRow">
  <td><a href="admin_email_notify_configure.asp<% = strQsSID1 %>">Email Settings</a><br />
Configure email settings and enable email features such as email notification, email account activation, etc.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_seo_settings.asp<% = strQsSID1 %>">Search Engine Optimisation (SEO) Settings</a><br />
   Configure Search Engine settings for your forum.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_date_time_configure.asp<% = strQsSID1 %>">Date and Time Settings</a><br />
This option allows you to set the global date and time format as well as any time off-set you need to have if your hosting is within a different time zone. </td>
 </tr>
<tr class="tableRow">
  <td><a href="admin_ads_settings.asp<% = strQsSID1 %>">Affiliate and Text/Banner Ads</a><br />
Monetize your forum content by affiliating links or add Text/Banners Ads to your forum including Google Adsense.</td>
 </tr>
</table>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Skin</td>
 </tr>
 <tr>
  <td class="tableRow"><a href="admin_skin_configure.asp<% = strQsSID1 %>">Configure Forum Skin</a><br />
   From this option you can apply a new skin to your forums and change the name, look, and feel of your forum. </td>
 </tr>
</table>
<%
End If

If (Request.QueryString("C") = "setup" OR Request.QueryString("C") = "admin") Then
	
%>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Web Wiz Forums </td>
 </tr>
 <tr>
  <td class="tableRow"><a href="admin_server_test.asp<% = strQsSID1 %>">Server Requirements Test </a><br />
Check that the server your site is hosted on and your web host have the correct requirements to run Web Wiz Forums. </td>
 </tr>
 <tr>
  <td class="tableRow"><a href="admin_update_check.asp<% = strQsSID1 %>">Check for updates</a><br />
Check for updates to the Forum Software</td>
 </tr>
 <tr>
  <td class="tableRow"><a href="https://www.webwiz.net/license/" target="_blank">License Agreement</a><br />
   View a copy of the latest License Agreement for this software. </td>
 </tr><%

If blnLCode Then

%>
 <tr>
  <td class="tableRow"><a href="http://www.webwizforums.com" target="_blank">About</a><br />
   About Web Wiz Web Wiz Forums.</td>
 </tr><%

End If

%>
</table>
<%

End If




'If they want security menu
If Request.QueryString("C") = "security" Then

%>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Permissions</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_group_permissions_form.asp<% = strQsSID1 %>">Set Group Permissions</a><br />
Setup or change Group Permissions </td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_find_user.asp<% = strQsSID1 %>">Set Individual Member Permissions</a><br />
From this option you can configure permissions  for individual Members.</td>
 </tr>
</table>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Security and Anti-SPAM </td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_close_forums.asp<% = strQsSID1 %>">Security Lock Down </a><br />
From this option you can Lock the Forums so that no-one can post, register, login. etc. on the forum. Useful if you forum comes under attack or you just need to close it for maintenance. </td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_spam_filter_configure.asp<% = strQsSID1 %>">SPAM Filter</a><br />
   Configure the SPAM Filter with SPAM that you want to filter from your forum and the action to take if SPAM is detected.</td>
 </tr>
  <tr class="tableRow">
  <td><a href="admin_stopforumspam.asp<% = strQsSID1 %>">StopForumSpam Settings</a><br />
   StopForumSpam is a free Database List of Spammers that persist in abusing forums and blogs.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_bad_word_filter_configure.asp<% = strQsSID1 %>">Swear/Bad Word Filter</a><br /> 
   Configure the Swear Word Filter to block words that you feel are not appropriate to your forum.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_spam_configure.asp<% = strQsSID1 %>">Anti-SPAM Flood Control  Configuration</a><br />
Configure Anti-SPAM Flood Control settings so you don't get members spamming the forum with 1,000's of unwanted or abusive posts.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_ip_blocking.asp<% = strQsSID1 %>">IP Address Banning</a><br />
   Ban IP addresses and ranges. Anyone in a blacklisted ranges will only have limited capabilities within the forum system. </td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_email_domain_blocking.asp<% = strQsSID1 %>">Email Address Banning</a><br />
Use this option to block Email address and Email Domains from being registered on the board.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_registration_country_blocking.asp<% = strQsSID1 %>">Registration Country Block List</a><br />
Use this option to block countries that you do not permit for new registrations.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_suspend_registration.asp<% = strQsSID1 %>">Suspend New Registrations</a><br />
From this option you can Suspend New Users from Registering to use the forum.</td>
 </tr><%
 	
 	'If super admin let 'em change password
 	If lngLoggedInUserID <> 1 AND blnDemoMode = False Then
%>
 <tr class="tableRow">
  <td><a href="admin_change_admin_username.asp<% = strQsSID1 %>">Change Super Admin Username and Password</a><br />
This is highly recommended for higher security to prevent unauthorised persons access this Admin Control Panel.</td>
 </tr><%

	End If
%>
</table>
<%

End If

'If they want tools menu
If Request.QueryString("C") = "tools" Then

%>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Forum Tools </td>
 </tr>
 <%
       
'If this is an access database show the compact and repair feature
If strDatabaseType = "Access" AND strAccessDriver = "Microsoft.Jet.OLEDB.4.0" Then 

%>
 <tr class="tableRow">
  <td><a href="admin_compact_access_db.asp<% = strQsSID1 %>">Compact and Repair Access Database</a><br />
   Form here you can compact and repair your Forums Access database to increase performance.</td>
 </tr>
 <%
  
End If

%>
 <tr class="tableRow">
  <td><a href="admin_resync_forum_post_count.asp<% = strQsSID1 %>">Re-sync Topic and Post Count</a><br />
   Re-sync the Topic and Post Count display for the forum's</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_archive_topics_form.asp<% = strQsSID1 %>">Batch Lock Old Topics</a><br />
   Batch lock old Topics allows you to batch lock Topics that haven't been posted in for sometime.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_batch_delete_posts_form.asp<% = strQsSID1 %>">Batch Delete Topics</a><br />
   Clean out the Forum Database by batch deleting topics that have not been posted in for sometime.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_batch_move_topics_form.asp<% = strQsSID1 %>">Batch Move Topics</a><br />
   Batch move Topics from one forum to another.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_batch_delete_pm_form.asp<% = strQsSID1 %>">Batch Delete Private Messages</a><br />
   Clean out the Forum Database by batch deleting old Private Messages.</td>
 </tr>
 <tr class="tableRow">
  <td><a href="admin_batch_delete_members_form.asp<% = strQsSID1 %>">Batch Delete Members</a><br />
   Clean out the Forum Database by batch deleting Members who have never posted.</td>
 </tr>
</table>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Import Tool</td>
 </tr>
 <tr>
  <td class="tableRow"><a href="admin_db_import_members_form.asp<% = strQsSID1 %>">Import Members from External Database</a><br />
   This tool allows you to import members from an external database, such as another forum system, CMS, etc. </td>
 </tr>
</table><%

End If


If Request.QueryString("C") = "upgrades" Then
	
%>
<br />
<table width="100%" border="0" align="center" cellpadding="4" cellspacing="1" class="tableBorder">
 <tr>
  <td class="tableLedger">Upgrades</td>
 </tr>
 <tr>
  <td class="tableRow"><a href="admin_server_test.asp<% = strQsSID1 %>">Server Requirements Test </a><br />
Check that the server your site is hosted on and your web host have the correct requirements to run Web Wiz Forums. </td>
 </tr>
 <tr>
  <td class="tableRow"><a href="admin_update_check.asp<% = strQsSID1 %>">Check for updates</a><br />
Check for updates to the Forum Software</td>
 </tr>
 </tr>
</table>
<%

End If


%>
<!-- #include file="includes/admin_footer_inc.asp" -->

