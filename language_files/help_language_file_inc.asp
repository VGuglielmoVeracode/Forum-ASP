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


'Global
'---------------------------------------------------------------------------------
Const strTxtForumHelp = "Forum Help"
Const strTxtChooseAHelpTopic = "Choose a help topic"
Const strTxtLoginAndRegistration = "Registration and Logging into the Forum"
Const strTxtUserPreferencesAndForumSettings = "User Preferences and Forum Settings"
Const strTxtPostingIssues = "Posting Issues"
Const strTxtMessageFormatting = "Message Formatting"
Const strTxtUsergroups = "User Groups"
Const strTxtPrivateMessaging = "Private Messaging"

Const strTxtWhyCantILogin = "Why can't I login?"
Const strTxtDoINeedToRegister = "Do I need to register?"
Const strTxtLostPasswords = "Lost Passwords"
Const strTxtIRegisteredInThePastButCantLogin = "I Registered in the past but can't login"


Const strTxtInformationOnCookiesSetByThisForum = "Information on Cookies set by this Forum"
Const strTxtWhatAreCookies = "What are Cookies?"


Const strTxtHowDoIChangeMyForumSettings = "How do I change my forum settings?"
Const strTxtForumTimesAndDates = "Forum Times and Dates are not set to my local time"
Const strTxtWhatDoesMyRankIndicate = "What does my rank indicate?"
Const strTxtCanIChangeMyRank = "Can I change my rank?"
Const strTxtWhatWebBrowserCanIUseForThisForum = "What web browser can I use for this forum?"

Const strTxtHowPostMessageInTheForum = "How do I post a message in the forum?"
Const strTxtHowDeletePosts = "How do I delete posts?"
Const strTxtHowEditPosts = "How do I  edit posts?"
Const strTxtHowSignaturToMyPost = "How do I add a signature to my post?"
Const strTxtHowCreatePoll = "How do I create a poll?"
Const strTxtWhyNotViewForum = "Why can I not view a forum?"
Const strTxtMyPostIsHiddenOrPendingApproval = "My post is not displayed, is 'Hidden', or 'Pending Approval'"
Const strTxtInternetExplorerWYSIWYGPosting = "Rich Text Editor (WYSIWYG) posting issues"

Const strTxtWhatForumCodes = "What are BBcodes?"
Const strTxtCanIUseHTML = "Can I use HTML?"
Const strTxtWhatEmoticons = "What are Emoticons (Smilies)"
Const strTxtCanPostImages = "Can I post images?"
Const strTxtWhatClosedTopics = "What are closed topics?"

Const strTxtWhatForumAdministrators = "What are Forum Administrators?"
Const strTxtWhatForumModerators = "What are Forum Moderators?"
Const strTxtWhatUsergroups = "What are User Groups?"

Const strTxtWhatIsPrivateMessaging = "What is Private Messaging?"
Const strTxtIPrivateMessages = "I cannot send Private Messages"
Const strTxtIPrivateMessagesToSomeUsers = "I cannot send Private Messages to some users"
Const strTxtHowCanPreventSendingPrivateMessages = "How can I prevent someone from sending me Private Messages"

Const strTxtRSSFeeds = "RSS Feeds"
Const strTxtHowDoISubscribeToRSSFeeds = "How Do I Subscribe to the Forums RSS Feeds?"

Const strTxtCalendarSystem = "Calendar System"
Const strTxtWhatIsTheCalendarSystem = "What is the Calendar System?"
Const strTxtHowDoICreateCalendarEvent = "How do I create a Calendar Event?"

Const strTxtAbout = "About"
Const strTxtWhatSoftwareIsUsedForThisForum = "What software is used for this forum?"


Const strTxtFAQ1 = "To login into the Forum you must use the Username and Password that you entered when registering for the Forum. If you have not yet registered then you must first do so in order to login. If you have registered and still are unable to login then first check that you have all cookies enabled on your web browser, you may need to add this web site to your browsers list of trusted web sites. If you are banned from the Forums then this may prevent you form logging in, in which case check with the Forums administrator."
Const strTxtFAQ2 = "You may not need to register to post in the Forums, it is up to the Forums Administrator as to whether they allow you to post in the Forums as an unregistered user. However, by registering you will be able to use additional features. It only takes a few minutes to register, so it is recommended that you do so."
Const strTxtFAQ3 = "If you have lost your password then don't panic. Although passwords can not be retrieved they can be reset. To reset your password click on the Login button and at the bottom of the login page you will have a link to the lost password page to request that a new password be emailed to you. If this option is not available or you do not have a valid email address in your profile then you need to contact the Forum Administrator or a Moderator and ask them to change your password for you."
Const strTxtFAQ4 = "It could be that you haven't posted anything for a while, or never posted anything. It is common for the Forum Administrator to periodically delete users from the database to free up usernames and reduce the size of the database."
Const strTxtFAQ5 = "You can change your forum settings, profile information, registration details, etc. from your <a href=""member_control_panel.asp"">Member Control Panel Menu</a>, once you have logged into the Forum. You will be able to control many aspects and access member features from this central menu."
Const strTxtFAQ6 = "The time used in the Forums has been set by the Forum Administrator. To change the Times and Dates to your own local settings then simply edit your 'Forum Preferences' through your <a href=""member_control_panel.asp"">Member Control Panel Menu</a>, and set by how many hours the time is offset from your own local time. The Forums are not designed to adjust between standard and daylight saving times, so you may need to adjust the time offset in your Forum Preferences during these months."
Const strTxtFAQ7 = "Ranks in the forums indicate which user group you are a member of and to identify users, for example, moderators and administrators may have a special rank. Depending on the setup of the forum you may be able to use different features of the forum depending on which rank you belong to."
Const strTxtFAQ8 = "Normally you can not, but if the forum administrator has setup ranks using the ladder system you may be able to move up groups in the forum by the number of posts you have made."
Const strTxtFAQ9 = "To post a message in the Forums click on the relevant button on the forum or topic screens. Depending on how the forum administrator has setup the forum depends if you need to login first before you can post a message. The facilities available to you in each of the forums are listed at the bottom of the topic screen."
Const strTxtFAQ10 = "Unless you are a Forum Moderator or an Administrator you can only delete your own posts and only if the forum has been setup with the relevant rights for you to be able to delete your post. If someone has replied to your post then you will no-longer be able to delete it."
Const strTxtFAQ11 = "Unless you are a Forum Moderator or an Administrator you can only edit your own posts and only if the forum administrator has created the relevant rights for you to do so. When you edit your posts depending on the forum setup, it will display the username, time, and date of when the post was edited at the bottom of the post."
Const strTxtFAQ12 = "If the forum administrator has allowed the use of signatures in the forums you can add a signature to the bottom of your posts. To do so you need to first create a signature in your 'Profile Information' through your <a href=""member_control_panel.asp"">Member Control Panel Menu</a>, once you have done this you can add your signature to the bottom of your posts by checking the 'Show Signature' checkbox at the bottom of the posting form."
Const strTxtFAQ13 = "If you have sufficient rights to create a poll in a forum you will see a 'New Poll' button at the top of the screen on the forum and topic screens. When creating a poll you need to enter a poll question and at least two options for the poll. You may also select whether people can vote multiple times or just once in the poll."
Const strTxtFAQ14 = "Some forums are setup to allow only certain users or groups of users to access them. To view, read, post, etc. in a forum you may first need permission which only a forum moderator or a forum administrator can grant you."
Const strTxtFAQ28 = "This means that you have submitted a New Post or Topic to a forum that requires the a forum admin or moderator to approve your post before it is displayed to other members or the general public. The Post is still available for the member who submitted the Post to edit or delete during this time."
Const strTxtFAQ15 = "If you are using Internet Explorer 5+ (windows only), Firefox, Opera 9+, or Netscape 7+, and if the forum administrator has enabled it you should have a Rich Text (WYSIWYG) Editor to type your messages with. If you find that you are having problems posting using the WYSIWYG editor then you can disable this WYSIWYG Editor by simply editing your profile and selecting to turn off the WYSIWYG Posting Editor."
Const strTxtFAQ16 = "Forum BBcodes allow you to format the messages you post in the forums. BBcodes are very similar to HTML except tags are enclosed in square brackets, [ and ], rather than, &lt; and &gt;. You can also disable BBcodes when posting a message. <a href=""javascript:winOpener('BBcodes.asp','codes',1,1,610,500)"">Click here to view BBcodes that are available in the forums</a>."
Const strTxtFAQ17 = "HTML cannot be used in your posts, this is done for security reasons as malicious HTML code can be used to destroy the layout of the forum or even crash a web browser when a user tries to view a post."
Const strTxtFAQ18 = "Emoticons or Smilies are small graphical images that can be used to express feelings or show emotions. If the forum administrators have allowed Emoticons in the forums then you can see them next to the posting form when posting a message. To add an emoticon to your post then simply click on the emoticon you would like to add to your post."
Const strTxtFAQ19 = "Images can be added to your posts, if the forum administrators have allowed uploading of images you can use this to upload an image from your own computer into your message. However, if image uploading is not available then you will need to link to the image that needs to be stored on a publicly accessible web server, e.g. http://www.mysite.com/my-picture.jpg."
Const strTxtFAQ20 = "Closed Topics are set this way by forum administrators or moderators. Once a Topic is closed you will no-longer be able to post a reply in that topic or vote in a poll."
Const strTxtFAQ21 = "Forums Administrators are people who have the highest level of control over the forums, they have the ability to turn on and off features on the forums, ban users, remove users, edit and delete posts, create users groups, etc."
Const strTxtFAQ22 = "Moderators are individuals or groups of users who look after the day to day running of the forums. They have the power to edit, delete, move, close, unclose, topics and posts, in the forum they moderate. Moderators generally are there to prevent people from posting offensive or abusive material."
Const strTxtFAQ23 = "User Groups are a way to group users. Each user is a member of a usergroup and each group can be assigned individual rights in forums, to read, view, post, create polls, etc."
Const strTxtFAQ24 = "There maybe several reasons for this, you are not logged in, you are not registered, or the forum administrators have disabled the Private Messaging system."
Const strTxtFAQ25 = "This maybe because the person you are trying to send a Private Message to has blocked you from being able to send them Private Messages, the persons 'Inbox' is full, or you have exceeded the number of Private Messages you can send in this hour. If any of these are case you should receive a message informing you of this if you try to send the member a Private Message."
Const strTxtFAQ26 = "If you do find that you are getting unwanted Private Messages from a user you can block them from sending you Private Messages. To do this go into the Private Messaging system and go to your buddy list. Go to add the user as a buddy, but instead choose the option from the drop down list 'Not to message you', this will prevent this user from sending you anymore Private Messages."
Const strTxtFAQ27 = "If the forum administrator has enabled Private Messaging you can use the forums built in Private Messaging System to send Private Messages to other forum members. The system works very much like email, but within the Forum System."
Const strTxtFAQ29 = "RSS, which stands for Really Simple Syndication, is a technology using XML that allows web sites to syndicate web site content. If the forum administer has enabled RSS Feeds you can subscribe to any of the RSS Feeds available in the forum so you are always up-to-date with the latest Posts or Calendar Events."
Const strTxtFAQ30 = "There are many ways you can subscribe to one or all of the RSS Feed's available within the forum, by either; using a Web Browser that supports Live RSS Bookmarks (Firefox, IE7, Safari, Opera), using an online RSS News Reader, using Windows Vista Sidebar, installing an RSS News Reader on your computer, or using an email program that supports RSS like Mozilla Thunderbird."
Const strTxtFAQ31 = "The Calendar System is a built in Calendar System that works very much like a normal calendar or diary, with Week View, Month View, and Year View. You can select what Week, Month, or Year, you wish to view and see Calendar Events and member birthdays for your selected date."
Const strTxtFAQ32 = "If the forum administrator has granted a member or group permission to create a Calendar Event you can select an 'Event Date' when creating a New Topic. This Event can then be viewed in like any other Forum Topic or through the Calendar System where it is displayed under the date the event occurs on."
Const strTxtFAQ33 = "The software used to run this forum is called, Web Wiz Forums. Web Wiz Forums is a Bulletin Board System written using Microsoft's Active Server Pages (ASP) and designed to run on Windows based web servers. Web Wiz Forums is simple to install on your own web site and is available for free from <a href=""http://www.webwizforums.com"">www.webwizforums.com</a>."
Const strTxtFAQ34 = "You can use this forum with almost any web browser that supports XHTML and JavaScript. For full functionality we recommend that you use Internet Explorer 7+, Firefox, Opera 9+, and Netscape 9+.  Safari and versions of Internet Explorer prior to version 7, will run but with reduced functionality."
Const strTxtFAQ36 = "Cookies are small files which are stored on a user's computer. They are designed to hold a modest amount of data specific to a particular client and website, and can be accessed either by the web server or the client computer. This allows members to login to the forum and display pages tailored to the forum member for items such as new Posts since the members last visit."

Const strTxtFAQ35 = "The Forum requires the use of cookies in order for forum services provided to Forum Members to function correctly. The following Cookies are set by the Forum."
Const strTxtFAQ35sID = "sID - This is a temporary cookie used to associate your browser with your forum session"
Const strTxtFAQ35sLID = "sLID - This cookie holds Members Login settings and is also used for auto login to the forum"
Const strTxtFAQ35lVisit = "lVisit - This cookie holds the date and time of your last visit to the forum in order to notify which Topics and Posts are new since your last visit."
Const strTxtFAQ35fID = "fID - This cookie is used for Private Forums that require a password in order to login."
Const strTxtFAQ35pID = "pID - This is a security cookie set when voting in Forum Polls."
Const strTxtFAQ35MobileView = "MobileView - This Cookie is set if you turn off the Mobile View when visiting the forum using a Mobile Browser."

%>