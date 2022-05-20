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






Dim strLoggedInUsername		'Holds a logged in users username
Dim intGroupID			'Holds the group ID number the member is a group of
Dim lngLoggedInUserID		'Holds a logged in users ID number
Dim blnActiveMember		'Set to false if the member is no longer allowed to post messages on the forum
Dim strDateFormat		'Holds the users date format
Dim strTimeOffSet		'Holds the users time offset in + or -
Dim intTimeOffSet		'Holds the users time offset
Dim blnReplyNotify		'Set to true if the user wants to be notified of replies to posts
Dim blnAttachSignature		'Set to true if the user always wants there signature attached
Dim blnWYSIWYGEditor		'Set to true if the user wants to use the IE WYSIWYG post editor
Dim strLoggedInUserEmail	'Holds the logged in users email address
Dim intNoOfPms			'Holds the number of PM's the user has
Dim dtmUserLastVisitDate	'Holds the last visit date of the user
Dim blnLoggedInUserSignature	'set to true if the user has enetered a signature
Dim blnLoggedInUserEmail	'Set to true if the user has entered there e-mail
Dim blnBanned			'Holds if the user is banned (suspended)
Dim blnGuest			'set to true for the Guest account (Group ID 2)
Dim dtmLastVisitDate		'Holds the last visit date for the user


Dim sarySessionData		'Holds the users session data
Dim strWebsiteName		'Holds the website name
Dim strMainForumName		'Holds the forum name
Dim strForumPath		'Holds the virtual path to the forum
Dim strForumEmailAddress	'Holds the forum e-mail address
Dim blnRTEEditor		'Set to true if the Rich Text Editor(RTE) is enabled
Dim blnEmail			'Boolean set to true if e-mail is on
Dim strMailComponent		'Email coponent the forum useses
Dim strMailServer		'Forums incomming mail server
Dim strLoggedInUserCode		'Holds the user code of the user
Dim blnLCode			'set to true
Dim blnAdmin			'set to true if the user is a forum admininstrator (Group ID 1)
Dim blnModerator		'Set to true if the user is a forum moderator
Dim intTopicPerPage		'Holds the number of topics to show on each page
Dim strTitleImage		'Holds the path and name for the title image for the forum
Dim blnEmoticons		'Set to true if emoticons are turned on
Dim blnGuestPost		'Set to true if guests can post
Dim blnAvatar			'Set to true if the forum can use avatars
Dim blnEmailActivation		'Set to true if the e-mail activation is turned on
Dim blnSendPost			'Set to true if post is sent with e-mail notification
Dim intNumHotViews		'Holds the number of how many views a topic has before it becomes a hot topic
Dim intNumHotReplies		'Holds the number of replies before a topic becomes a hot topic
Dim blnPrivateMessages		'Set to true if private messages are allowed
Dim intPmInbox			'Holds the number of private messages allowed by each user
Dim intThreadsPerPage		'Holds the number of threads shown per page
Dim intSpamTimeLimitSeconds	'Holds the number of secounds between posts
Dim intSpamTimeLimitMinutes	'Holds the number of minutes the user can post five posts in
Dim intMaxPollChoices		'Holds the maximum allowed number of poll choices
Dim blnEmailMessenger		'Set to true if the email system is on
Dim blnActiveUsers		'Set to true if active users is enabled
Dim blnForumClosed		'Set to true of the forum is cloded for maintence
Dim blnShowEditUser		'Set to true if we are to show the username and time a post is edited
Dim blnShowProcessTime		'Set to true if we are to show how long the page took to be processed on the server
Dim dblStartTime		'Holds the start time for the page process
Dim blnDisplayForumClosed	'Set to true if we are looking at the closed forum page
Dim blnFlashFiles		'Set to true if Flash support is enabled
Dim strWebsiteURL 		'Holds the URL to the sites homepage
Dim blnShowMod			'Set to true if mod groups are shown on the main forum page
Dim blnAvatarUploadEnabled	'Set to true if avatars are enabled
Dim blnRegistrationSuspeneded	'Set to true if new registrations are suspended
Dim strImageTypes		'Holds the types of images allowed in the forum
Dim blnLongRegForm		'Set to true if the reg form is to be the long version
Dim blnCAPTCHAsecurityImages	'Set to true if CAPTCHA is required for Guest Posting
Dim strPathToRTEFiles		'Holds the path to the RTE files
Dim saryPermissions		'Holds the array for permissions
Dim blnTopicIcon		'Holds if message icons are enabled
Dim strMailServerUser		'Holds the SMTP server username
Dim strMailServerPass		'Holds the SMTP server password
Dim strNavSpacer		'Navigation spacer
Dim strImagePath		'Image path
Dim strCSSfile			'Path and name of CSS skin file
Dim saryActiveUsers		'Holds the active users array
Dim strQsSID			'Holds the users session ID displayed in pages
Dim strSessionData		'Holds the users session data
Dim strSessionID		'Holds  the forums internal sesison ID
Dim strQsSID1			'Holds the session ID for ? URL links
Dim strQsSID2			'Holds the session ID for &amp; URL links
Dim strQsSID3			'Holds the session ID for & URL links in redirects
Dim blnCalendar			'Set to true if Calendar is enabled
Dim blnGuestSessions		'Set to true if guest sessions are enabled
Dim blnMemberApprove		'Set to true if new members need to be approved by forum admin
Dim intLoginAttempts		'Holds the number of login attempts by the user
Dim blnRSS			'Set to true if RSS feed is enabled
Dim strInstallID		'Holds the forums unique install ID
Dim intPmFlood			'PM Flood control amount
Dim blnACode			'Set to true
Dim blnCAPTCHAabout		'CAPTCHA about
Dim sarryUnReadPosts		'Array to hold un-read post ID's
Dim sarryUnReadComments		'Array to hold uread comments
Dim strBreadCrumbTrail		'New variable that holds the bread crumb trail for the page
Dim strStatusBarTools		'As is says on the tin (or variable in this case)
Dim intForumID			'Holds the ID number of the forum
Dim strLinkPage			'Holds the page to link to
Dim strLinkPageTitle		'Holds the page to link to's title
Dim strLinkPageSelectID		'Holds the page to link select form ID
Dim strUploadComponent		'Holds upload component
Dim strUploadFileTypes		'Holds upload file types
Dim lngUploadMaxImageSize	'Holds max upload image size
Dim lngUploadMaxFileSize	'Holds upload max file size
Dim intUploadAllocatedSpace	'Holds upload allocated user space
Dim strUploadOriginalFilePath	'For security the users upload folder is set in the common.asp on each page, but for admin purposes may need to get the orginal upload path
Dim blnWebWizNewsPad		'Set to true if Web Wiz NewsPad integration is enabled
Dim strWebWizNewsPadURL		'URL to Web Wiz NewsPad
Dim strClientBrowserVersion	'Holds the browser version
Dim strForumImageType		'Holds the type of image to use (PNG or GIF)
Dim saryConfiguration		'Holds the configutaion/settings data
Dim strAvatarTypes		'Holds the avatar types
Dim intMaxAvatarSize		'Holds the max avatar size
Dim lngMostEverActiveUsers	'Holds the number of the most ever active users
Dim dtmMostEverActiveDate	'Holds the date of the most ever active users
Dim lngMailServerPort		'Holds the port number of the mail server
Dim strPageEncoding		'Page encoding
Dim strTextDirection		'Writing direction for lanaguge
Dim strCookiePrefix		'Cookie prefix
Dim strCookiePath		'Cookie path
Dim blnDatabaseHeldSessions	'Holds session data in database
Dim blnNewUserCode		'User code update settings
Dim blnModeratorProfileEdit	'Set to true it moderatrs can edit member profiles
Dim blnForumViewing		'Shows how many people are viewing forums
Dim blnDetailedErrorReporting	'Shows detailed error messages
Dim blnDisplayBirthdays		'Shows user birthdays within the calendar system
Dim blnDisplayTodaysBirthdays	'Shows todays birthdays on the home page
Dim intRssTimeToLive		'The RSS Time to Live value
Dim intRSSmaxResults		'Max results to return for RSS Feeds
Dim strDefaultPostOrder		'The default order to display posts in
Dim strCookieDomain		'Holds the cookie domain if set
Dim blnYouTube			'YouTube Content
Dim strPMoverAction		'What to do if the member has exceeded their private messages
Dim blnPmFlashFiles		'PM Falsh files
Dim blnPmYouTube		'PM YouTube
Dim intMinPasswordLength	'Minimum Password length
Dim intMinUsernameLength	'Minimum Username length
Dim blnQuickReplyForm		'Set to true if quick reply is enabled
Dim intIncorrectLoginAttempts	'This is the number of incorrect login attempts before CAPTCH is disaplyed
Dim intPointsReply		'The number of points a user gets for repling
Dim intPointsTopic		'The number of points a user gets for creating a new topic
Dim intEditedTimeDelay		'This is the amount of time before the edited notes are added to posts
Dim blnEnforceComplexPasswords	'Set to true if complex passwords is enabled
Dim blnRealNameReq		'If real name required
Dim blnLocationReq		'If location is required
Dim blnTopicRating		'If topic rating is permitted
Dim blnBoldNewTopics		'Bold topics
Dim blnBoldToday		'Bold today date
Dim strBoardMetaDescription	'Meta description
Dim strBoardMetaKeywords	'Meta keywords
Dim blnDynamicMetaTags		'Dynamic meta tags for topics and forums
Dim blnSearchEngineSessions	'Search engine sessions
Dim blnNoFollowTagInLinks	'Sets nofollow on users links
Dim blnSignatures		'If signatures are allowed
Dim intPmOutbox			'Number of PM's in outbox
Dim blnWindowsAuthentication	'Set to true if windows autrhntication is enabled
Dim strEmailErrorMessage	'Holds the error message if an error sending an email
Dim blnHomePage			'If users are allowd to post homepages
Dim blnSeoTitleQueryStrings	'If SEO Titles are added to querystrings
Dim strOSType			'Holds the OS type
Dim intEditPostTimeFrame	'Holds the time within that members are allowed to edit posts
Dim blnChatRoom			'Set to true if chat room enabled
Dim blnPmIgnoreSpamFilter	'Set to true if spam filter is ignored in private messages
Dim blnMobileBrowser		'Set to true if a mobile browser
Dim blnMobileClassicView	'Set to true if mobile browser is in classic view
Dim blnMobileView		'Set to tyrue to enable mobile view
Dim strRegistrationRules	'Holds the rules for registration
Dim strCustRegItemName1		'Holds the name of the custom registration item
Dim blnReqCustRegItemName1	'If cust item is required
Dim blnViewCustRegItemName1	'If members can view custom item in profile
Dim strCustRegItemName2		'Holds the name of the custom registration item
Dim blnReqCustRegItemName2	'If cust item is required
Dim blnViewCustRegItemName2	'If members can view custom item in profile
Dim strCustRegItemName3		'Holds the name of the custom registration item
Dim blnReqCustRegItemName3	'If cust item is required
Dim blnViewCustRegItemName3	'If members can view custom item in profile
Dim strStatsTrackingCode	'Holds stats code (eg. Google Analytics)
Dim strForumHeaderAd		'Holds ad code
Dim strForumPostAd		'Holds ad code
Dim strVigLinkKey		'Holds VigLink Code	
Dim blnShowLatestPosts		'Set to true if showing latest posts
Dim strAnswerPosts		'Holds the status of Answer posts
Dim blnShareTopicLinks		'Display shared topic links
Dim intPointsAnswered		'The number of points for an anwser
Dim intPointsThanked		'The number of posts for being thanked
Dim blnPostThanks		'Set to true if members can be thanked for posts
Dim blnTwitterTweet		'Set to true if displaying twitter tweet buttons
Dim intNoOfInboxPms		'Hold stotal number of PM's in a users inbox
Dim strAnswerPostsWording	'The wording for anwser posts
Dim blnFacebookLike		'Set to true if displaying facebook like buttons
Dim strFacebookPageID		'Holds the facebook page ID
Dim strFacebookImage		'Holds the image to use for Facebook sharing
Dim blnHttpXmlApi		'Set to true if HTTP XML API is enabled
Dim blnUploadSecurityCheck	'Set to true if scanning uploads for malcious code
Dim blnShowHeaderFooter		'Set to true if showing custom header/footer
Dim strHeader			'Holds forum custom header
Dim strFooter			'Holds forum custom footer
Dim blnShowMobileHeaderFooter	'Set to true if showing custom mobile header/footer
Dim strHeaderMobile		'Holds forum custom header for mobile
Dim strFooterMobile		'Holds forum custom footer for mobiles
Dim blnEmailNotificationSendAll	'Set to true if sending all post notifications
Dim intMaxImageWidth 		'Max width of images
Dim intMaxImageHeight 		'Max height of images
Dim blnDisplayForumStats	'Set to true if displaying forum stats on homepage
Dim blnDisplayMemberList	'Set to true if displaying member list
Dim strForumsMessage		'Message that appears across all forums
Dim blnUrlRewrite		'Set to true if URL Rewriting is enabled
Dim blnGooglePlusOne		'Set to true for Google Plus1 buttons
Dim blnModViewIpAddresses
Dim intSearchTimeDefault	'Holds the defualt search date for searhes
Dim strMemberProfileView	'Holds who can view member profiles
Dim blnStopForumSpam
Dim strStopForumSpamApiKey
Dim lngStopForumSpamBlocked
Dim blnStopForumSpamUsername
Dim strSslEnabled
Dim blnSslEnabledPage
Dim strRegistrationRedirect
Dim blnEuCookieLaw
Dim blnLocalLinksSameWindow
Dim blnRegistrationCAPTCHA
Dim strCountryBlockRegList
Dim intSessionTimeout
Dim blnSessionIpCheck
Dim blnUploadDuplicateFiles
Dim blnCovertTextLinkstoHyperLink
Dim blnHttpsWebWizApi
Dim strRegPrivacyNotice



'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
Const strCAPTCHAversion = "4.04 wwf"
Const blnDemoMode = False
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******   





'Forum Permissions (Forum Level)
Dim blnRead
Dim blnPost
Dim blnReply
Dim blnEdit
Dim blnDelete
Dim blnPriority
Dim blnPollCreate
Dim blnVote
Dim blnCheckFirst
Dim blnEvents

'Group Permissions (Group Level)
Dim blnAttachments
Dim blnImageUpload
Dim blnGroupSignatures
Dim blnGroupURLs
Dim blnGroupImages


'Initialise variables
strLoggedInUsername = strTxtGuest
blnActiveMember = true
blnBanned = false
blnLoggedInUserEmail = false
blnLoggedInUserSignature = false
intGroupID = 2
lngLoggedInUserID = 2
blnAdmin = false
blnModerator = false
blnGuest = true
intTimeOffSet = 0
strTimeOffSet = "+"
blnWYSIWYGEditor = True
strPathToRTEFiles = ""
dblStartTime = Timer()
If Request.QueryString("03333408000") Then Server.Execute("includes/egg_inc.asp")
blnMobileClassicView = False
blnSslEnabledPage = True '- Change November 2016 to make all pages SSL Enabled when SSL is stwiched on from the admin area


'Force Global Veriable Reload
If Request.QueryString("reload") Then
	Application.Lock
	Application(strAppPrefix & "blnConfigurationSet") = false
	Application(strAppPrefix & "saryAppDateTimeFormatData") = false
	Application.UnLock
End If




'******************************************
'***    Read in Configuration Data     ****
'******************************************

'This function when called with load in the global variables at application level 
'(application level variables are used for improved performance and less database hits)
Public Sub getForumConfigurationData()

	'Read in the Forum configuration
	If isEmpty(Application(strAppPrefix & "blnConfigurationSet")) OR isNull(Application(strAppPrefix & "blnConfigurationSet")) OR Application(strAppPrefix & "blnConfigurationSet") = "" OR NOT Application(strAppPrefix & "blnConfigurationSet") = True Then

		'Initialise the SQL variable with an SQL statement to get the configuration details from the database
		strSQL = "SELECT " & strDbTable & "SetupOptions.Option_Item, " & strDbTable & "SetupOptions.Option_Value " & _
		"FROM " & strDbTable & "SetupOptions" &  strDBNoLock & ";"
		
		'Set error trapping
		On Error Resume Next
	
		'Query the database
		rsCommon.Open strSQL, adoCon
		
		'If an error has occurred write an error to the page
		If Err.Number <> 0 Then	Call errorMsg("An error initialising Global Variables from database.", "getForumConfigurationData()_read_global_variables", "global_variables_inc.asp")
		
		'Disable error trapping
		On Error goto 0			
		
		'If there is config deatils in the recordset then read them in
		If NOT rsCommon.EOF Then
			
			'Place into an array for performance
			saryConfiguration = rsCommon.GetRows()
			
			'Clean up
			rsCommon.Close
		
			'SQL Query Array Look Up table
			'0 = tblSetupOptions.Option_Item
			'1 = tblSetupOptions.Option_Value
			
		
			'Read in the Configuration Settings
			blnACode = CBool(getConfigurationItem("A_code", "bool"))
			blnActiveUsers = CBool(getConfigurationItem("Active_users", "bool"))
			blnAvatar = CBool(getConfigurationItem("Avatar", "bool"))
			blnCalendar = CBool(getConfigurationItem("Calendar", "bool"))
			blnCAPTCHAsecurityImages = CBool(getConfigurationItem("CAPTCHA", "bool"))
			blnEmailActivation = CBool(getConfigurationItem("Email_activate", "bool"))
			blnEmail = CBool(getConfigurationItem("email_notify", "bool"))
			blnSendPost = CBool(getConfigurationItem("Email_post", "bool"))
			blnEmailMessenger = CBool(getConfigurationItem("Email_sys", "bool"))
			blnEmoticons = CBool(getConfigurationItem("Emoticons", "bool"))
			blnFlashFiles = CBool(getConfigurationItem("Flash", "bool"))
			strForumEmailAddress = getConfigurationItem("forum_email_address", "string")
			strMainForumName = getConfigurationItem("forum_name", "string")
			strForumPath = getConfigurationItem("forum_path", "string")
			blnForumClosed = CBool(getConfigurationItem("Forums_closed", "bool"))
			blnGuestSessions = CBool(getConfigurationItem("Guest_SID", "bool"))
			intNumHotReplies = CInt(getConfigurationItem("Hot_replies", "numeric"))
			intNumHotViews = CInt(getConfigurationItem("Hot_views", "numeric"))
			blnRTEEditor = CBool(getConfigurationItem("IE_editor", "bool"))
			strInstallID = getConfigurationItem("Install_ID", "string")
			'blnLCode = CBool(getConfigurationItem("L_code", "bool"))
			blnLongRegForm = CBool(getConfigurationItem("Long_reg", "bool"))
			strMailComponent = getConfigurationItem("mail_component", "string")
			strMailServerPass = getConfigurationItem("Mail_password", "string")
			strMailServer = getConfigurationItem("mail_server", "string")
			strMailServerUser = getConfigurationItem("Mail_username", "string")
			blnMemberApprove = CBool(getConfigurationItem("Member_approve", "bool"))
			blnWebWizNewsPad = CBool(getConfigurationItem("NewsPad", "bool"))
			strWebWizNewsPadURL = getConfigurationItem("NewsPad_URL", "string")
			intPmInbox = CInt(getConfigurationItem("PM_inbox", "numeric"))
			intPmOutbox = CInt(getConfigurationItem("PM_outbox", "numeric"))
			intPmFlood = CInt(getConfigurationItem("PM_Flood", "numeric"))
			blnPrivateMessages = CBool(getConfigurationItem("Private_msg", "bool"))
			blnShowProcessTime = CBool(getConfigurationItem("Process_time", "bool"))
			blnRegistrationSuspeneded = CBool(getConfigurationItem("Reg_closed", "bool"))
			blnRSS = CBool(getConfigurationItem("RSS", "bool"))	
			blnShowEditUser = CBool(getConfigurationItem("Show_edit", "bool"))
			blnShowMod = CBool(getConfigurationItem("Show_mod", "bool"))
			strCSSfile = getConfigurationItem("Skin_file", "string")
			strImagePath = getConfigurationItem("Skin_image_path", "string")
			strNavSpacer = getConfigurationItem("Skin_nav_spacer", "string")
			intSpamTimeLimitMinutes = CInt(getConfigurationItem("Spam_minutes", "numeric"))
			intSpamTimeLimitSeconds = CInt(getConfigurationItem("Spam_seconds", "numeric"))
			intThreadsPerPage = CInt(getConfigurationItem("Threads_per_page", "numeric"))
			strTitleImage = getConfigurationItem("Title_image", "string")
			blnTopicIcon = CBool(getConfigurationItem("Topic_icon", "bool"))
			intTopicPerPage = CInt(getConfigurationItem("Topics_per_page", "numeric"))
			intUploadAllocatedSpace = CLng(getConfigurationItem("Upload_allocation", "numeric"))
			blnAvatarUploadEnabled = CBool(getConfigurationItem("Upload_avatar", "bool"))
			intMaxAvatarSize = CInt(getConfigurationItem("Upload_avatar_size", "numeric"))
			intMaxImageWidth = CInt(getConfigurationItem("Upload_img_width", "numeric"))
			intMaxImageHeight = CInt(getConfigurationItem("Upload_img_height", "numeric"))
			strAvatarTypes = getConfigurationItem("Upload_avatar_types", "string")
			strUploadComponent = getConfigurationItem("Upload_component", "string")
			lngUploadMaxFileSize = CLng(getConfigurationItem("Upload_files_size", "numeric"))
			strUploadFileTypes = getConfigurationItem("Upload_files_type", "string")
			lngUploadMaxImageSize = CLng(getConfigurationItem("Upload_img_size", "numeric"))
			strImageTypes = getConfigurationItem("Upload_img_types", "string")
			intMaxPollChoices = CInt(getConfigurationItem("Vote_choices", "numeric"))
			strWebsiteName = getConfigurationItem("website_name", "string") 
			strWebsiteURL = getConfigurationItem("website_path", "string")
			strCountryBlockRegList = getConfigurationItem("country_reg_block_list", "string")
			
			

			lngMostEverActiveUsers = CLng(getConfigurationItem("Most_active_users", "numeric"))
			dtmMostEverActiveDate = CDate(getConfigurationItem("Most_active_date", "date"))
			lngMailServerPort = CInt(getConfigurationItem("Mail_server_port", "numeric"))
			strPageEncoding = getConfigurationItem("Page_encoding", "string")
			strTextDirection = getConfigurationItem("Text_direction", "string")
			strCookiePrefix = getConfigurationItem("Cookie_prefix", "string")
			strCookiePath = getConfigurationItem("Cookie_path", "string")
			blnDatabaseHeldSessions =  CBool(getConfigurationItem("Session_db", "bool"))
			
			
			blnNewUserCode = CBool(getConfigurationItem("Tracking_code_update", "bool"))
			blnModeratorProfileEdit = CBool(getConfigurationItem("Mod_profile_edit", "bool"))
			blnForumViewing = CBool(getConfigurationItem("Active_users_viewing", "bool"))
			blnDetailedErrorReporting = CBool(getConfigurationItem("Detailed_error_reporting", "bool"))
			blnDisplayBirthdays = CBool(getConfigurationItem("Show_birthdays", "bool"))
			blnDisplayTodaysBirthdays = CBool(getConfigurationItem("Show_todays_birthdays", "bool"))
			intRssTimeToLive = CInt(getConfigurationItem("RSS_TTL", "numeric"))
			intRSSmaxResults = CInt(getConfigurationItem("RSS_max_results", "numeric"))
			strDefaultPostOrder = getConfigurationItem("Post_order", "string")
			strCookieDomain = getConfigurationItem("Cookie_domain", "string")
			blnYouTube = CBool(getConfigurationItem("YouTube", "bool"))
			strPMoverAction = getConfigurationItem("PM_overusage_action", "string")
			blnPmFlashFiles = CBool(getConfigurationItem("PM_Flash", "bool"))
			blnPmYouTube = CBool(getConfigurationItem("PM_YouTube", "bool"))
			intMinPasswordLength = CInt(getConfigurationItem("Min_password_length", "numeric"))
			intMinUsernameLength = CInt(getConfigurationItem("Min_usename_length", "numeric"))
			blnQuickReplyForm = CBool(getConfigurationItem("Quick_reply", "bool"))
			intIncorrectLoginAttempts = CInt(getConfigurationItem("Login_attempts", "numeric"))
			intPointsReply = CInt(getConfigurationItem("Points_reply", "numeric"))
			intPointsTopic = CInt(getConfigurationItem("Points_topic", "numeric"))
			intPointsAnswered = CInt(getConfigurationItem("Points_answer", "numeric"))
			intPointsThanked = CInt(getConfigurationItem("Points_thanked", "numeric"))
			intEditedTimeDelay = CInt(getConfigurationItem("Edited_by_delay", "numeric"))
			blnEnforceComplexPasswords = CBool(getConfigurationItem("Password_complexity", "bool"))
			blnRealNameReq = CBool(getConfigurationItem("Real_name", "bool"))
			blnLocationReq = CBool(getConfigurationItem("Location", "bool"))
			blnTopicRating = CBool(getConfigurationItem("Topic_rating", "bool"))
			blnBoldNewTopics = CBool(getConfigurationItem("Topics_new_bold", "bool"))
			blnBoldToday = CBool(getConfigurationItem("Date_today_bold", "bool"))
			strBoardMetaDescription = getConfigurationItem("Meta_description", "string")
			strBoardMetaKeywords = getConfigurationItem("Meta_keywords", "string")
			blnDynamicMetaTags = CBool(getConfigurationItem("Meta_tags_dynamic", "bool"))
			blnSearchEngineSessions = CBool(getConfigurationItem("Search_eng_sessions", "bool"))
			blnNoFollowTagInLinks = CBool(getConfigurationItem("Hyperlinks_nofollow", "bool"))
			blnSignatures = CBool(getConfigurationItem("Signatures", "bool"))
			blnHomePage = CBool(getConfigurationItem("Homepage", "bool"))
			blnSeoTitleQueryStrings = CBool(getConfigurationItem("SEO_title", "bool"))
			intEditPostTimeFrame = CInt(getConfigurationItem("Edit_post_time_frame", "numeric"))
			blnChatRoom = CBool(getConfigurationItem("Chat_room", "bool"))
			blnPmIgnoreSpamFilter = CBool(getConfigurationItem("PM_spam_ignore", "bool"))
			blnMobileView = CBool(getConfigurationItem("Mobile_View", "bool"))
			strRegistrationRules = getConfigurationItem("Registration_Rules", "string")
			
			strCustRegItemName1 = getConfigurationItem("Cust_item_name_1", "string")
			blnReqCustRegItemName1 = CBool(getConfigurationItem("Cust_item_name_req_1", "bool"))
			blnViewCustRegItemName1 = CBool(getConfigurationItem("Cust_item_name_view_1", "bool"))
			strCustRegItemName2 = getConfigurationItem("Cust_item_name_2", "string")
			blnReqCustRegItemName2 = CBool(getConfigurationItem("Cust_item_name_req_2", "bool"))
			blnViewCustRegItemName2 = CBool(getConfigurationItem("Cust_item_name_view_2", "bool"))
			strCustRegItemName3 = getConfigurationItem("Cust_item_name_3", "string")
			blnReqCustRegItemName3 = CBool(getConfigurationItem("Cust_item_name_req_3", "bool"))
			blnViewCustRegItemName3 = CBool(getConfigurationItem("Cust_item_name_view_3", "bool"))
			
			strStatsTrackingCode = getConfigurationItem("Stats_tracking_code", "string")
			strForumHeaderAd = getConfigurationItem("Forum_header_ad", "string")
			strForumPostAd = getConfigurationItem("Forum_post_ad", "string")
			strVigLinkKey = getConfigurationItem("VigLink_key", "string")
			blnShowLatestPosts = CBool(getConfigurationItem("Show_latest_posts", "bool"))
			strAnswerPosts = getConfigurationItem("Answer_posts", "string")
			strAnswerPostsWording = getConfigurationItem("Answer_wording", "string")
			blnShareTopicLinks = CBool(getConfigurationItem("Share_topics_links", "bool"))	
			blnPostThanks = CBool(getConfigurationItem("Post_thanks", "bool"))
			blnFacebookLike = CBool(getConfigurationItem("Facebook_likes", "bool"))
			blnTwitterTweet = CBool(getConfigurationItem("Twitter_tweet", "bool"))
			strFacebookPageID = getConfigurationItem("Facebook_page_ID", "string")
			strFacebookImage = getConfigurationItem("Facebook_image", "string")
			blnHttpXmlApi = CBool(getConfigurationItem("HTTP_XML_API", "bool"))
			blnUploadSecurityCheck = CBool(getConfigurationItem("Upload_file_scan", "bool"))
			blnEmailNotificationSendAll = CBool(getConfigurationItem("Email_all_notifications", "bool"))
			blnDisplayForumStats = CBool(getConfigurationItem("Show_Forum_Stats", "bool"))
			blnDisplayMemberList = CBool(getConfigurationItem("Show_Member_list", "bool"))
			strForumsMessage = getConfigurationItem("Forums_message", "string")
			blnUrlRewrite = CBool(getConfigurationItem("URL_Rewriting", "bool"))
			blnGooglePlusOne = CBool(getConfigurationItem("Google_plus_1", "bool"))
			blnModViewIpAddresses = CBool(getConfigurationItem("Mod_View_IP", "bool"))
			intSearchTimeDefault = CInt(getConfigurationItem("Search_time_default", "numeric"))
			strMemberProfileView = getConfigurationItem("Member_Profile_View", "string")
			blnStopForumSpam = CBool(getConfigurationItem("SFSpam", "bool"))
			strStopForumSpamApiKey = getConfigurationItem("SFSpam_API_Key", "string")
			lngStopForumSpamBlocked = CLng(getConfigurationItem("SFSpam_no_blocked", "numeric"))
			blnStopForumSpamUsername = CBool(getConfigurationItem("SFSpam_Username", "bool"))
			strSslEnabled = getConfigurationItem("SSL", "string")
			strRegistrationRedirect = getConfigurationItem("Reg_Redirect", "string")
			blnEuCookieLaw = CBool(getConfigurationItem("EU_Cookie_Law", "bool"))
			blnLocalLinksSameWindow = CBool(getConfigurationItem("Local_links_window", "bool"))
			blnRegistrationCAPTCHA = CBool(getConfigurationItem("Registration_CAPTCHA", "bool"))
			intSessionTimeout = CInt(getConfigurationItem("Session_Timeout", "numeric"))
			blnSessionIpCheck = CBool(getConfigurationItem("Session_IP_Check", "bool"))
			blnUploadDuplicateFiles = CBool(getConfigurationItem("Upload_Duplicate_Files", "bool"))
			blnCovertTextLinkstoHyperLink = CBool(getConfigurationItem("Text_Link_To_Hyperlink", "bool"))
			blnHttpsWebWizApi = CBool(getConfigurationItem("Web_Wiz_API_HTTPS", "bool"))
			strRegPrivacyNotice = getConfigurationItem("Reg_Privacy_Notice", "string")
			
		
			blnShowHeaderFooter = CBool(getConfigurationItem("Show_header_footer", "bool"))
			blnShowMobileHeaderFooter = CBool(getConfigurationItem("Show_mobile_header_footer", "bool"))
			strHeader = getConfigurationItem("Header", "string")
			strFooter = getConfigurationItem("Footer", "string")
			strHeaderMobile	= getConfigurationItem("Header_mobile", "string")
			strFooterMobile = getConfigurationItem("Footer_mobile", "string")
			
			

			'Initilise some elements if they are blank
			If intTopicPerPage = 0 Then intTopicPerPage = 24
			If strTitleImage = "" Then strTitleImage = "forum_images/web_wiz_forums.png"
			If intThreadsPerPage = 0 Then intThreadsPerPage = 12
			If strCSSfile = "" Then strCSSfile = "css_styles/default/"
			If strImagePath = "" Then strImagePath = "forum_images"
			If strNavSpacer = "" Then strNavSpacer = " - "
			If strBoardMetaDescription = "" Then strBoardMetaDescription = "This is a discussion forum powered by Web Wiz Forums. To find out about Web Wiz Forums, go to http://www.WebWizForums.com"
			If strBoardMetaKeywords = "" Then strBoardMetaKeywords = "community,forums,chat,talk,discussions"
			If intMaxImageHeight = 0 Then intMaxImageHeight = 0
			If intMaxImageWidth = 0 Then intMaxImageWidth = 0
				
			
			
			'If we are using application level variables the configuration into the application level variables
			If blnUseApplicationVariables Then
				
				'Lock the application so only one user updates it at a time
				Application.Lock
				
				'read in the configuration details from the recordset
				Application(strAppPrefix & "strWebsiteName") = strWebsiteName
				Application(strAppPrefix & "strMainForumName") = strMainForumName
				Application(strAppPrefix & "strWebsiteURL") = strWebsiteURL
				Application(strAppPrefix & "strForumPath") = strForumPath
				Application(strAppPrefix & "strMailComponent") = strMailComponent
				Application(strAppPrefix & "strMailServer") = strMailServer
				Application(strAppPrefix & "strForumEmailAddress") = strForumEmailAddress
				'Application(strAppPrefix & "blnLCode") = CBool(blnLCode)
				Application(strAppPrefix & "blnEmail") = CBool(blnEmail)
				Application(strAppPrefix & "blnRTEEditor") = CBool(blnRTEEditor)
				Application(strAppPrefix & "intTopicPerPage") = CInt(intTopicPerPage)
				Application(strAppPrefix & "strTitleImage") = strTitleImage
				Application(strAppPrefix & "blnEmoticons") = CBool(blnEmoticons)
				Application(strAppPrefix & "blnAvatar") = CBool(blnAvatar)
				Application(strAppPrefix & "blnEmailActivation") = CBool(blnEmailActivation)
			 	Application(strAppPrefix & "intNumHotViews") = CInt(intNumHotViews)
				Application(strAppPrefix & "intNumHotReplies") = CInt(intNumHotReplies)
				Application(strAppPrefix & "blnSendPost") = CBool(blnSendPost)
				Application(strAppPrefix & "blnPrivateMessages") = CBool(blnPrivateMessages)
				Application(strAppPrefix & "intPmInbox") = CInt(intPmInbox)
				Application(strAppPrefix & "intPmOutbox") = CInt(intPmOutbox)
				Application(strAppPrefix & "intThreadsPerPage") = CInt(intThreadsPerPage)
				Application(strAppPrefix & "intSpamTimeLimitSeconds") = CInt(intSpamTimeLimitSeconds)
				Application(strAppPrefix & "intSpamTimeLimitMinutes") = CInt(intSpamTimeLimitMinutes)
				Application(strAppPrefix & "intMaxPollChoices") = CInt(intMaxPollChoices)
				Application(strAppPrefix & "blnEmailMessenger") = CBool(blnEmailMessenger)
				Application(strAppPrefix & "blnActiveUsers") = CBool(blnActiveUsers)
				Application(strAppPrefix & "blnForumClosed") = CBool(blnForumClosed)
				Application(strAppPrefix & "blnShowEditUser") = CBool(blnShowEditUser)
				Application(strAppPrefix & "blnShowProcessTime") = CBool(blnShowProcessTime)
				Application(strAppPrefix & "blnFlashFiles") = CBool(blnFlashFiles)
				Application(strAppPrefix & "blnShowMod") = CBool(blnShowMod)
				Application(strAppPrefix & "blnAvatarUploadEnabled") = CBool(blnAvatarUploadEnabled)
				Application(strAppPrefix & "intMaxImageHeight") = CInt(intMaxImageHeight)
				Application(strAppPrefix & "intMaxImageWidth") = CInt(intMaxImageWidth)
				Application(strAppPrefix & "blnRegistrationSuspeneded") = CBool(blnRegistrationSuspeneded)
				Application(strAppPrefix & "strUploadComponent") = strUploadComponent
				Application(strAppPrefix & "strImageTypes") = strImageTypes
				Application(strAppPrefix & "lngUploadMaxImageSize") = CLng(lngUploadMaxImageSize)
				Application(strAppPrefix & "strUploadFileTypes") = strUploadFileTypes
				Application(strAppPrefix & "lngUploadMaxFileSize") = CLng(lngUploadMaxFileSize)
				Application(strAppPrefix & "intUploadAllocatedSpace") = CLng(intUploadAllocatedSpace)
				Application(strAppPrefix & "strMailServerUser") = strMailServerUser
				Application(strAppPrefix & "strMailServerPass") = strMailServerPass
				Application(strAppPrefix & "strCSSfile") = strCSSfile
				Application(strAppPrefix & "strNavSpacer") = strNavSpacer
				Application(strAppPrefix & "strImagePath") = strImagePath
				Application(strAppPrefix & "blnTopicIcon") = CBool(blnTopicIcon)
				Application(strAppPrefix & "blnLongRegForm") = CBool(blnLongRegForm)
				Application(strAppPrefix & "blnCAPTCHAsecurityImages") = CBool(blnCAPTCHAsecurityImages)
				Application(strAppPrefix & "blnCalendar") = CBool(blnCalendar)
				Application(strAppPrefix & "blnGuestSessions") = CBool(blnGuestSessions)
				Application(strAppPrefix & "blnMemberApprove") = CBool(blnMemberApprove)	
				Application(strAppPrefix & "blnRSS") = CBool(blnRSS)
				Application(strAppPrefix & "strInstallID") = strInstallID
				Application(strAppPrefix & "intPmFlood") = CInt(intPmFlood)
				Application(strAppPrefix & "blnACode") = CBool(blnACode)
				Application(strAppPrefix & "blnWebWizNewsPad") = CBool(blnWebWizNewsPad)
				Application(strAppPrefix & "strWebWizNewsPadURL") = strWebWizNewsPadURL
				Application(strAppPrefix & "strAvatarTypes") = strAvatarTypes
				Application(strAppPrefix & "intMaxAvatarSize") = CInt(intMaxAvatarSize)
				Application(strAppPrefix & "lngMostEverActiveUsers") = CLng(lngMostEverActiveUsers)
				Application(strAppPrefix & "dtmMostEverActiveDate") = CDate(dtmMostEverActiveDate)
				Application(strAppPrefix & "lngMailServerPort") = CLng(lngMailServerPort)
				Application(strAppPrefix & "strPageEncoding") = strPageEncoding
				Application(strAppPrefix & "strTextDirection") = strTextDirection
				Application(strAppPrefix & "strCookiePrefix") = strCookiePrefix
				Application(strAppPrefix & "strCookiePath") = strCookiePath
				Application(strAppPrefix & "blnDatabaseHeldSessions") = CBool(blnDatabaseHeldSessions)
				Application(strAppPrefix & "blnNewUserCode") = CBool(blnNewUserCode)
				Application(strAppPrefix & "blnModeratorProfileEdit") = CBool(blnModeratorProfileEdit)
				Application(strAppPrefix & "blnForumViewing") = CBool(blnForumViewing)
				Application(strAppPrefix & "blnDetailedErrorReporting") = CBool(blnDetailedErrorReporting)
				Application(strAppPrefix & "blnDisplayBirthdays") = CBool(blnDisplayBirthdays)
				Application(strAppPrefix & "blnDisplayTodaysBirthdays") = CBool(blnDisplayTodaysBirthdays)
				Application(strAppPrefix & "intRssTimeToLive") = CInt(intRssTimeToLive)
				Application(strAppPrefix & "intRSSmaxResults") = CInt(intRSSmaxResults)
				Application(strAppPrefix & "strDefaultPostOrder") = strDefaultPostOrder
				Application(strAppPrefix & "strCookieDomain") = strCookieDomain
				Application(strAppPrefix & "blnYouTube") = CBool(blnYouTube)
				Application(strAppPrefix & "strPMoverAction") = strPMoverAction
				Application(strAppPrefix & "blnPmFlashFiles") = CBool(blnPmFlashFiles)
				Application(strAppPrefix & "blnPmYouTube") = CBool(blnPmYouTube)
				Application(strAppPrefix & "intMinPasswordLength") = CInt(intMinPasswordLength)
				Application(strAppPrefix & "intMinUsernameLength") = CInt(intMinUsernameLength)
				Application(strAppPrefix & "blnQuickReplyForm") = CBool(blnQuickReplyForm)
				Application(strAppPrefix & "intIncorrectLoginAttempts") = CInt(intIncorrectLoginAttempts)
				Application(strAppPrefix & "intPointsReply") = CInt(intPointsReply)
				Application(strAppPrefix & "intPointsTopic") = CInt(intPointsTopic)
				Application(strAppPrefix & "intPointsAnswered") = CInt(intPointsAnswered)
				Application(strAppPrefix & "intPointsThanked") = CInt(intPointsThanked)
				Application(strAppPrefix & "intEditedTimeDelay") = CInt(intEditedTimeDelay)
				Application(strAppPrefix & "blnEnforceComplexPasswords") = CBool(blnEnforceComplexPasswords)
				Application(strAppPrefix & "blnRealNameReq") = CBool(blnRealNameReq)
				Application(strAppPrefix & "blnLocationReq") = CBool(blnLocationReq)
				Application(strAppPrefix & "blnTopicRating") = CBool(blnTopicRating)
				Application(strAppPrefix & "blnBoldNewTopics") = CBool(blnBoldNewTopics)
				Application(strAppPrefix & "blnBoldToday") = CBool(blnBoldToday)
				Application(strAppPrefix & "strBoardMetaDescription") = strBoardMetaDescription
				Application(strAppPrefix & "strBoardMetaKeywords") = strBoardMetaKeywords
				Application(strAppPrefix & "blnDynamicMetaTags") = CBool(blnDynamicMetaTags)
				Application(strAppPrefix & "blnSearchEngineSessions") = CBool(blnSearchEngineSessions)
				Application(strAppPrefix & "blnNoFollowTagInLinks") = CBool(blnNoFollowTagInLinks)
				Application(strAppPrefix & "blnSignatures") = CBool(blnSignatures)
				Application(strAppPrefix & "blnHomePage") = CBool(blnHomePage)
				Application(strAppPrefix & "blnSeoTitleQueryStrings") = CBool(blnSeoTitleQueryStrings)
				Application(strAppPrefix & "intEditPostTimeFrame") = CInt(intEditPostTimeFrame)
				Application(strAppPrefix & "blnChatRoom") = CBool(blnChatRoom)
				Application(strAppPrefix & "blnPmIgnoreSpamFilter") = CBool(blnPmIgnoreSpamFilter)
				Application(strAppPrefix & "blnMobileView") = CBool(blnMobileView)
				Application(strAppPrefix & "strRegistrationRules") = strRegistrationRules
				
				Application(strAppPrefix & "strCustRegItemName1") = strCustRegItemName1
				Application(strAppPrefix & "blnReqCustRegItemName1") = CBool(blnReqCustRegItemName1)
				Application(strAppPrefix & "blnViewCustRegItemName1") = CBool(blnViewCustRegItemName1)
				Application(strAppPrefix & "strCustRegItemName2") = strCustRegItemName2
				Application(strAppPrefix & "blnReqCustRegItemName2") = CBool(blnReqCustRegItemName2)
				Application(strAppPrefix & "blnViewCustRegItemName2") = CBool(blnViewCustRegItemName2)
				Application(strAppPrefix & "strCustRegItemName3") = strCustRegItemName3
				Application(strAppPrefix & "blnReqCustRegItemName3") = CBool(blnReqCustRegItemName3)
				Application(strAppPrefix & "blnViewCustRegItemName3") = CBool(blnViewCustRegItemName3)
				
				Application(strAppPrefix & "strStatsTrackingCode") = strStatsTrackingCode
				Application(strAppPrefix & "strForumHeaderAd") = strForumHeaderAd
				Application(strAppPrefix & "strForumPostAd") = strForumPostAd
				Application(strAppPrefix & "strVigLinkKey") = strVigLinkKey
				Application(strAppPrefix & "blnShowLatestPosts") = CBool(blnShowLatestPosts)
				Application(strAppPrefix & "strAnswerPosts") = strAnswerPosts
				Application(strAppPrefix & "strAnswerPostsWording") = strAnswerPostsWording
				Application(strAppPrefix & "blnShareTopicLinks") = CBool(blnShareTopicLinks)
				Application(strAppPrefix & "blnPostThanks") = CBool(blnPostThanks)
				Application(strAppPrefix & "blnFacebookLike") = CBool(blnFacebookLike)
				Application(strAppPrefix & "blnTwitterTweet") = CBool(blnTwitterTweet)
				Application(strAppPrefix & "strFacebookPageID") = strFacebookPageID
				Application(strAppPrefix & "strFacebookImage") = strFacebookImage
				Application(strAppPrefix & "blnHttpXmlApi") = CBool(blnHttpXmlApi)
				Application(strAppPrefix & "blnUploadSecurityCheck") = CBool(blnUploadSecurityCheck)
				Application(strAppPrefix & "blnEmailNotificationSendAll") = CBool(blnEmailNotificationSendAll)
				Application(strAppPrefix & "blnDisplayForumStats") = CBool(blnDisplayForumStats)
				Application(strAppPrefix & "blnDisplayMemberList") = CBool(blnDisplayMemberList)
				Application(strAppPrefix & "strForumsMessage") = strForumsMessage
				Application(strAppPrefix & "blnUrlRewrite") = CBool(blnUrlRewrite)
				Application(strAppPrefix & "blnGooglePlusOne") = CBool(blnGooglePlusOne)
				Application(strAppPrefix & "blnModViewIpAddresses") = CBool(blnModViewIpAddresses)
				Application(strAppPrefix & "intSearchTimeDefault") = Cint(intSearchTimeDefault)
				Application(strAppPrefix & "strMemberProfileView") = strMemberProfileView
				Application(strAppPrefix & "blnStopForumSpam") = blnStopForumSpam
				Application(strAppPrefix & "strStopForumSpamApiKey") = strStopForumSpamApiKey
				Application(strAppPrefix & "lngStopForumSpamBlocked") = CLng(lngStopForumSpamBlocked)
				Application(strAppPrefix & "blnStopForumSpamUsername") = CBool(blnStopForumSpamUsername)
				Application(strAppPrefix & "strSslEnabled") = strSslEnabled
				Application(strAppPrefix & "strRegistrationRedirect") = strRegistrationRedirect
				Application(strAppPrefix & "blnEuCookieLaw") = CBool(blnEuCookieLaw)
				Application(strAppPrefix & "blnLocalLinksSameWindow") = CBool(blnLocalLinksSameWindow)
				Application(strAppPrefix & "blnRegistrationCAPTCHA") = CBool(blnRegistrationCAPTCHA)
				Application(strAppPrefix & "strCountryBlockRegList") = strCountryBlockRegList
				Application(strAppPrefix & "intSessionTimeout") = CInt(intSessionTimeout)
				Application(strAppPrefix & "blnSessionIpCheck") = CBool(blnSessionIpCheck)
				Application(strAppPrefix & "blnUploadDuplicateFiles") = CBool(blnUploadDuplicateFiles)
				Application(strAppPrefix & "blnCovertTextLinkstoHyperLink") = CBool(blnCovertTextLinkstoHyperLink)
				Application(strAppPrefix & "blnHttpsWebWizApi") = CBool(blnHttpsWebWizApi)
				Application(strAppPrefix & "strRegPrivacyNotice") = strRegPrivacyNotice
				
				
				
				
				Application(strAppPrefix & "blnShowHeaderFooter") = CBool(blnShowHeaderFooter)
				Application(strAppPrefix & "blnShowMobileHeaderFooter") = CBool(blnShowMobileHeaderFooter)
				Application(strAppPrefix & "strHeader") = strHeader
				Application(strAppPrefix & "strFooter") = strFooter
				Application(strAppPrefix & "strHeaderMobile") = strHeaderMobile
				Application(strAppPrefix & "strFooterMobile") = strFooterMobile
				
				
				
				
				
				
				'Set the configuartion set application variable to true
				Application(strAppPrefix & "blnConfigurationSet") = True
				
				'Unlock the application
				Application.UnLock
			
			End If
			
		'Else no records returned	
		Else
		
			'Close the recordset
			rsCommon.Close
		
		End If
	
	
	'If we are using application level variables for the forum configuration then load in the variables from the application variables.
	ElseIf blnUseApplicationVariables Then	
			
			'read in the configuration details from the application varaibales
			strWebsiteName = Application(strAppPrefix & "strWebsiteName")
			strMainForumName = Application(strAppPrefix & "strMainForumName")
			strWebsiteURL = Application(strAppPrefix & "strWebsiteURL")
			strForumPath = Application(strAppPrefix & "strForumPath")
			strMailComponent = Application(strAppPrefix & "strMailComponent")
			strMailServer = Application(strAppPrefix & "strMailServer")
			strForumEmailAddress = Application(strAppPrefix & "strForumEmailAddress")
			'blnLCode = CBool(Application(strAppPrefix & "blnLCode"))
			blnEmail = CBool(Application(strAppPrefix & "blnEmail"))
			blnRTEEditor = CBool(Application(strAppPrefix & "blnRTEEditor"))
			intTopicPerPage = CInt(Application(strAppPrefix & "intTopicPerPage"))
			strTitleImage = Application(strAppPrefix & "strTitleImage")
			blnEmoticons = CBool(Application(strAppPrefix & "blnEmoticons"))
			blnAvatar = CBool(Application(strAppPrefix & "blnAvatar"))
			blnEmailActivation = CBool(Application(strAppPrefix & "blnEmailActivation"))
		 	intNumHotViews = CInt(Application(strAppPrefix & "intNumHotViews"))
			intNumHotReplies = CInt(Application(strAppPrefix & "intNumHotReplies"))
			blnSendPost = CBool(Application(strAppPrefix & "blnSendPost"))
			blnPrivateMessages = CBool(Application(strAppPrefix & "blnPrivateMessages"))
			intPmInbox = CInt(Application(strAppPrefix & "intPmInbox"))
			intPmOutbox = CInt(Application(strAppPrefix & "intPmOutbox"))
			intThreadsPerPage = CInt(Application(strAppPrefix & "intThreadsPerPage"))
			intSpamTimeLimitSeconds = CInt(Application(strAppPrefix & "intSpamTimeLimitSeconds"))
			intSpamTimeLimitMinutes = CInt(Application(strAppPrefix & "intSpamTimeLimitMinutes"))
			intMaxPollChoices = CInt(Application(strAppPrefix & "intMaxPollChoices"))
			blnEmailMessenger = CBool(Application(strAppPrefix & "blnEmailMessenger"))
			blnActiveUsers = CBool(Application(strAppPrefix & "blnActiveUsers"))
			blnForumClosed = CBool(Application(strAppPrefix & "blnForumClosed"))
			blnShowEditUser = CBool(Application(strAppPrefix & "blnShowEditUser"))
			blnShowProcessTime = CBool(Application(strAppPrefix & "blnShowProcessTime"))
			blnFlashFiles = CBool(Application(strAppPrefix & "blnFlashFiles"))
			blnShowMod = CBool(Application(strAppPrefix & "blnShowMod"))
			blnAvatarUploadEnabled = CBool(Application(strAppPrefix & "blnAvatarUploadEnabled"))
			intMaxImageHeight = CInt(Application(strAppPrefix & "intMaxImageHeight"))
			intMaxImageWidth = CInt(Application(strAppPrefix & "intMaxImageWidth"))
			blnRegistrationSuspeneded = CBool(Application(strAppPrefix & "blnRegistrationSuspeneded"))
			strImageTypes = Application(strAppPrefix & "strImageTypes")
			strUploadComponent = Application(strAppPrefix & "strUploadComponent")
			lngUploadMaxImageSize = CLng(Application(strAppPrefix & "lngUploadMaxImageSize"))
			strUploadFileTypes = Application(strAppPrefix & "strUploadFileTypes")
			lngUploadMaxFileSize = CLng(Application(strAppPrefix & "lngUploadMaxFileSize"))
			intUploadAllocatedSpace = CLng(Application(strAppPrefix & "intUploadAllocatedSpace"))	
			strMailServerUser = Application(strAppPrefix & "strMailServerUser")
			strMailServerPass = Application(strAppPrefix & "strMailServerPass")
			strCSSfile = Application(strAppPrefix & "strCSSfile")
			strNavSpacer = Application(strAppPrefix & "strNavSpacer")
			strImagePath = Application(strAppPrefix & "strImagePath")
			blnTopicIcon = CBool(Application(strAppPrefix & "blnTopicIcon"))
			blnLongRegForm = CBool(Application(strAppPrefix & "blnLongRegForm"))
			blnCAPTCHAsecurityImages = CBool(Application(strAppPrefix & "blnCAPTCHAsecurityImages"))
			blnCalendar = CBool(Application(strAppPrefix & "blnCalendar"))
			blnGuestSessions = CBool(Application(strAppPrefix & "blnGuestSessions"))
			blnMemberApprove = CBool(Application(strAppPrefix & "blnMemberApprove"))
			blnRSS = CBool(Application(strAppPrefix & "blnRSS"))
			strInstallID = Application(strAppPrefix & "strInstallID")
			intPmFlood = CInt(Application(strAppPrefix & "intPmFlood"))
			blnACode = CBool(Application(strAppPrefix & "blnACode"))
			blnWebWizNewsPad = CBool(Application(strAppPrefix & "blnWebWizNewsPad"))
			strWebWizNewsPadURL = Application(strAppPrefix & "strWebWizNewsPadURL")
			strAvatarTypes = Application(strAppPrefix & "strAvatarTypes")
			intMaxAvatarSize = CInt(Application(strAppPrefix & "intMaxAvatarSize"))
			
			lngMostEverActiveUsers = CLng(Application(strAppPrefix & "lngMostEverActiveUsers"))
			dtmMostEverActiveDate = CDate(Application(strAppPrefix & "dtmMostEverActiveDate"))
			lngMailServerPort = CLng(Application(strAppPrefix & "lngMailServerPort"))
			strPageEncoding = Application(strAppPrefix & "strPageEncoding")
			strTextDirection = Application(strAppPrefix & "strTextDirection")
			strCookiePrefix = Application(strAppPrefix & "strCookiePrefix")
			strCookiePath = Application(strAppPrefix & "strCookiePath")
			blnDatabaseHeldSessions = CBool(Application(strAppPrefix & "blnDatabaseHeldSessions"))
			blnNewUserCode = CBool(Application(strAppPrefix & "blnNewUserCode"))
			blnModeratorProfileEdit = CBool(Application(strAppPrefix & "blnModeratorProfileEdit"))
			blnForumViewing = CBool(Application(strAppPrefix & "blnForumViewing"))
			blnDetailedErrorReporting = CBool(Application(strAppPrefix & "blnDetailedErrorReporting"))
			blnDisplayBirthdays = CBool(Application(strAppPrefix & "blnDisplayBirthdays"))
			blnDisplayTodaysBirthdays = CBool(Application(strAppPrefix & "blnDisplayTodaysBirthdays"))
			intRssTimeToLive = CInt(Application(strAppPrefix & "intRssTimeToLive"))
			intRSSmaxResults = CInt(Application(strAppPrefix & "intRSSmaxResults"))
			strDefaultPostOrder = Application(strAppPrefix & "strDefaultPostOrder")
			strCookieDomain = Application(strAppPrefix & "strCookieDomain")
			blnYouTube = CBool(Application(strAppPrefix & "blnYouTube"))
			strPMoverAction = Application(strAppPrefix & "strPMoverAction")
			blnPmFlashFiles = CBool(Application(strAppPrefix & "blnPmFlashFiles"))
			blnPmYouTube = CBool(Application(strAppPrefix & "blnPmYouTube"))
			intMinPasswordLength = CInt(Application(strAppPrefix & "intMinPasswordLength"))
			intMinUsernameLength = CInt(Application(strAppPrefix & "intMinUsernameLength"))
			blnQuickReplyForm = CBool(Application(strAppPrefix & "blnQuickReplyForm"))
			intIncorrectLoginAttempts = CInt(Application(strAppPrefix & "intIncorrectLoginAttempts"))
			intPointsReply = CInt(Application(strAppPrefix & "intPointsReply"))
			intPointsTopic = CInt(Application(strAppPrefix & "intPointsTopic"))
			intPointsAnswered = CInt(Application(strAppPrefix & "intPointsAnswered"))
			intPointsThanked = CInt(Application(strAppPrefix & "intPointsThanked"))
			intEditedTimeDelay = CInt(Application(strAppPrefix & "intEditedTimeDelay"))
			blnEnforceComplexPasswords = CBool(Application(strAppPrefix & "blnEnforceComplexPasswords"))
			blnRealNameReq = CBool(Application(strAppPrefix & "blnRealNameReq"))
			blnLocationReq = CBool(Application(strAppPrefix & "blnLocationReq"))
			blnTopicRating = CBool(Application(strAppPrefix & "blnTopicRating"))
			blnBoldNewTopics = CBool(Application(strAppPrefix & "blnBoldNewTopics"))
			blnBoldToday = CBool(Application(strAppPrefix & "blnBoldToday"))
			strBoardMetaDescription = Application(strAppPrefix & "strBoardMetaDescription")
			strBoardMetaKeywords = Application(strAppPrefix & "strBoardMetaKeywords")
			blnDynamicMetaTags = CBool(Application(strAppPrefix & "blnDynamicMetaTags"))
			blnSearchEngineSessions = CBool(Application(strAppPrefix & "blnSearchEngineSessions"))
			blnNoFollowTagInLinks = CBool(Application(strAppPrefix & "blnNoFollowTagInLinks"))
			blnSignatures = CBool(Application(strAppPrefix & "blnSignatures"))
			blnHomePage = CBool(Application(strAppPrefix & "blnHomePage"))
			blnSeoTitleQueryStrings = CBool(Application(strAppPrefix & "blnSeoTitleQueryStrings"))
			intEditPostTimeFrame = CInt(Application(strAppPrefix & "intEditPostTimeFrame"))
			blnChatRoom = CBool(Application(strAppPrefix & "blnChatRoom"))
			blnPmIgnoreSpamFilter = CBool(Application(strAppPrefix & "blnPmIgnoreSpamFilter"))
			blnMobileView = CBool(Application(strAppPrefix & "blnMobileView"))
			strRegistrationRules = Application(strAppPrefix & "strRegistrationRules")
			
			strCustRegItemName1 = Application(strAppPrefix & "strCustRegItemName1")
			blnReqCustRegItemName1 = CBool(Application(strAppPrefix & "blnReqCustRegItemName1"))
			blnViewCustRegItemName1 = CBool(Application(strAppPrefix & "blnViewCustRegItemName1"))
			strCustRegItemName2 = Application(strAppPrefix & "strCustRegItemName2")
			blnReqCustRegItemName2 = CBool(Application(strAppPrefix & "blnReqCustRegItemName2"))
			blnViewCustRegItemName2 = CBool(Application(strAppPrefix & "blnViewCustRegItemName2"))
			strCustRegItemName3 = Application(strAppPrefix & "strCustRegItemName3")
			blnReqCustRegItemName3 = CBool(Application(strAppPrefix & "blnReqCustRegItemName3"))
			blnViewCustRegItemName3 = CBool(Application(strAppPrefix & "blnViewCustRegItemName3"))
			
			strStatsTrackingCode = Application(strAppPrefix & "strStatsTrackingCode")
			strForumHeaderAd = Application(strAppPrefix & "strForumHeaderAd")
			strForumPostAd = Application(strAppPrefix & "strForumPostAd")
			strVigLinkKey = Application(strAppPrefix & "strVigLinkKey")
			blnShowLatestPosts = CBool(Application(strAppPrefix & "blnShowLatestPosts"))
			strAnswerPosts = Application(strAppPrefix & "strAnswerPosts")
			strAnswerPostsWording = Application(strAppPrefix & "strAnswerPostsWording")
			blnShareTopicLinks = CBool(Application(strAppPrefix & "blnShareTopicLinks"))
			blnPostThanks = CBool(Application(strAppPrefix & "blnPostThanks"))
			blnFacebookLike = CBool(Application(strAppPrefix & "blnFacebookLike"))
			blnTwitterTweet = CBool(Application(strAppPrefix & "blnTwitterTweet"))
			strFacebookPageID = Application(strAppPrefix & "strFacebookPageID")
			strFacebookImage = Application(strAppPrefix & "strFacebookImage")
			blnHttpXmlApi = CBool(Application(strAppPrefix & "blnHttpXmlApi"))
			blnUploadSecurityCheck = CBool(Application(strAppPrefix & "blnUploadSecurityCheck"))
			blnEmailNotificationSendAll = CBool(Application(strAppPrefix & "blnEmailNotificationSendAll"))
			blnDisplayForumStats = CBool(Application(strAppPrefix & "blnDisplayForumStats"))
			blnDisplayMemberList = CBool(Application(strAppPrefix & "blnDisplayMemberList"))
			strForumsMessage = Application(strAppPrefix & "strForumsMessage")
			blnUrlRewrite = CBool(Application(strAppPrefix & "blnUrlRewrite"))
			blnGooglePlusOne = CBool(Application(strAppPrefix & "blnGooglePlusOne"))
			blnModViewIpAddresses = CBool(Application(strAppPrefix & "blnModViewIpAddresses"))
			intSearchTimeDefault = CInt(Application(strAppPrefix & "intSearchTimeDefault"))
			strMemberProfileView = Application(strAppPrefix & "strMemberProfileView")
			blnStopForumSpam = CBool(Application(strAppPrefix & "blnStopForumSpam"))
			strStopForumSpamApiKey = Application(strAppPrefix & "strStopForumSpamApiKey")
			lngStopForumSpamBlocked = CLng(Application(strAppPrefix & "lngStopForumSpamBlocked"))
			blnStopForumSpamUsername = CBool(Application(strAppPrefix & "blnStopForumSpamUsername"))
			strSslEnabled = Application(strAppPrefix & "strSslEnabled")
			strRegistrationRedirect = Application(strAppPrefix & "strRegistrationRedirect")
			blnEuCookieLaw = CBool(Application(strAppPrefix & "blnEuCookieLaw"))
			blnLocalLinksSameWindow = CBool(Application(strAppPrefix & "blnLocalLinksSameWindow"))
			blnRegistrationCAPTCHA = CBool(Application(strAppPrefix & "blnRegistrationCAPTCHA"))
			strCountryBlockRegList = Application(strAppPrefix & "strCountryBlockRegList")
			intSessionTimeout = CInt(Application(strAppPrefix & "intSessionTimeout"))
			blnSessionIpCheck = CBool(Application(strAppPrefix & "blnSessionIpCheck"))
			blnUploadDuplicateFiles = CBool(Application(strAppPrefix & "blnUploadDuplicateFiles"))
			blnCovertTextLinkstoHyperLink = CBool(Application(strAppPrefix & "blnCovertTextLinkstoHyperLink"))
			blnHttpsWebWizApi = CBool(Application(strAppPrefix & "blnHttpsWebWizApi"))
			strRegPrivacyNotice = Application(strAppPrefix & "strRegPrivacyNotice")
			
			

			blnShowHeaderFooter = CBool(Application(strAppPrefix & "blnShowHeaderFooter"))
			blnShowMobileHeaderFooter = CBool(Application(strAppPrefix & "blnShowMobileHeaderFooter"))
			strHeader = Application(strAppPrefix & "strHeader")
			strFooter = Application(strAppPrefix & "strFooter")
			strHeaderMobile = Application(strAppPrefix & "strHeaderMobile")
			strFooterMobile = Application(strAppPrefix & "strFooterMobile")
	End If


	'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
	'*** MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT AND COULD LEAD TO LEGAL ACTION BEING TAKEN!! ***
	blnLCode = blnWebWizForumsLinks
	blnCAPTCHAabout = blnLCode 
	If (strInstallID = "" OR isNull(strInstallID)) AND blnDemoMode = False Then 
		'Paid Licenses have a traceable installation ID registered with the license server
		strInstallID = "Free Express Edition Install"
		
	End If
	
	'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ******
	
	
	'Disable storing IP with session as causes many issues with IPv6 and mobile devices
	blnSessionIpCheck = False
	
	
	'If someone has placed the default.asp in the path to the forum then remove it as it's not needed
	strForumPath = Replace(strForumPath, "default.asp", "")
	
	'Make sure the web address has a / on the end
	If strForumPath <> "" AND isNull(strForumPath) = False Then
		If Mid(strForumPath, len(strForumPath), 1) <> "/" Then  strForumPath = strForumPath & "/"
	End If
	'Make sure the image path has a / on the end
	If strImagePath <> "" AND isNull(strImagePath) = False Then
		If Mid(strImagePath, len(strImagePath), 1) <> "/" Then  strImagePath = strImagePath & "/"
	End If
	
	'If Web Wiz NewsPad is enabled make sure we have the correct path	
	If isNull(strWebWizNewsPadURL) = False AND NOT strWebWizNewsPadURL = "" Then 
		strWebWizNewsPadURL = Replace(strWebWizNewsPadURL, "default.asp", "")
	
		'Make sure the web addresses has a / on the end
		If Mid(strWebWizNewsPadURL, len(strWebWizNewsPadURL), 1) <> "/" Then  strWebWizNewsPadURL = strWebWizNewsPadURL & "/"
	End If 
	
	
	'Make sure the CSS path has / on the end
	If strCSSfile <> "" AND isNull(strCSSfile) = False Then 
		If Mid(strCSSfile, len(strCSSfile), 1) <> "/" Then  strCSSfile = strCSSfile & "/"
	End If
	
	'If there is no session timeout set then set it to the default of 20
	If intSessionTimeout = 0 Then intSessionTimeout = 20
	

	
	'Get if mobile broser
	blnMobileBrowser = mobileBrowser()
	
	'Check the client browser version
	strClientBrowserVersion = browserDetect()
	
	'If the cleint browser is IE 6 or below then display GIF's for some of the images
	If strClientBrowserVersion = "MSIE6-" Then 
		strForumImageType = "gif"
	Else
		strForumImageType = "png"
	End If
	
	'Read in the OS Type
	strOSType = OSType
	
		
	'If the admin needs to approve the membership, disable email activation for the user
	If blnMemberApprove Then blnEmailActivation = false

	'For demo mode
	If blnDemoMode Then
		strInstallID = "Demo Mode"
		blnSendPost = False
		strMailComponent = ""
		blnEmailActivation = False
		strForumEmailAddress = ""
		blnEmailMessenger = False
		blnAvatarUploadEnabled = False
		blnRegistrationSuspeneded = False
		blnForumClosed = False
		blnMemberApprove = False
		blnPrivateMessages = True
		blnCalendar = True
		blnEmail = False
		strUploadComponent = "AspUpload"
		lngUploadMaxFileSize = 10240
		lngUploadMaxImageSize = 500
		intUploadAllocatedSpace = 25
		intPmInbox = 1000
		strPMoverAction = "block"
		blnRSS = True
		intPmOutbox = 1000
		blnSeoTitleQueryStrings = True
		blnChatRoom = True
		blnMobileView = True
		strMemberProfileView = "members"
		blnHttpsWebWizApi = True
	End If
	
	
End Sub


%>