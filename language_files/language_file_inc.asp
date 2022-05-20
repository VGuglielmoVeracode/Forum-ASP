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
Const strTxtWelcome = "Welcome"
Const strTxtAllForums = "All Forums"
Const strTxtTopics = "Topics"
Const strTxtPosts = "Posts"
Const strTxtLastPost = "Last Post"
Const strTxtPostPreview = "Post Preview"
Const strTxtAt = "at"
Const strTxtBy = "By"
Const strTxtOn = "on"
Const strTxtProfile = "Member Profile"
Const strTxtSearch = "Search"
Const strTxtQuote = "Quote"
Const strTxtVisit = "Visit"
Const strTxtView = "View"
Const strTxtHome = "Home"
Const strTxtHomepage = "Homepage"
Const strTxtEdit = "Edit"
Const strTxtDelete = "Delete"
Const strTxtEditProfile = "Edit Profile"
Const strTxtLogOff = "Logout"
Const strTxtRegister = "Register"
Const strTxtLogin = "Login"
Const strTxtMembersList = "Display List of Forum Members"
Const strTxtForumLocked = "Forum Locked"
Const strTxtSearchTheForum = "Forum Search"
Const strTxtPostReply = "Post Reply"
Const strTxtNewTopic = "New Topic"
Const strTxtNoForums = "There are no Forums to display"
Const strTxtReturnToDiscussionForum = "Return to the Discussion Forum"
Const strTxtMustBeRegistered = "You must be a registered forum member in order to use this feature."
Const strTxtClearForm = "Clear Form"
Const strTxtYes = "Yes"
Const strTxtNo = "No"
Const strTxtForumLockedByAdmim = "This function has been disabled.<br />This Forum has been locked by the Forum Administrator."
Const strTxtRequiredFields = "Indicates a required field"

Const strTxtForumJump = "Forum Jump"
Const strTxtSelectForum = "Select Forum"

Const strTxtNoMessageError = "Message Box - Enter a Message to Submit"
Const strTxtErrorDisplayLine = "_______________________________________________________________"
Const strTxtErrorDisplayLine1 = "The form has not been submitted because there are problem(s) with the form."
Const strTxtErrorDisplayLine2 = "Please correct the problem(s) and re-submit the form."
Const strTxtErrorDisplayLine3 = "The following field(s) need to be corrected: -"



'default.asp
'---------------------------------------------------------------------------------
Const strTxtCookies = "Cookies and JavaScript must be enabled on your web browser in order to use this forum"
Const strTxtForum = "Forum"
Const strTxtLatestForumPosts = "Latest Forum Posts"
Const strTxtForumStatistics = "Forum Statistics"
Const strTxtNoForumPostMade = "There have been no Forum Posts"
Const strTxtThereAre = "There are"
Const strTxtPostsIn = "Posts in"
Const strTxtTopicsIn = "Topics in"
Const strTxtLastPostBy = "Last Post by"
Const strTxtForumMembers = "Forum Members"
Const strTxtTheNewestForumMember = "The Newest Forum Member is"


'forum_topics.asp
'---------------------------------------------------------------------------------
Const strTxtThreadStarter = "Topic Starter"
Const strTxtReplies = "Replies"
Const strTxtViews = "Views"
Const strTxtDeleteTopicAlert = "Are you sure you want to delete this topic?"
Const strTxtDeleteTopic = "Delete Topic"
Const strTxtNextTopic = "Next Topic"
Const strTxtLastTopic = "Last Topic"
Const strTxtShowTopics = "Show Topics"
Const strTxtNoTopicsToDisplay = "There are no messages posted in this forum in the "

Const strTxtAll = "All"
Const strTxtLastWeek = "Last Week"
Const strTxtLastTwoWeeks = "Last Two Weeks"
Const strTxtLastMonth = "Last Month"
Const strTxtLastTwoMonths = "Last Two Months"
Const strTxtLastSixMonths = "Last Six Months"
Const strTxtLastYear = "Last Year"


'forum_posts.asp
'---------------------------------------------------------------------------------
Const strTxtLocation = "Location"
Const strTxtJoined = "Joined"
Const strTxtForumAdministrator = "Forum Administrator"
Const strTxtForumModerator = "Forum Moderator"
Const strTxtDeletePostAlert = "Are you sure you want to delete this post?"
Const strTxtEditPost = "Edit Post"
Const strTxtDeletePost = "Delete Post"
Const strTxtSearchForPosts = "Search for posts by"
Const strTxtSubjectFolder = "Subject"
Const strTxtPrintVersion = "Printable Version"
Const strTxtEmailTopic = "Email This Topic"
Const strTxtSorryNoReply = "You can NOT post a reply."
Const strTxtThisForumIsLocked = "This forum has been locked by a Forum Administrator."
Const strTxtPostAReplyRegister = "If you wish to post a reply to this topic you must first"
Const strTxtNeedToRegister = "If you are not already registered you must first"
Const strTxtSmRegister = "register"
Const strTxtNoThreads = "There are no posts in the database relating to this topic"
Const strTxtNotGiven = "Not Given"


'search_form.asp
'---------------------------------------------------------------------------------
Const strTxtSearchFormError = "Search For - Enter something to search for"


'search.asp
'---------------------------------------------------------------------------------
Const strTxtSearchResults = "Search Results"
Const strTxtHasFound = "has found"
Const strTxtResults = "results"
Const strTxtNoSearchResults = "Your search failed to find any results"
Const strTxtClickHereToRefineSearch = "Click here to amend your search"
Const strTxtSearchFor = "Search For"
Const strTxtSearchIn = "Search In"
Const strTxtSearchOn = "Search On"
Const strTxtAllWords = "All Words"
Const strTxtAnyWords = "Any Words"
Const strTxtPhrase = "Phrase"
Const strTxtTopicSubject = "Topic Subject"
Const strTxtMessageBody = "Message Body"
Const strTxtAuthor = "Author"
Const strTxtSearchForum = "Search Forum"
Const strTxtSortResultsBy = "Sort Results By"
Const strTxtLastPostTime = "Last Post Time"
Const strTxtTopicStartDate = "Topic Start Date"
Const strTxtSubjectAlphabetically = "Subject Alphabetically"
Const strTxtNumberViews = "Number of Views"
Const strTxtStartSearch = "Start Search"


'printer_friendly_posts.asp
'---------------------------------------------------------------------------------
Const strTxtPrintPage = "Print Page"
Const strTxtPrintedFrom = "Printed From"
Const strTxtForumName = "Forum Name"
Const strTxtForumDiscription = "Forum Description"
Const strTxtURL = "URL"
Const strTxtPrintedDate = "Printed Date"
Const strTxtTopic = "Topic"
Const strTxtPostedBy = "Posted By"
Const strTxtDatePosted = "Date Posted"


'emoticons.asp
'---------------------------------------------------------------------------------
Const strTxtEmoticonSmilies = "Emoticons"


'login.asp
'---------------------------------------------------------------------------------
Const strTxtSorryUsernamePasswordIncorrect = "The Username or Password entered is incorrect."
Const strTxtPleaseTryAgain = "Please try again."
Const strTxtUsername = "Username"
Const strTxtPassword = "Password"
Const strTxtLoginUser = "Forum Login"
Const strTxtClickHereForgottenPass = "Forgotten Your Password?"
Const strTxtErrorUsername = "Username  - Enter your Forum Username"
Const strTxtErrorPassword = "Password  - Enter your Forum Password"


'forgotten_password.asp
'---------------------------------------------------------------------------------
Const strTxtForgottenPassword = "Retrieve Login"
Const strTxtNoRecordOfUsername = "No match can be found for the details you have entered."
Const strTxtNoEmailAddressInProfile = "Your login details can not be emailed to you as your profile does not contain an email address."
Const strTxtReregisterForForum = "You will need to re-register to use the forum."
Const strTxtPasswordEmailToYou = "Your login details and new password has been emailed to you."
Const strTxtPleaseEnterYourUsername = "Please fill out the form below with your username or email address and your username and new password will be emailed to your registered email address."
Const strTxtEmailPassword = "Request New Password"

Const strTxtEmailPasswordRequest = "A request has been made to retrieve your login details for the forum,"
Const strTxtEmailPasswordRequest2 = "You will find your login details with your new password below."
Const strTxtEmailPasswordRequest3 = "To go to the forum now click on the link below: -"


'forum_password_form.asp
'---------------------------------------------------------------------------------
Const strTxtForumLogin = "Forum Login"
Const strTxtErrorEnterPassword = "Password  - Enter a Password to use this Forum"
Const strTxtPasswordRequiredForForum = "This is a private forum and requires that you enter a forum password to proceed."
Const strTxtForumPasswordIncorrect = "The Password entered is incorrect."
Const strTxtAutoLogin = "Keep me Logged-in on this computer (requires cookies)"
Const strTxtLoginToForum = "Login To Forum"


'profile.asp
'---------------------------------------------------------------------------------
Const strTxtNoUserProfileFound = "No profile can be found for this user"
Const strTxtRegisteredToViewProfile = "You must be a registered forum member to view profiles."
Const strTxtMemberNo = "Member No."
Const strTxtEmailAddress = "Email Address"
Const strTxtPrivate = "Private"


'new_topic_form.asp
'---------------------------------------------------------------------------------
Const strTxtPostNewTopic = "Create New Topic"
Const strTxtErrorTopicSubject = "Subject   - Enter a Subject for your new Topic"
Const strTxtForumMemberSuspended = "This function is disabled as your Forum Membership has been suspended!"

'edit_post_form.asp
'---------------------------------------------------------------------------------
Const strTxtNoPermissionToEditPost = "You do not have permission to edit this post!"
Const strTxtReturnForumTopic = "Return to Forum Topic"


'email_topic.asp
'---------------------------------------------------------------------------------
Const strTxtEmailTopicToFriend = "Email Topic To a Friend"
Const strTxtFriendSentEmail = "Your Friend has been sent the email"
Const strTxtFriendsName = "Friends Name"
Const strTxtFriendsEmail = "Friends Email"
Const strTxtYourName = "Your Name"
Const strTxtYourEmail = "Your Email"
Const strTxtSendEmail = "Send Email"
Const strTxtMessage = "Message"

Const strTxtEmailFriendMessage = "I thought you might be interested in a post on"
Const strTxtFrom = "from"

Const strTxtErrorFrinedsName = "Friends Name  - Enter your Friends Name"
Const strTxtErrorFriendsEmail = "Friends Email  - Enter a valid email address for your friend"
Const strTxtErrorYourName = "Your Name  - Enter your Name"
Const strTxtErrorYourEmail = "Your Email  - Enter your valid email address"
Const strTxtErrorEmailMessage = "Message  - Enter a message for your friend"


'members.asp
'---------------------------------------------------------------------------------
Const strTxtForumMembersList = "Forum Members List"
Const strTxtMemberSearch = "Member Search"

Const strTxtRegistered = "Registered"
Const strTxtSend = "Send"
Const strTxtNext = "Next"
Const strTxtPrevious = "Previous"
Const strTxtPage = "Page"

Const strTxtErrorMemberSerach = "Member Search - Enter a Members Username to search for"



'register.asp
'---------------------------------------------------------------------------------
Const strTxtRegisterNewUser = "Register New Member"

Const strTxtProfileUsernameLong = "This is the name displayed when you use the forum"
Const strTxtRetypePassword = "Retype Password"
Const strTxtProfileEmailLong = "Not required, but useful if you wish to be notified when someone answers one of your posts or if you lose your password."
Const strTxtShowHideEmail = "Show my Email Address"
Const strTxtShowHideEmailLong = "Hide your email address if you want it to be kept private from other users."
Const strTxtSelectCountry = "Select Country"
Const strTxtProfileAutoLogin = "Automatically log me in when I return to the Forum"
Const strTxtSignature = "Signature"
Const strTxtSignatureLong = "Enter a signature that you would like shown at the bottom of your Forum Posts"

Const strTxtErrorUsernameChar = "Username  - Your Username must be at least"
Const strTxtErrorPasswordChar = "Password  - Your Password must be at least"
Const strTxtErrorPasswordNoMatch = "Password Error - The passwords entered do not match"
Const strTxtErrorValidEmail = "Email  - Enter your valid email address"
Const strTxtErrorValidEmailLong = "If you don't want to enter your email address then leave the email field blank"
Const strTxtErrorNoEmailToShow = "You can not show your email address if you haven\'t entered one!"
Const strTxtErrorSignatureToLong = "Signature  - Your signature has to many characters"
Const strTxtUpdateProfile = "Update Profile"


Const strTxtUsrenameGone = "The Username you requested is already taken, not permitted, or is less than the required number of characters.\n\nPlease choose another Username."
Const strTxtEmailThankYouForRegistering = "Thank-you for taking the time to register to use the forum"
Const strTxtEmailYouCanNowUseTheForumAt = "Your login details can be found below and now you have registered for a new account you can post new messages and reply to existing ones on the"
Const strTxtEmailForumAt = "forum at"
Const strTxtEmailToThe = "to "


'register_new_user.inc
'---------------------------------------------------------------------------------
Const strTxtEmailAMeesageHasBeenPosted = "A message has been posted on"
Const strTxtEmailClickOnLinkBelowToView = "To view and/or reply to the post then click on the link below"
Const strTxtEmailAMeesageHasBeenPostedOnForumNum = "A message has been posted in the forum number"


'forum_terms.asp
'---------------------------------------------------------------------------------
Const strTxtForumRulesAndPolicies = "Forum Rules and Policies"
Const srtTxtAccept = "Accept"




'New from version 6
'---------------------------------------------------------------------------------
Const strTxtHi = "Hi"
Const strTxtInterestingForumPostOn = "Interesting Forum post on"
Const strTxtForumLostPasswordRequest = "Forum Login Details Request"
Const strTxtLockForum = "Lock Forum"
Const strTxtLockedTopic = "Closed&nbsp;Topic"
Const strTxtUnLockTopic = "Un-Lock Topic"
Const strTxtTopicLocked = "Topic Closed"
Const strTxtUnForumLocked = "Un-Lock Forum"
Const strTxtThisTopicIsLocked = "This topic is closed."
Const strTxtThatYouAskedKeepAnEyeOn = "that you asked us to keep an eye on."
Const strTxtTheTopicIsNowDeleted = "The Topic has now been Deleted."
Const strTxtOf = "of"
Const strTxtTheTimeNowIs = "The time now is"
Const strTxtYouLastVisitedOn = "You last visited"
Const strTxtSendMsg = "Send PM"
Const strTxtSendPrivateMessage = "Send Private Message"
Const strTxtActiveUsers = "Active Users"
Const strTxtMembers = "Member(s)"
Const strTxtEnterTextYouWouldLikeIn = "Enter the text that you would like in"
Const strTxtEmailAddressAlreadyUsed = "The email address entered has already been used to register another member."
Const strTxtIP = "IP"
Const strTxtIPLogged = "IP Logged"
Const strTxtPages = "Pages"
Const strTxtCharacterCount = "Character Count"
Const strTxtAdmin = "Admin"


Const strTxtType = "Group"
Const strTxtActive = "Active"
Const strTxtGuest = "Guest"
Const strTxtAccountStatus = "Account Status"
Const strTxtNotActive = "Not Active"



Const strTxtEmailRequiredForActvation = "Required to be able to receive an email to activate your membership"
Const strTxtToActivateYourMembershipFor = "To activate your membership for"
Const strTxtForumClickOnTheLinkBelow = "click on the link below"
Const strTxtForumAdmin = "Forum Admin"
Const strTxtViewLastPost = "View Last Post"
Const strTxtSelectAvatar = "Select Avatar"
Const strTxtAvatar = "Avatar"
Const strTxtSelectAvatarDetails = "This is the small icon shown next to your posts. Either select one from the list or type the path in to your own Avatar (must be 64 x 64 pixels)"
Const strTxtForumCodesInSignature = "can be used in your signature"

Const strTxtHighPriorityPost = "Announcement"
Const strTxtPinnedTopic = "Sticky Topic"

Const strTxtOpenForum = "Open Forum"
Const strTxtReadOnly = "Read Only"
Const strTxtPasswordRequired = "Password Required"
Const strTxtNoAccess = "No Access"

Const strTxtFont = "Font"
Const strTxtSize = "Size"
Const strTxtForumCodes = "BBcodes"

Const strTxtNormal = "Normal Topic"
Const strTxtTopAllForums = "Announcement (all forums)"
Const strTopThisForum = "Announcement (this forum)"


Const strTxtMarkAllPostsAsRead = "Mark all posts as read"
Const strTxtDeleteCookiesSetByThisForum = "Delete cookies set by this forum"


'forum_codes
'---------------------------------------------------------------------------------
Const strTxtYouCanUseForumCodesToFormatText = "You can use the following Forum Codes to Format your text"
Const strTxtTypedForumCode = "Typed Forum Code"
Const strTxtConvetedCode = "Converted Code"
Const strTxtTextFormating = "Text Formatting"
Const strTxtImagesAndLinks = "Images and Links"
Const strTxtMyLink = "My Link"
Const strTxtMyEmail = "My Email"



'insufficient_permission.asp
'---------------------------------------------------------------------------------
Const strTxtAccessDenied = "Access Denied"
Const strTxtInsufficientPermison = "Only members with sufficient permission can access this page."


'activate.asp
'---------------------------------------------------------------------------------
Const strTxtYourForumMemIsNowActive = "Your Forum membership is active."
Const strTxtErrorWithActvation = "There is a problem activating your membership.<br /><br />Please contact the "


'register_mail_confirm.asp
'---------------------------------------------------------------------------------
Const strTxtYouShouldReceiveAnEmail = "<strong>Your Forum Membership needs to be activated!</strong> <br /><br />An activation email will be sent in a few moments to the email address you gave when registering.<br />Click on the link in this email to activate your Forum Membership."
Const strTxtThankYouForRegistering = "Thank-you for registering to use"
Const strTxtIfErrorActvatingMembership = "If you have a problem activating your membership please contact the"


'active_users.asp
'---------------------------------------------------------------------------------
Const strTxtActiveForumUsers = "Active Forum Users"
Const strTxtAddMToActiveUsersList = "Add me to Active Users list"
Const strTxtLoggedIn = "Logged In"
Const strTxtLastActive = "Last Active"
Const strTxtBrowser = "Browser"
Const strTxtOS = "OS"
Const strTxtMinutes = "minutes"
Const strTxtAnnoymous = "Anonymous"



'not_posted.asp
'---------------------------------------------------------------------------------
Const strTxtMessageNotPosted = "Message Not Posted"
Const strTxtDoublePostingIsNotPermitted = "Double posting is not permitted; your message has been posted already."
Const strTxtSpammingIsNotPermitted = "Spamming is not permitted!"
Const strTxtYouHaveExceededNumOfPostAllowed = "You have exceeded the number of posts permitted in the time span.<br /><br />Please try again later."
Const strTxtYourMessageNoValidSubjectHeading = "Your message did not contain a valid subject heading and/or message body."


'active_topics.asp
'---------------------------------------------------------------------------------
Const strTxtActiveTopics = "New Posts"
Const strTxtLastVisitOn = "Last visit on"
Const strTxtLastFifteenMinutes = "Last 15 minutes"
Const strTxtLastThirtyMinutes = "Last 30 minutes"
Const strTxtLastFortyFiveMinutes = "Last 45 minutes"
Const strTxtLastHour = "Last hour"
Const strTxtLastTwoHours = "Last 2 hours"
Const strTxtYesterday = "Yesterday"
Const strTxtNoActiveTopicsSince = "There are no New Topics since"
Const strTxtToDisplay = "to display"
Const strTxtThereAreCurrently = "There are currently"



'pm_check.inc
'---------------------------------------------------------------------------------
Const strTxtNewPMsClickToGoNowToPM = "new Private Message(s).\n\nClick OK to go to your Private Messenger."


'display_forum_topics.inc
'---------------------------------------------------------------------------------
Const strTxtFewYears = "few years"
Const strTxtWeek = "week"
Const strTxtTwoWeeks = "two weeks"
Const strTxtMonth = "month"
Const strTxtTwoMonths = "two months"
Const strTxtSixMonths = "6 months"
Const strTxtYear = "year"



Const strTxtHasBeenSentTo = "has been sent to"
Const strTxtCharactersInYourSignatureToLong = "characters in your signature, you must shorten it to below 200"
Const strTxtSorryYourSearchFoundNoMembers = "Your search found no forum members that match your criteria, try to broaden your search criteria"
Const strTxtCahngeOfEmailReactivateAccount = "If you change your email address you will be sent an email to re-activate your account"
Const strTxtAddToBuddyList = "Add to Buddy List"


'register_mail_confirm.asp
'---------------------------------------------------------------------------------
Const strTxtYourEmailAddressHasBeenChanged = "Your email address has been changed, <br />you will have to re-activate your forum membership before you can use the forum again."
Const strTxtYouShouldReceiveAReactivateEmail = "<strong>Your Forum Membership needs to be re-activated!</strong><br /><br />An activation email will be sent in a few moments to the email address in your profile.<br />Click on the link in this email to re-activate your Forum Membership."


'Preview signature window
'---------------------------------------------------------------------------------
Const strTxtSignaturePreview = "Signature Preview"
Const strTxtPostedMessage = "Posted Message"



'New from version 7
'---------------------------------------------------------------------------------

Const strTxtMemberlist = "Memberlist"
Const strTxtForums = "Forum(s)"
Const strTxtOurUserHavePosted = "Our users have posted"
Const strTxtInTotalThereAre = "In total there are"
Const strTxtOnLine = "online"
Const strTxtWeHave = "We have"
Const strTxtActivateAccount = "Activate Account"
Const strTxtSorryYouDoNotHavePermissionToPostInTisForum = "You do not have permission to post new topics in this forum"
Const strTxtSorryYouDoNotHavePerimssionToReplyToPostsInThisForum = "You do not have permission to reply to posts in this forum"
Const strTxtSorryYouDoNotHavePerimssionToReplyIPBanned = "You can not reply to posts, your IP address or range is not permitted.<br />Please contact the forum administrator if you feel this is in error."
Const strTxtLoginSm = "login"
Const strTxtYourProfileHasBeenUpdated = "Your profile has been updated."
Const strTxtPosted = "Posted:"
Const strTxtBackToTop = "Back to Top"
Const strTxtNewPassword = "New Password"
Const strTxtRetypeNewPassword = "Retype New Password"
Const strTxtRegards = "Regards"
Const strTxtClickTheLinkBelowToUnsubscribe = "If you no-longer wish to receive email notification for this Topic or Forum click on the link below "
Const strTxtPostsPerDay = "posts per day"
Const strTxtGroup = "Group"
Const strTxtLastVisit = "Last Visit"
Const strTxtPrivateMessage = "Private Message"
Const strTxtSorryFunctionNotPermiitedIPBanned = "This function is not available as you are using an IP address or range that is not permitted.<br />Please contact the forum administrator if you feel this is in error."
Const strTxtEmailAddressBlocked = "The email address or domain entered has been blocked by the forum administrator.\n\nPlease use a different email address or email domain."
Const strTxtTopicAdmin = "Topic Admin"
Const strTxtMovePost = "Move Post"
Const strTxtPrevTopic = "Prev Topic"
Const strTxtTheMemberHasBeenDleted = "The Member has been Deleted."
Const strTxtThisPageWasGeneratedIn = "This page was generated in"
Const strTxtSeconds = "seconds."
Const strTxtEditBy = "Edited by"
Const strTxtWrote = "wrote"
Const strTxtEnable = "Enable"
Const strTxtToFormatPosts = "to format post"
Const strTxtFlashFilesImages = "Adobe Flash"
Const strTxtSessionIDErrorCheckCookiesAreEnabled = "A security error has occurred with authentication.<br /><br />Please ensure that all cookies are enabled on your web browser, you are not using a saved or cached copy of the page, and your Firewall/Proxy settings are not hiding or masking your IP address."
Const strTxtName = "Name"
Const strTxtModerators = "Moderators"
Const strTxtMore = "more..."
Const strTxtNewRegSuspendedCheckBackLater = "New registrations are currently suspended, please check back again later."
Const strTxtMoved = "Moved"
Const strTxtNoNameError = "Name   - Enter your name"
Const strTxtHelp = "Help"

'PM system
'---------------------------------------------------------------------------------
Const strTxtPrivateMessenger = "Private Messenger"
Const strTxtUnreadMessage = "Unread Message"
Const strTxtReadMessage = "Read Message"
Const strTxtNew = "new"
Const strTxtYouHave = "You have"
Const strTxtNewMsgsInYourInbox = "unread message(s) in your inbox!"
Const strTxtNoneSelected = "None Selected"
Const strTxtAddBuddy = "Add Buddy"


'active_topics.asp
'---------------------------------------------------------------------------------
Const strTxtSelectMember = "Select Member"
Const strTxtSelect = "Select"
Const strTxtNoMatchesFound = "No matches found"


'active_topics.asp
'---------------------------------------------------------------------------------
Const strTxtLastFourHours = "Last 4 hours"
Const strTxtLastSixHours = "Last 6 hours"
Const strTxtLastEightHours = "Last 8 hours"
Const strTxtLastTwelveHours = "Last 12 hours"
Const strTxtLastSixteenHours = "Last 16 hours"


'permissions
'---------------------------------------------------------------------------------
Const strTxtYou = "You"
Const strTxtCan = "can"
Const strTxtCannot = "cannot"
Const strTxtpostNewTopicsInThisForum = "post new topics in this forum"
Const strTxtReplyToTopicsInThisForum = "reply to topics in this forum"
Const strTxtEditYourPostsInThisForum = "edit your posts in this forum"
Const strTxtDeleteYourPostsInThisForum = "delete your posts in this forum"
Const strTxtCreatePollsInThisForum = "create polls in this forum"
Const strTxtVoteInPOllsInThisForum = "vote in polls in this forum"


'register.asp
'---------------------------------------------------------------------------------
Const strTxtRegistrationDetails = "Registration Details"
Const strTxtProfileInformation = "Profile Information"
Const strTxtForumPreferences = "Forum Preferences"
Const strTxtICQNumber = "ICQ Number"
Const strTxtAIMAddress = "AIM Address"
Const strTxtMSNMessenger = "Windows Live Messenger"
Const strTxtYahooMessenger = "Yahoo Messenger"
Const strTxtOccupation = "Occupation"
Const strTxtInterests = "Interests"
Const strTxtDateOfBirth = "Date of Birth"
Const strTxtNotifyMeOfReplies = "Notify me of replies to posts"
Const strTxtSendsAnEmailWhenSomeoneRepliesToATopicYouHavePostedIn = "Sends an email when someone replies to a topic you have posted in. This can be changed whenever you make a post."
Const strTxtNotifyMeOfPrivateMessages = "Notify me by email when I receive a Private Message"
Const strTxtAlwaysAttachMySignature = "Always attach my signature to posts"
Const strTxtEnableTheWindowsIEWYSIWYGPostEditor = "Enable the WYSIWYG post editor <br /><span class=""smText"">Only browsers that are detected as being Rich Text Enabled will have this feature available when posting.</span>"
Const strTxtTimezone = "Time offset from forum time"
Const strTxtPresentServerTimeIs = "Present server date and time is: "
Const strTxtDateFormat = "Date Format"
Const strTxtDayMonthYear = "Day/Month/Year"
Const strTxtMonthDayYear = "Month/Day/Year"
Const strTxtYearMonthDay = "Year/Month/Day"
Const strTxtYearDayMonth = "Year/Day/Month"
Const strTxtHours = "hours"
Const strTxtDay = "Day"
Const strTxtCMonth = "Month"
Const strTxtCYear = "Year"
Const strTxtRealName = "Real Name"
Const strTxtMemberTitle = "Member Title"


'Polls
'---------------------------------------------------------------------------------
Const strTxtCreateNewPoll = "Create New Poll"
Const strTxtPollQuestion = "Poll&nbsp;Question"
Const strTxtPollChoice = "Poll Choice"
Const strTxtErrorPollQuestion = "Poll Question  - Enter a Question for this Poll"
Const strTxtErrorPollChoice = "Poll Choice  - Enter a least two choices for this Poll"
Const strTxtSorryYouDoNotHavePermissionToCreatePollsForum = "You do not have permission to create polls in this forum"
Const strTxtAllowMultipleVotes = "Allow Multiple Votes in this Poll"
Const strTxtMakePollOnlyNoReplies = "Make Poll only (no replies allowed)"
Const strTxtYourNoValidPoll = "Your Poll did not contain a valid Question or Choices."
Const strTxtPoll = "Poll:"
Const strTxtVote = "Vote"
Const strTxtVotes = "Votes"
Const strTxtCastMyVote = " Cast My Vote"
Const strTxtPollStatistics = "Poll Statistics"
Const strTxtThisTopicIsClosedNoNewVotesAccepted = "This topic is closed, no new votes accepted"
Const strTxtYouHaveAlreadyVotedInThisPoll = "You have already voted in this poll"
Const strTxtThankYouForCastingYourVote = "Thank-you for casting your vote."
Const strsTxYouCanNotNotVoteInThisPoll = "You can not vote in this poll"
Const strTxtYouDidNotSelectAChoiceForYourVote = "Your vote was not cast.\n\nYou did not select a Poll Choice to vote for."
Const strTxtThisIsAPollOnlyYouCanNotReply = "This is a Poll only, you can not post a reply."


'Email Notify
'---------------------------------------------------------------------------------
Const strTxtWatchThisTopic = "Watch This Topic"
Const strTxtUn = "Un-"
Const strTxtWatchThisForum = "Watch This Forum"
Const strTxtYouAreNowBeNotifiedOfPostsInThisForum = "You will now be notified by email of all Posts in this Forum.\n\nTo un-subscribe from this forum click on \'Un-Watch This Forum\' \n or go to your \'Email Notification Subscriptions\' page found in your Member Control Panel."
Const strTxtYouAreNowNOTBeNotifiedOfPostsInThisForum = "You will now not be notified by email of Posts in this Forum.\n\nTo re-subscribe to this forum click on \'Watch This Forum\' \n or go to your \'Email Notification Subscriptions\' page found in your Member Control Panel."
Const strTxtYouWillNowBeNotifiedOfAllReplies = "You will now be notified by email of all Replies in this Topic.\n\nTo un-subscribe from this topic click on \'Un-Watch This Topic\' \n or go to your \'Email Notification Subscriptions\' page found in your Member Control Panel."
Const strTxtYouWillNowNOTBeNotifiedOfAllReplies = "You will now not be notified by email of Replies in this Topic.\n\nTo re-subscribe to this topic click on \'Watch This Topic\'."


'email_messenger.asp
'---------------------------------------------------------------------------------
Const strTxtEmailMessenger = "Email Messenger"
Const strTxtRecipient = "Recipient"
Const strTxtNoHTMLorForumCodeInEmailBody = "Please note that the email is sent in plain text only (no HTML or forum codes).<br /><br />Please include a return email address within the message as the 'from' email address is set to that of the forum."
Const strTxtYourEmailHasBeenSentTo = "Your Email has been sent to"
Const strTxtYouCanNotEmail = "You can not email"
Const strTxtYouDontHaveAValidEmailAddr = "you do not have a valid email address in your forum profile."
Const strTxtTheyHaveChoosenToHideThierEmailAddr = "they have chosen to hide their email address."
Const strTxtTheyDontHaveAValidEmailAddr = "they do not have a valid email address in their profile."
Const strTxtSendACopyOfThisEmailToMyself = "Send a copy of this email to myself"
Const strTxtTheFollowingEmailHasBeenSentToYouBy = "The following email has been sent to you by"
Const strTxtFromYourAccountOnThe = "from the forum your participate in on "
Const strTxtIfThisMessageIsAbusive = "If this message is spam or you find it offensive please contact the webmaster of the forum at the following address"
Const strTxtIncludeThisEmailAndTheFollowing = "Include this email and the following"
Const strTxtMessageSent = "Message sent"



'forum_closed.asp
'---------------------------------------------------------------------------------
Const strTxtForumClosed = "Forums Closed"
Const strTxtSorryTheForumsAreClosedForMaintenance = "The forums are presently closed for maintenance.<br />Please check back again later."


'report_post.asp
'---------------------------------------------------------------------------------
Const strTxtReportPost = "Report Post"
Const strTxtSendReport = "Send Report"
Const strTxtProblemWithPost = "Problem With Post"
Const strTxtPleaseStateProblemWithPost = "Please state below the issue with the post, a copy of the post will be emailed to the forum moderators and/or forum administrators so they can deal with it appropriately."
Const strTxtTheFollowingReportSubmittedBy = "The following report has been submitted by"
Const strTxtWhoHasTheFollowingIssue = "who has the following issue with this post"
Const strTxtToViewThePostClickTheLink = "To view the post then click on the link below"
Const strTxtIssueWithPostOn = "Issue With Post on"
Const strTxtYourReportEmailHasBeenSent = "Your email has been sent to the forum moderators and/or forum administrators so they can deal with it appropriately."


'New from version 7.5
'---------------------------------------------------------------------------------
Const strTxtQuickLogin = "Quick Login"
Const strTxtThisTopicWasStarted = "This topic was started: "
Const strTxtResendActivationEmail = "Resend Activation Email"
Const strTxtNoOfStars = "Number of Stars"
Const strTxtOnLine2 = "Online"
Const strTxtCode = "Code"
Const strTxtCodeandFixedWidthData = "Code and Fixed Width Data"
Const strTxtQuoting = "Quoting"
Const strTxtMyCodeData = "My Code or Fixed Width Data"
Const strTxtQuotedMessage = "Quoted Message"
Const strTxtWithUsername = "With Username"
Const strTxtGo = "Go"
Const strTxtDataBasedOnActiveUsersInTheLastXMinutes = "This data is based on users active over the past twenty minutes"
Const strTxtSoftwareVersion = "Software Version"
Const strTxtForumMembershipNotAct = "Your Forum Membership is not yet activated!"
Const strTxtMustBeRegisteredToPost = "You must be a registered forum member in order to post in this forum."
Const strTxtMemberCPMenu = "Member Control Panel"
Const strTxtYouCanAccessCP = "You can access forum tools and change your Forum Preferences through your "
Const strTxtEditMembersSettings = "Edit Members Settings"
Const strTxtSecurityCodeConfirmation = "Security Code Confirmation (required)"
Const strTxtUniqueSecurityCode = "Security Code"
Const strTxtEnterCAPTCHAcode = "Please enter the Security Code exactly as shown in image format.<br />Cookies must be enabled on your web browser."
Const strTxtErrorSecurityCode = "Security Code  - You must enter the CAPTCHA security code image as shown"
Const strTxtSecurityCodeDidNotMatch = "The Security Code entered did not match that displayed.\n\nA new Security Code has been generated."

'login_user_test.asp
'---------------------------------------------------------------------------------
Const strTxtSuccessfulLogin = "Successful Login"
Const strTxtSuccessfulLoginReturnToForum = "Your login was successful, please wait while you are returned to the Forum"
Const strTxtUnSuccessfulLoginText = "Your login was un-successful due to a session tracking error. <br /><br />Please ensure that cookies are enabled on your web browser and you are not hiding or masking your IP address."
Const strTxtUnSuccessfulLoginReTry = "Click here to retry logging into the Forum."
Const strTxtToActivateYourForumMem = "To activate your forum membership click on the link in the activation email you should have received after registering."

'email_notify_subscriptions.asp
'---------------------------------------------------------------------------------
Const strTxtEmailNotificationSubscriptions = "Email Notification Subscriptions"
Const strTxtSelectForumErrorMsg = "Select Forum - Select a Forum to Subscribe to"
Const strTxtYouHaveNoSubToEmailNotify = "You have no subscriptions to email notification"
Const strTxtThatYouHaveSubscribedTo = "that you have Subscribed to Email Notification"
Const strTxtUnsusbribe = "Unsubscribe"
Const strTxtAreYouWantToUnsubscribe = "Are you sure you want to Un-subscribe from these"



'New from version 7.51
'---------------------------------------------------------------------------------
Const strTxtSubscribeToForum = "Subscribe to Email Notification of New Posts"
Const strTxtSelectForumToSubscribeTo = "Select Forum to Subscribe to"


'New from version 7.7
'---------------------------------------------------------------------------------
Const strTxtOffLine = "Offline"


'New from version 7.8
'---------------------------------------------------------------------------------
Const strTxtConfirmOldPass = "Confirm Old Password"
Const strTxtConformOldPassNotMatching = "The Password Confirmation does not match that held on record for your account.\n\nIf you wish to change your password, please enter your old password correctly."



'New from version 8.0
'---------------------------------------------------------------------------------
Const strTxtSub = "Sub"
Const strTxtHidden = "Hidden"
Const strTxtHidePost = "Hide Post"
Const strTxtAreYouSureYouWantToHidePost = "Are you sure you want to hide this post?"
Const strTxtModeratedPost = "Pre-Approved Post"
Const strTxtYouArePostingModeratedForum = "You are posting in a moderated forum."
Const strTxtBeforePostDisplayedAuthorised = "Your post will first need to be approved by a moderator before it will be viewable to others on the forum."
Const strTxtHiddenTopics = "Moderated Topics"
Const strTxtVerifiedBy = "Verified By"
Const strTxtYourEmailHasChanged = "Your email address for"
Const strTxtPleaseUseLinkToReactivate = "has changed, please use the link below to re-activate your account"
Const strTxtToday = "Today"
Const strTxtPreviewPost = "Preview Post"
Const strTxtEmailNotify = "Email Notify me of Replies"
Const strTxtAvatarUpload = "Avatar Upload"
Const strTxtClickOnEmoticonToAdd = "Click on the emoticon you would like to add to your message."
Const strTxtUpdatePost = "Update Post"
Const strTxtShowSignature = "Show Signature"
Const strTxtQuickReply = "Quick Reply"
Const strTxtCategory = "Category"
Const strTxtReverseSortOrder = "Reverse Sort Order"
Const strTxtSendPM = "Send Private Message"
Const strTxtSearchKeywords = "Search Keywords"
Const strTxtSearchbyKeyWord = "Search by Keyword"
Const strTxtSearchbyUserName = "Search by Member Name (Optional)"
Const strTxtMatch = "Match"
Const strTxtSearchOptions = "Search Options"
Const strTxtCtrlApple = "(Hold down the 'control' or 'apple' key to select more than one)"
Const strTxtFindPosts = "Find Posts"
Const strTxtAndNewer = "and Newer"
Const strTxtAndOlder = "and Older"
Const strTxtAnyDate = "Any Date"
Const strTxtNumberReplies = "Number of Replies"
Const strTxtExactMatch = "Exact Match"
Const strTxtSearhExpiredOrNoPermission = "Either the search has expired or you do not have permission to run this search."
Const strTxtCreateNewSearch = "Create New Search"
Const strTxtNoSearchResultsFound = "No Search Results Found"
Const strTxtSearchError = "Search Error"
Const strTxtSearchWordLengthError = "One or more of the words used in your search criteria is below 3 characters"
Const strTxtIPSearchError = "Your IP address has exceeded the number of allowed searches<br /><br />Please allow 30 seconds between searches or log into the forum"
Const strTxtResultsIn = "results in"
Const strTxtSecounds = "seconds"
Const strTxtFor = "for"
Const strTxtThisSearchWasProcessed = "This search was processed"
Const strTxtError = "Error"
Const strTxtReply = "Reply"
Const strTxtClose = "Close"
Const strTxtActiveStats = "Active Stats"
Const strTxtInformation = "Information"
Const strTxtCommunicate = "Communicate"
Const strTxtDisplayResultsAs = "Display Results As"
Const strTxtViewPost = "View Post"
Const strTxtPasswordRequiredViewPost = "Password Required to View Post"
Const strTxtNewestPostFirst = "Newest Post First"
Const strTxtOldestPostFirst = "Oldest Post First"
Const strTxtMessageIcon = "Message Icon"
Const strTxtSkypeName = "Skype Name"
Const strTxtLastPostDetailNotHiddenDetails = "Please note:- Last Post details don't include details of hidden posts."
Const strTxtOriginallyPostedBy = "Originally posted by"
Const strTxtViewingTopic = "Viewing Topic"
Const strTxtViewingIndex = "Viewing Index"
Const strTxtForumIndex = "Forum Index"
Const strTxtIndex = "Index"
Const strTxtViewing = "Viewing"
Const strTxtSearchingForums = "Searching Forums"
Const strTxtSearchingFor = "Searching For"
Const strTxtWritingPrivateMessage = "Writing Private Message"
Const strTxtViewingPrivateMessage = "Viewing Private Message"
Const strTxtEditingPost = "Editing Post"
Const strTxtWritingReply = "Writing Reply"
Const strTxtWritingNewPost = "Writing New Post"
Const strTxtCreatingNewPost = "Creating New Poll"
Const strTxtWhatsGoingOn = "What's Going On?"
Const strTxtLoadNewCode = "Refresh Image"
Const strTxtApprovePost = "Approve Post"
Const strTxtSuspendUser = "Suspend Member"
Const strTxtAdminNotes = "Admin/Moderator Notes"
Const strTxtAdminNotesAbout = "These are notes left by Admin and Moderators, about this member, such as any warnings the member has been given (max 250 characters)"
Const strTxtAge = "Age"
Const strTxtUnknown = "Unknown"
Const strTxtSuspended = "Suspended"
Const strTxtEmailNewUserRegistered = "The following new member has registered on "
Const strTxtToActivateTheNewMembershipFor = "To activate the new membership for "
Const strTxtNewMemberActivation = "New Member Activation"
Const strTxtEmailYouCanNowUseOnceYourAccountIsActivatedTheForumAt = "Your login details can be found below. Once the Forum Admin approves your new account you will be able to post new messages and reply to existing ones on the"
Const strTxtYouAdminNeedsToActivateYourMembership = "<strong>Your Forum Membership needs to be activated by the Forum Admin!</strong>"
Const strTxtEmailYourForumMembershipIsActivatedThe = "Your forum membership is now active and you can now post new messages and reply to existing ones on the"
Const strTxtTheAccountIsNowActive = "The Account is now Active!!"
Const strTxtErrorOccuredActivatingTheAccount = "An error has oocured while attempting to activate the account"
Const strTxtMustBeLoggedInAsAdminActivateAccount = "You need to be logged in as the Forum Admin in order to activate new forum members.<br />Once logged in as Forum Admin, click the link in the activation email again."
Const strTxtTodaysBirthdays = "Today's Birthdays"
Const strTxtCalendar = "Calendar"
Const strTxtEventDate = "Event Date"
Const strTxtEvent = "Event"
Const strTxtCalendarEvent = "Calendar Event"
Const strTxtLast = "Last"
Const strTxtRSS = "RSS Feed"
Const strTxtNewPostFeed = "New Post Feed"
Const strTxtLastTwoDays = "Last 2 Days"
Const strTxtThisRSSFileIntendedToBeSyndicated = "This file is intended to be viewed in an RSS News Aggregator or syndicated to another site."
Const strTxtCurrentFeedContent = "Current Feed Content"
Const strTxtSyndicatedForumContent = "Syndicated Forum Content"
Const strTxtSubscribeNow = "Subscribe To RSS Feed Now!"
Const strTxtSubscribeWithWebBasedRSS = "with web-based news readers. Click your choice below"
Const strTxtWithOtherReaders = "with a RSS News Reader installed on your computer"
Const strTxtSelectYourReader = "Select Your Reader"
Const strTxtThisIsAnXMLFeedOf = "This is an XML content feed of"
Const strTxtDirectLinkToThisPost = "Direct Link To This Post"
Const strTxtWhatIsAnRSSFeed = "What is an RSS Feed?"


'New from version 8.02
'---------------------------------------------------------------------------------
Const strTxtSecurityCodeDidNotMatch2 = "The Security Code entered did not match that displayed."




'New from version 8.05
'---------------------------------------------------------------------------------
Const strTxtPleaseDontForgetYourPassword = "Please do not forget your password as it has been encrypted in our database and we cannot retrieve it for you. However, should you forget your password you can request a new one sent to your email address."
Const strTxtActivationEmail = "Activation Email" 
Const strTxtTopicReplyNotification = "Topic Reply"
Const strTxtUserNameOrEmailAddress = "Username or Email Address"
Const strTxtAnonymousMembers = "Anonymous Member(s)"
Const strTxtGuests = "Guest(s)"
Const strTxtNewPosts = "New Posts"
Const strTxtNoNewPosts = "no new posts"
Const strTxtFullReplyEditor = "Full Reply Editor"


'New from version 9
'---------------------------------------------------------------------------------
Const strTxtForumHome = "Forum Home"
Const strTxtNewMessages = "New Messages"
Const strTxtFAQ = "FAQ"
Const strTxtUnAnsweredTopics = "Un-Answered Topics"
Const strTxtShowPosts = "Show Posts"
Const strTxtModeratorTools = "Moderator Tools"
Const strTxtResyncTopicPostCount = "Resync Forum Stats"
Const strTxtAdminControlPanel = "Admin Control Panel"
Const strTxtAdvancedSearch = "Advanced Search"
Const strTxtLockTopic = "Lock Topic"
Const strTxtHideTopic = "Hide Topic"
Const strTxtShowTopic = "Show Topic"
Const strTxtTopicOptions = "Topic Options"
Const strTxtForumOptions = "Forum Options"
Const strTxtFindMembersPosts = "Find Members Posts"
Const strTxtMembersProfile = "Members Profile"
Const strTxtVisitMembersHomepage = "Visit Members Homepage"
Const strTxtFirstPage = "First Page"
Const strTxtLastPage = "Last Page"
Const strTxtPostOptions = "Post Options"
Const strTxtBlockUsersIP = "Block Poster's IP"
Const strTxtCreateNewTopic = "Create New Topic"
Const strTxtNewPoll = "New Poll"
Const strTxtControlPanel = "Control Panel"
Const strTxtSubscriptions = "Subscriptions"
Const strTxtMessenger = "Messenger"
Const strTxtBuddyList = "Buddy List"
Const strTxtProfile2 = "Profile"
Const strTxtSubscribe = "Subscribe"
Const strTxtMultiplePages = "Multiple Pages"
Const strTxtCurrentPage = "Current Page"
Const strTxtRefreshPage = "Refresh Page"
Const strTxtAnnouncements = "Announcements"
Const strTxtHiddenTopic = "Moderated Topic"
Const strTxtHot = "Hot"
Const strTxtLocked = "Locked"
Const strTxtNewPost = "New Post"
Const strTxtPoll2 = "Poll"
Const strTxtSticky = "Sticky"
Const strTxtForumPermissions = "Forum Permissions"
Const strTxtForumWithSubForums = "Forum with Sub Forum(s)"
Const strTxtPostNewTopic2 = "Post New Topic"
Const strTxtViewDropDown = "View Drop Down"
Const strTxtFull = "full"
Const strNotYetRegistered = "Not Yet Registered?"
Const strTxtNewsletterSubscription = "Newsletter Subscription"
Const strTxtSignupToRecieveNewsletters = "Signup to receive Newsletters from " 
Const strTxtNewsBulletins = "Blog Posts"
Const strTxtPublished = "Published"
Const strTxtStartDate = "Start Date"
Const strTxtEndDate = "End Date"
Const strTxtNotRequiredForSingleDateEvents = "not required for single date events"
Const strTxtIn = "in"
Const strTxtGender = "Gender"
Const strTxtMale = "Male"
Const strTxtFemale = "Female"



Const strTxtFileAlreadyUploaded = "You have already uploaded a file of the same name"
Const strTxtSelectOrRenameFile = "Please select another file, or rename the file and try again."
Const strTxtAllotedFileSpaceExceeded = "You have exceeded the allotted file space of"
Const strTxtDeleteFileOrImagesUingCP = "Please delete any unused files or images using the File Manager available in your Member Control Panel"



'File Manager
Const strTxtFileManager = "File Manager"
Const strTxtFileName = "File Name"
Const strTxtFileSize = "File Size"
Const strTxtFileType = "File Type"
Const strTxtFileExplorer = "File Explorer"
Const strTxtFileProperties = "File Properties"
Const strTxtFilePreview = "File Preview"
Const strTxtAllocatedFileSpace = "Allocated File Space"
Const strTxtYouHaveUsed  = "You have used"
Const strTxtFromYour = "from your"
Const strTxtOfAllocatedFileSpace = "of allocated file space"
Const strTxtYourFileSpaceIs = "Your file space is"
Const strTxtDownloadFile = "Download File"
Const strTxtNewUpload = "New Upload"
Const strTxtDeleteFile = "Delete File"
Const strTxtRenameFile = "Rename File"
Const strTxtAreYouSureDeleteFile = "Are you sure you want to delete the file"
Const strTxtNoFileSelected = "There is no file selected"
Const strTxtTheFileNowDeleted = "The file has been deleted"
Const strTxtYourFileHasBeenSuccessfullyUploaded = "Your file has been successfully uploaded to the server"
Const strTxtSelectUploadType = "Select Upload Type"
Const strTxtYouTube = "YouTube"
Const strTxtUploadFolderEmpty = "Upload Folder Empty"




'New from version 9.04
'---------------------------------------------------------------------------------

Const strTxtAutologinOnlyAppliesToSession = "Only applies to this session, you will need to select not to auto-login next time you login to the forum"
Const strTxtViewUnreadPost = "View First Unread Post"



'New from version 9.51
'---------------------------------------------------------------------------------

Const strTxtPendingApproval = "Pending Approval"
Const strTxtThatRequiresApproval = "that requires approval."

Const strTxtMovieProperties = "Movie Properties"
Const strTxtMovieType = "Movie Type"
Const strTxtYouTubeFileName = "YouTube File Name"
Const strTxtFlashMovieURL = "Flash Movie URL"


'New from version 9.52
'---------------------------------------------------------------------------------
Const strTxtThroughTheirForumProfileAtLinkBelow = "through their forum profile at the link below"
Const strTxtYouCanNotEmailTisTopicToAFriend = "You can not email this topic to a friend"
Const strTxtToReplyPleaseEmailContact = "To reply to this email contact"
Const strTxtInsertMovie = "Insert Movie"

'New from version 9.54
'---------------------------------------------------------------------------------
Const strTxtTheEmailFailedToSendPleaseContactAdmin = "The email failed to send. Please contact the Forum Administrator with the error message below."
Const strTxtFindMember = "Find Member"
Const strTxtSearchForTopicsThisMemberStarted = "Search for Topics This Member Started"
Const strTxtMemberName2 = "Member Name"
Const strTxtSearchTimeoutPleaseNarrowSearchTryAgain = "A timeout expired while performing the search. Please narrow your search criteria and try again"


'New from version 9.55
'---------------------------------------------------------------------------------
Const strTxtTheFileFailedTheSecurityuScanAndHasBeenDeleted = "The file failed the security scan and has been deleted as it may contain malicous code"


'New from version 9.56
'---------------------------------------------------------------------------------
Const strTxtShareTopic = "Share Topic"
Const strTxtPostThisTopicTo = "Post this Topic to"

'New from version 9.61
'---------------------------------------------------------------------------------
Const strTxtSponsor = "Sponsored Links"

'New from version 9.64
'---------------------------------------------------------------------------------
Const strTxtIPAddress = "IP Address"

'New from version 9.65
'---------------------------------------------------------------------------------
Const strTxtTranslate = "Translate"

'New from version 9.66
'---------------------------------------------------------------------------------
Const strTxtConfirmEmail = "Confirm Email"
Const strTxtErrorConfirmEmail = "Confirm Email \t- The Email Fields do not match"

'New from version 9.67
'---------------------------------------------------------------------------------
Const strTxtThereMayAlsoBeOtherMessagesPostedOn = "There may also be other replies, but you will not receive any more notifications until you visit the forum again"
Const strTxtWarningYourSessionHasExpiredRefreshPageFormDataWillBeLost = "WARNING\nYour session has expired! Click 'Refresh\' on your browser to reload the page.\n** Any form data you have entered may be lost! **"


'New from version 9.70
'---------------------------------------------------------------------------------
Const strTxtNoFollowAppliedToAllLinks = "NoFollow is applied to all links from this forum"


'New from version 9.71
'---------------------------------------------------------------------------------
Const strTxtViewIn = "View in"
Const strTxtMoble = "Mobile"
Const strTxtClassic = "Classic"


'New from version 10
'---------------------------------------------------------------------------------
Const strTxtStatus = "Status"
Const strTxtTheEmailAddressEnteredIsInvalid = "The email address entered is invalid"
Const strTxtMostUsersEverOnlineWas = "Most users ever online was"
Const strTxtTypeTheNameOfMemberInBoxBelow = "Type in members name or part of the members name you wish to find, then click the 'Search' button. For example, if you wish to find a member called 'wiz123' type in 'wiz' into the search box to find all members names beginning with the name 'wiz'."
Const strTxtSelectNameOfMemberFromDropDownBelow = "Below is a list of members who match your search criteria, select the member you are looking for and click the 'Select' button to insert this members name into the form this box opened from."
Const strTxtCharacters = "characters"
Const strTxtMxLFailedLoginAttemptsMade = "More than the maximum failed login attempts have been made on this account.<br />Please enter your details again, including security code."
Const strTxtNumberOfPoints = "Number of Points"
Const strTxtPoints = "Points"
Const strTxtPasswordNotComplex = "Your Password does not meet the Password Complexity requirements.\n\n Your Password must contain at least; 1 Uppercase Character, 1 Lowercase Character, 1 Number and 1 Non-Alphanumeric Symbol."
Const strTxtLadderGroup = "Ladder Group"
Const strTxtNone = "None"
Const strTxtRealNameError = "Real Name  - You must enter your Real Name"
Const strTxtLocationError = "Location  - You must enter your Location"
Const strTxtNotRequired = "Not Required"
Const strTxtRating = "Rating"
Const strTxtTopicRating = "Topic Rating"
Const strTxtAverage = "Average"
Const strTxtRateTopic = "Rate Topic"
Const strTxtYouHaveAlreadyRatedThisTopic = "Our records show you have already voted for this Topic!"
Const strTxtYouCanNotRateATopicYouStarted = "You can not rate a Topic that you started!"
Const strTxtThankYouForRatingThisTopic = "Thank-you for rating this Topic."
Const strTxtExcellent = "Excellent"
Const strTxtPoor = "Poor"
Const strTxtGood = "Good"
Const strTxtTerrible = "Terrible"
Const strTxtRateThisTopicAs = "Click here to; Rate this Topic as"
Const strTxtSortOrder = "Sort Order"
Const strTxtMost = "Most"
Const strTxtHighestRating = "Highest Rating"
Const strTxtBy1 = "by"
Const strTxtMembers2 = "Members"
Const strTxtEvents = "Events"
Const strTxtYouAreOnlyPermittedToEditPostWithin = "You are only permitted to Edit Post within"
Const strTxtChat = "Chat"
Const strTxtOnlineMembers = "Online Members"
Const strTxtReason = "Reason"
Const strTxtYourMessageWasRejectedByTheSpamFilters = "Your message was rejected by the SPAM Filters"
Const strYouMustEnterYour = "You must enter your"
Const strTxtViewUnreadPost1 = "View Unread Post"
Const strTxtSetAsAnswer = "Set As Answer"
Const strTxtUnSetAsAnswer = "Un-Set As Answer"
Const strTxtExternalLinkTo = "External Link To"
Const strTxtYouMustBeARegisteredMemberAndPostAReplyToViewMessage = "You must be a registered member and post a reply in this topic to view the hidden content within this post"
Const strTxtHideContent = "Hide Content"
Const strTxtPostContentHiddenUntilReply = "Hidden Post Content, revealed when Member Posts Reply"


Const strTxtThanks = "Thanks"
Const strTxtThanked = "Thanked"
Const strTxtYouMustHaveAnActiveMemberAccount  = "You must be an Active Member Account"
Const strTxtYouCanNotThankYourself = "You cannot thank yourself"
Const strTxtYouHaveAlreadySaidThanksForThisPost = "You have already said thanks for this Post"
Const strTxtHasBeenThankedForTheirPost = "has been thanked for their Post"
Const strTxtYourPmInboxIsFullPleaseDeleteOldPMs = "Your Private Message Inbox is full!\nPlease delete any old or unwanted Private Messages."
Const strTxtShare = "Share"
Const strTxtShareThisPageOnTheseSites ="Share this page on these sites"

Const strTxtFacebook = "Facebook"
Const strTxtLinkedIn = "LinkedIn"
Const strTxtTwitter = "Twitter"

Const strTxtAnswer = "Answer"
Const strTxtResolution = "Resolution"
Const strTxtOfficialResponse = "Official Response"



'New from version 10.02
'---------------------------------------------------------------------------------
Const strTxtYouDoNotHavePermissionToViewPorfiles = "You do not have permission to view member profiles."
Const strTxtSearchRobots = "Search Robot(s)"


'New from version 10.10 beta
'---------------------------------------------------------------------------------
Const strTxtToViewTheDetailsOf = "To view details of"
Const strTxtBeforeApprovingClickTheLinkBelow = "*before* approving, click on the link below"
Const strTxtToViewInfoOnIPAddressOf = "To view information on the IP Address of "

Const strTxtIfYourAreSureThat = "If you are sure that"
Const strTxtIsASmammerSubmitToSpamDatabase = "is a Spammer, click the link below to submit their details to StopForumSpam Database"

Const strTxtInformationGiveFoundInSpamDatabaseAndRejected = "Information given in your Registration has been flagged by 3rd party Anti-Spam Databases and has therefore been rejected."
Const strSubmitAsSpammer = "Submit As Spammer"
Const strTxtSearchEngine = "Search Engine"
Const strTxtSearchEngines = "Search Engine(s)"
Const strTxtSpamCleaner = "Spam Cleaner"
Const strTxtSubmit = "Submit"
Const srtTxtIAgreeToThe = "I agree to the"

Const strTxtHour = "hour"
Const strTxtMinute = "minute"
Const strTxtSecond = "second"
Const strTxtAgo = "ago"

Const strTxtErrorTerms = "Terms  - You must accept the Forum Rules and Policies"


'New from version 10.11
'---------------------------------------------------------------------------------
Const strTxtCookiePrivacyNotice = "Cookie Privacy Notice"


'New from version 10.12
'---------------------------------------------------------------------------------
Const strTxtSubmittedToStopForumSpam = "Submitted to StopForumSpam"

'New from version 10.18
'---------------------------------------------------------------------------------
Const strTxtRegistrationToThisForumIsNotPermittedFromYourCountry = "Registration to this forum is not permitted from your country."


'New from version 11.08
'---------------------------------------------------------------------------------
Const strTxtSpoiler = "Spoiler"
Const strTxtShow = "Show"
Const strTxtHide = "Hide"
Const strTxtSpoilerContent = "Spoiler Content"
Const strTxtHiddenSpolierContentShowenWhehnButtonClicked = "Hidden Spoiler Content, revealed when 'Show' button is clicked"


'New from version 12.00
'---------------------------------------------------------------------------------
Const strTxtYourInformation = "Your Information"
Const strTxtIPAddresses = "IP Addresses"
Const strTxtBelowYouCanFindIPAddressesRecordedInWhenPosting = "Below you can find a list of IP addresses that were recorded when you posted within the forum"
Const strTxtMemberData = "Member Data"
Const strTxtViewYourPersonalData  = "View Your Personal Data"
Const strTxtHaveALookAtThePersonalDataTheForumHasOnYou  = "Have a look and see what Personal Data the Forum has on you."
%>