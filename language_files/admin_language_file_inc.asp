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


'Topic admin
'---------------------------------------------------------------------------------
Const strTxtUpdateTopic = "Update Topic"
Const strTxtTopicNotFoundOrAccessDenied = "Either the topic can not be found or you do not have permission to access this page"
Const strTxtMoveTopic = "Move Topic"
Const strTxtDeletePoll = "Delete Poll"
Const strTxtAreYouSureYouWantToDeleteThisPoll = "Are you sure you want to delete this poll?"



'move_post_form.asp
'---------------------------------------------------------------------------------
Const strTxtSelectTheForumYouWouldLikePostIn = "Select the Forum you would like this post to be in"
Const strTxtMovePostErrorMsg = "Forum\t- Please select a Forum to move the Post to"
Const strTxtSelectForumClickNext = "Step 1: - Select the Forum you would like to move the Post to.<br />Then click on the Next button."
Const strTxtSelectTopicToMovePostTo = "Step 2: - Select the Topic to move the Post to from the last 400 Topics in this forum. Alternatively type the name of a new Topic in the box at the bottom to place the Post in a new Topic."
Const strTxtSelectTheTopicYouWouldLikeThisPostToBeIn = "Select the Topic you would like this post to be in"
Const strTxtOrTypeTheSubjectOfANewTopic = "Or, type the Subject of a new Topic"



'pop_up_IP_blocking.asp
'---------------------------------------------------------------------------------
Const strTxtIPBlocking = "IP Blocking"
Const strTxtBlockedIPList = "Blocked IP Address List"
Const strTxtYouHaveNoBlockedIpAddesses = "You have no blocked IP address"
Const strTxtRemoveIP = "Remove IP"
Const strTxtBlockIPAddressOrRange = "Block IP Address or Range"
Const strTxtIpAddressRange = "IP Address/Range"
Const strTxtBlockIPAddressRange = "Block IP Address/Range"
Const strTxtBlockIPRangeWhildcardDescription = "The * wildcard character can be used to block IP ranges. <br />eg. To block the range '200.200.200.0 - 255' you would use '200.200.200.*'"
Const strTxtErrorIPEmpty = "IP Address/Range \t- Enter a an IP address or range to block"


'register.asp - in admin mode
'---------------------------------------------------------------------------------
Const strTxtAdminModeratorFunctions = "Admin and Moderator Functions"
Const strTxtUserIsActive = "Account Activated"
Const strTxtDeleteThisUser = "Delete this member?"
Const strTxtCheckThisBoxToDleteMember = "Check this box to delete this member, this cannot be undone."
Const strTxtNumberOfPosts = "Number of posts"
Const strTxtNonRankGroup = "Non Ladder Group"




'Topic admin
'---------------------------------------------------------------------------------
Const strTxtUpdateForum = "Update Forum"
Const strTxtForumNotFoundOrAccessDenied = "Either the forum can not be found or you do not have permission to access this page"





'New from version 7.02
'---------------------------------------------------------------------------------
Const strTxtShowMovedIconInLastForum = "Show moved icon in last forum"

'New from version 8.0
'---------------------------------------------------------------------------------
Const strTxtIfYouAreShowingTopic = "This will only hide or display this topic in the Topic Index, you still need to approve individual posts within the topic"
Const strTxtShowHiddenTopicsSince = "Show Pending Approval Topics/Posts Since"
Const strTxtHiddenTopicsPosts = "Pending Approval Topics/Posts"
Const strTxtNoHiddenTopicsPostsSince = "There are no Pending Approval Topics/Posts "


'New from version 10.0
'---------------------------------------------------------------------------------
Const strTxtMinPosts = "Min. Posts"
Const strTxtRankLadderGroup = "Ladder Group"


'New from version 10.04
'---------------------------------------------------------------------------------
Const strTxtDeleteMembersPosts = "Delete Members Posts"
Const strTxtBlockMembersIPpAddress = "Block Members IP Address"
Const strTxtSubmitToStopForumSpam = "Submit Member to StopForumSpam"
Const strTxtRemoveMembersSignature = "Remove Members Signature"
Const strTxtDeleteMembersPrivateMessages = "Delete Members Private Messages"

Const strTxtMembersDetailsFoundInStopForumSpamDatabase = "Member's details are in the StopForumSpam Database"
Const strTxtMembersLsstLoggedIpInBlockList = "Member's last logged in IP Address is in the IP Block List"
Const strTxtNoPPostsCanBeFound = "No Posts can be found for this Member"
Const strTxtNoPrivateMessagesCanBeFound = "No Private Message can be found for this Member"
Const strTxtThisMemberIsSuispended = "This Member is Suspended"
Const strTxtCleanedSpammerSpamCleanerOn = "cleaned Spammer using the Spam Cleaner on"

Const strTxtBlockedFromSpamFilter = "Blocked from Spam Cleaner"

Const strTxtSuspendingMembersBetterThanDeleteing = "Suspended Members is better than deleting them as they are not able to join your forum again using the same email address or username."
Const strTxtIfMemCreatedLessThanXXPostsDeleteThem = "If the Member has created less than 50 posts you can delete all their posts"
Const strTxtBlockIPAddressFromRegPostInForum = "Block this Members IP Address from registering or posting in your forum"
Const strTxtSubmitToStopForumSpamDatabaseToStopSpammer = "Submit this Members details to the StopForumSpam Database, to prevent this spammer from spamming other forums"
Const strTxtDeleteMembersPrivateMessagesSentToAndFrom = "Delete the Members Private Messages sent both to the Member and to others"
%>