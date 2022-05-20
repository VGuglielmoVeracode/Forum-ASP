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


'pm_welcome.asp
'---------------------------------------------------------------------------------
Const strTxtToYourPrivateMessenger = "to your Private Messenger"
Const strTxtPmIntroduction = "With this Private Messenger you can send and receive Messages to other Forum Members. Messages you send will be viewable only by your recipients."
Const strTxtInboxStatus = "Inbox Status"
Const strTxtGoToYourInbox = "Go to your inbox"
Const strTxtNoNewMsgsInYourInbox = "You have no new messages in your inbox."
Const strTxtYourLatestPrivateMessageIsFrom = "Your latest Private Message is from"
Const strTxtSentOn = "sent on"
Const strTxtPrivateMessengerOverview = "Private Messenger Overview"
Const strTxtInboxOverview = "Just like your Email 'inbox', the incoming Private Messages Inbox stores any Private Messages that have been sent to yourself, from here you can view or delete any messages you have received."
Const strTxtOutboxOverview = "This is where outgoing Private Messages are stored. Messages remain in your Outbox until the recipient has deleted the message."
Const strTxtBuddyListOverview = "The buddy list works a little like your email address book. You can add or delete forum members to your Buddy List for quick reference. You can also block those members you don't wish to receive Private Messages from."
Const strTxtNewMsgOverview = "From this option you can compose new Private Messages and send them to other forum members."

'pm_inbox.asp
'---------------------------------------------------------------------------------

Const strTxtInbox = "Inbox"
Const strTxtNewPrivateMessage = "New Private Message"
Const strTxtNoPrivateMessages = "You have no Private Messages in your"
Const strTxtRead = "Read"
Const strTxtMessageTitle = "Message Title"
Const strTxtMessageFrom = "Message From"
Const strTxtDate = "Date"
Const strTxtBlock = "block"
Const strTxtSentBy = "Sent by"
Const strTxtDeletePrivateMessageAlert = "Are you sure you want to delete the Private Message(s)?"
Const strTxtPrivateMessagesYouCanReceiveAnother = "Private Messages, you can receive another"
Const strTxtOutOf = "from a total of"
Const strTxtPreviousPrivateMessage = "Previous Private Message"
Const strTxtMeassageDeleted = "The Private Message(s) have been deleted from your"

'pm_message.asp
'---------------------------------------------------------------------------------
Const strTxtSorryYouDontHavePermissionsPM = "Sorry, you don't have permission to view this private message"
Const strTxtYouDoNotHavePermissionViewPM = "You do not have permission to view this Private Message."
Const strTxtNotificationReadPM = "Read Private Message Notification"
Const strTxtReplyToPrivateMessage = "Reply To Private Message"
Const strTxtAddToBuddy = "Add To Buddy List"
Const strTxtThisIsToNotifyYouThat = "This is to notify you that"
Const strTxtHasReadPM = "has read the Private Message"
Const strTxtYouSentToThemOn = "you sent to them on"


'pm_new_message_form.asp
'---------------------------------------------------------------------------------
Const strTxtSendNewMessage = "Send New Message"
Const strTxtPostMessage = "Post Message"
Const strTxtEmailNotifyWhenPMIsRead = "Email Notify me when message is read"
Const strTxtToUsername = "To&nbsp;Username"
Const strSelectFormBuddyList = "or select from Buddy List"
Const strTxtNoPMSubjectErrorMsg = "Subject \t\t- Enter a Subject for your new Private Message"
Const strTxtNoToUsernameErrorMsg = "To Username \t- Enter or Select a Username to send you Private Message to"
Const strTxtNoPMErrorMsg = "Message \t\t- Enter a Private Message to send"
Const strTxtSent = "Sent"

'pm_new_message.asp
'---------------------------------------------------------------------------------
Const strTxtAPrivateMessageHasBeenPosted = "A Private Message has been posted for you on"
Const strTxtClickOnLinkBelowForPM = "Click on the link below to go to the forum to view the Private Message"
Const strTxtNotificationPM = "Private Message Notification"
Const strTxtTheUsernameCannotBeFound = "The Username you have entered can not be found."
Const strTxtYourPrivateMessage = "Your Private Message"
Const strTxtHasNotBeenSent = "has not been sent!"
Const strTxtAmendYourPrivateMessage = "Amend your Private Message"
Const strTxtReturnToYourPrivateMessenger = "Return to your Private Messenger"
Const strTxtYouAreBlockedFromSendingPMsTo = "You are blocked from sending Private Messages to"
Const strTxtHasExceededMaxNumPPMs = "has exceeded the maximum number of Private Messages they are allowed to receive"
Const strTxtHasSentYouPM = "has sent you a Private Message with the following subject"
Const strTxtToViewThePrivateMessage = "To view the Private Message"


'pm_buddy_list.asp
'---------------------------------------------------------------------------------
Const strTxtNoBuddysInList = "You have no Buddies in your Buddy List"
Const strTxtDeleteBuddyAlert = "Are you sure you want to Delete this Buddy from your Buddy List?"
Const strTxtNoBuddyErrorMsg = "Member Name \t- Enter a Forum Member to add to your Buddy List"
Const strTxtBuddy = "Buddy"
Const strTxtDescription = "Description"
Const strTxtContactStatus = "Contact Status"
Const strTxtThisPersonCanNotMessageYou = "This person can not message you"
Const strTxtThisPersonCanMessageYou = "This person can message you"
Const strTxtAddNewBuddyToList = "Add New Buddy to List"
Const strTxtMemberName = "Member Name"
Const strTxtAllowThisMemberTo = "Allow this member to"
Const strTxtMessageMe = "Message me"
Const strTxtNotMessageMe = "Not message me"
Const strTxtHasNowBeenAddedToYourBuddyList = "has now been added to your Buddy List"
Const strTxtIsAlreadyInYourBuddyList = "is already in your Buddy List"
Const strTxtUserCanNotBeFoundInDatabase = "can not be found in the database.\n\nCheck you havn\'t misspelled the Members Username"



Const strTxtOutbox = "Outbox"
Const strTxtMessageTo = "Message To"
Const strTxtMessagesInOutBox = "Messages remain in your Outbox until the recipient deletes the message."

'New from version 7.02
'---------------------------------------------------------------------------------
Const strTxtYourInboxIs = "Your inbox is"
Const strTxtEmailThisPMToMe = "Email this Private Message to myself"
Const strTxtEmailBelowPrivateEmailThatYouRequested = "Below is the copy of the Private Message you requested emailed to yourself"
Const strTxtAnEmailWithPM = "An Email containing the Private Message has"
Const strTxtBeenSent = "been sent to your email address"
Const strTxtNotBeenSent = "not been sent to you as an error has occurred"
Const strTxtSelected = "Selected"


'New from version 8
'-----------------------------------------------------------------
Const strTxtYouAreOnlyPerToSend = "You are only permitted to send"
Const strTxtYouHaveExceededLimit = "Private Messages per hour. You have exceeded this limit"


'New from version 9
'-----------------------------------------------------------------
Const strTxtNewMessage = "New Message"



'New from version 10
'-----------------------------------------------------------------
Const strTxtYourOutboxIs = "Your outbox is"
Const strTxtPrivateMessagesYouCanSendAnother = "Private Messages, you can send another"
Const strTxtYouAreOnlyPerToSendAMaximum = "You are only permitted to send a maximum"
Const strTxtPMsYouHaveExceededLimit = "Private Messages. You have exceeded this limit"
Const strTxtToSendFutherPMsYouWillNeedToDelete = "To send further Private Messages you will need to delete unwanted Private Messages from your Outbox"

%>