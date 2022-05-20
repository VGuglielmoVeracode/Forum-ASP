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




'Dimension variables
Dim saryEmoticons(42,3)	 'If you add more emoticons increase the first number to the number of emoticons you have in the array below


saryEmoticons(1,1) = "Smile"			'Emoticon Name
saryEmoticons(1,2) = "[:)]"			'Forum code
saryEmoticons(1,3) = "smileys/smiley1.gif"	'URL/path to smiley

saryEmoticons(2,1) = "Tongue"
saryEmoticons(2,2) = "[:P]"
saryEmoticons(2,3) = "smileys/smiley17.gif"

saryEmoticons(3,1) = "Wink"
saryEmoticons(3,2) = "[;)]"
saryEmoticons(3,3) = "smileys/smiley2.gif"

saryEmoticons(4,1) = "Cry"
saryEmoticons(4,2) = "[:^(]"
saryEmoticons(4,3) = "smileys/smiley19.gif"

saryEmoticons(5,1) = "Big smile"
saryEmoticons(5,2) = "[:D]"
saryEmoticons(5,3) = "smileys/smiley4.gif"

saryEmoticons(6,1) = "LOL"
saryEmoticons(6,2) = "[LOL]"
saryEmoticons(6,3) = "smileys/smiley36.gif"

saryEmoticons(7,1) = "Dead"
saryEmoticons(7,2) = "[xx(]"
saryEmoticons(7,3) = "smileys/smiley11.gif"

saryEmoticons(8,1) = "Embarrassed"
saryEmoticons(8,2) = "[:$]"
saryEmoticons(8,3) = "smileys/smiley9.gif"

saryEmoticons(9,1) = "Confused"
saryEmoticons(9,2) = "[:s]"
saryEmoticons(9,3) = "smileys/smiley5.gif"

saryEmoticons(10,1) = "Clap"
saryEmoticons(10,2) = "[=D&gt;]"
saryEmoticons(10,3) = "smileys/smiley32.gif"

saryEmoticons(11,1) = "Angry"
saryEmoticons(11,2) = "[:x]"
saryEmoticons(11,3) = "smileys/smiley7.gif"

saryEmoticons(12,1) = "Ouch"
saryEmoticons(12,2) = "[8(]"
saryEmoticons(12,3) = "smileys/smiley18.gif"

saryEmoticons(13,1) = "Star"
saryEmoticons(13,2) = "[:*:]"
saryEmoticons(13,3) = "smileys/smiley10.gif"

saryEmoticons(14,1) = "Shocked"
saryEmoticons(14,2) = "[:o]"
saryEmoticons(14,3) = "smileys/smiley3.gif"

saryEmoticons(15,1) = "Sleepy"
saryEmoticons(15,2) = "[|)]"
saryEmoticons(15,3) = "smileys/smiley12.gif"

saryEmoticons(16,1) = "Unhappy"
saryEmoticons(16,2) = "[:(]"
saryEmoticons(16,3) = "smileys/smiley6.gif"

saryEmoticons(17,1) = "Approve"
saryEmoticons(17,2) = "[:^:]"
saryEmoticons(17,3) = "smileys/smiley14.gif"

saryEmoticons(18,1) = "Cool"
saryEmoticons(18,2) = "[8D]"
saryEmoticons(18,3) = "smileys/smiley16.gif"

saryEmoticons(19,1) = "Clown"
saryEmoticons(19,2) = "[:o)]"
saryEmoticons(19,3) = "smileys/smiley8.gif"

saryEmoticons(20,1) = "Evil Smile"
saryEmoticons(20,2) = "[}:)]"
saryEmoticons(20,3) = "smileys/smiley15.gif"

saryEmoticons(21,1) = "Disapprove"
saryEmoticons(21,2) = "[:V:]"
saryEmoticons(21,3) = "smileys/smiley13.gif"

saryEmoticons(22,1) = "Stern Smile"
saryEmoticons(22,2) = "[:|]"
saryEmoticons(22,3) = "smileys/smiley22.gif"

saryEmoticons(23,1) = "Thumbs Up"
saryEmoticons(23,2) = "[:Y:]"
saryEmoticons(23,3) = "smileys/smiley20.gif"

saryEmoticons(24,1) = "Thumbs Down"
saryEmoticons(24,2) = "[:N:]"
saryEmoticons(24,3) = "smileys/smiley21.gif"

saryEmoticons(25,1) = "Geek"
saryEmoticons(25,2) = "[:-B]"
saryEmoticons(25,3) = "smileys/smiley23.gif"

saryEmoticons(26,1) = "Ermm"
saryEmoticons(26,2) = "[:[]"
saryEmoticons(26,3) = "smileys/smiley24.gif"

saryEmoticons(27,1) = "Question"
saryEmoticons(27,2) = "[:?:]"
saryEmoticons(27,3) = "smileys/smiley25.gif"

saryEmoticons(28,1) = "Pinch"
saryEmoticons(28,2) = "[&gt;&lt;]"
saryEmoticons(28,3) = "smileys/smiley26.gif"

saryEmoticons(29,1) = "Heart"
saryEmoticons(29,2) = "[L]"
saryEmoticons(29,3) = "smileys/smiley27.gif"

saryEmoticons(30,1) = "Broken Heart"
saryEmoticons(30,2) = "[%(]"
saryEmoticons(30,3) = "smileys/smiley28.gif"

saryEmoticons(31,1) = "Wacko"
saryEmoticons(31,2) = "[8-}]"
saryEmoticons(31,3) = "smileys/smiley29.gif"

saryEmoticons(32,1) = "Pig"
saryEmoticons(32,2) = "[:@)]"
saryEmoticons(32,3) = "smileys/smiley30.gif"

saryEmoticons(33,1) = "Hug"
saryEmoticons(33,2) = "[&gt;:D&lt;]"
saryEmoticons(33,3) = "smileys/smiley31.gif"

saryEmoticons(34,1) = "Censored"
saryEmoticons(34,2) = "[XXX]"
saryEmoticons(34,3) = "smileys/smiley35.gif"

saryEmoticons(35,1) = "Yin Yang"
saryEmoticons(35,2) = "[%]"
saryEmoticons(35,3) = "smileys/smiley33.gif"

saryEmoticons(36,1) = "Nuke"
saryEmoticons(36,2) = "[!]"
saryEmoticons(36,3) = "smileys/smiley34.gif"

saryEmoticons(37,1) = "Exclamation"
saryEmoticons(37,2) = "[!]"
saryEmoticons(37,3) = "smileys/smiley37.gif"

saryEmoticons(38,1) = "Lamp"
saryEmoticons(38,2) = "[*]"
saryEmoticons(38,3) = "smileys/smiley38.gif"

saryEmoticons(39,1) = "Sick"
saryEmoticons(39,2) = "[+o(]"
saryEmoticons(39,3) = "smileys/smiley39.gif"

saryEmoticons(40,1) = "Party"
saryEmoticons(40,2) = "[<:o)]"
saryEmoticons(40,3) = "smileys/smiley40.gif"

saryEmoticons(41,1) = "Beer"
saryEmoticons(41,2) = "[beer]"
saryEmoticons(41,3) = "smileys/smiley41.gif"

saryEmoticons(42,1) = "Handshake"
saryEmoticons(42,2) = "[shake]"
saryEmoticons(42,3) = "smileys/smiley42.gif"

'If you add more emoticons don't forget to increase the number in the Dim statement at the top!
%>