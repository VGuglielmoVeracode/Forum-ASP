/* *******************************************************
Software: Web Wiz Forums
Info: http://www.webwizforums.com
Copyright: (C)2001-2012 Web Wiz Ltd. All rights reserved
******************************************************* */

var lines = 50;
var chatPointer = 0;
var chatter = new Array();
var message  = document.getElementById('message');
var chatBox = document.getElementById('ChatRoomBox');
var sessionID = document.getElementById('SID');
var memberPointer   = 0;
var members = new Array();
var membersBox = document.getElementById('ChatMembersBox');
var hasFocus = true;

//Initial server check time
var requestTimer = setTimeout('chatRead();', 1000);

//Setup window focus and blur
window.onblur = function () {hasFocus = false;};
window.onfocus = function () {hasFocus = true;};
message.onfocus = function () {hasFocus = true;};
	
//Send write response to server
function httpRequestWrite(url, post) {
  	xmlHttpWriteResponse = false;
  	
  	if (window.XMLHttpRequest){
		xmlHttpWriteResponse = new XMLHttpRequest();
	}else if (window.ActiveXObject){
		xmlHttpWriteResponse = new ActiveXObject('Msxml2.XMLHTTP');
		if (! xmlHttpWriteResponse){
			xmlHttpWriteResponse = new ActiveXObject('Microsoft.XMLHTTP');
		}
	}
  	
  	if (!xmlHttpWriteResponse) return false;
  	xmlHttpWriteResponse.onreadystatechange = alertWrite;
  	if (post == null) {
    		xmlHttpWriteResponse.open('GET', url, true);
    		xmlHttpWriteResponse.send(null);
  	} else {
    		xmlHttpWriteResponse.open('POST', url, true);
    		xmlHttpWriteResponse.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
    		xmlHttpWriteResponse.send(post);
  	}
}

//Process write reponse
function alertWrite() {
  	try {
    		if ((xmlHttpWriteResponse.readyState == 4 || xmlHttpWriteResponse.readyState=='complete') && (xmlHttpWriteResponse.status == 200)) {
			parse(xmlHttpWriteResponse.responseText);
		}
  	} catch(e) {
  	}
}

//Key events
function keyup(evtKey) {
  	if (window.event) key = window.event.keyCode;
  	else if (evtKey) key = evtKey.which;
  	else return true;
  	if (key == 13) {
  		chatWrite();
  	}else{
  		chatRead(true) 
	}
}

//Write chat
function chatWrite() {
  	httpRequestWrite('chat_server.asp', 'SID=' + escape(sessionID.value) + '&writeMsg=' + escape(message.value));
  	message.value = '';
  	clearTimeout(requestTimer);
  	requestTimer = setTimeout('chatRead();', 2000); //writing message 2 secound to read again
}

//Send read reponse to server
function httpRequestRead(url, post) {
	
  	xmlHttpReadResponse = false;
  	if (window.XMLHttpRequest){
		xmlHttpReadResponse = new XMLHttpRequest();
	}else if (window.ActiveXObject){
		xmlHttpReadResponse = new ActiveXObject('Msxml2.XMLHTTP');
		if (! xmlHttpReadResponse){
			xmlHttpReadResponse = new ActiveXObject('Microsoft.XMLHTTP');
		}
	}
	
	  if (!xmlHttpReadResponse) return false;
	  xmlHttpReadResponse.abort();
	  xmlHttpReadResponse.onreadystatechange = alertRead;
	  if (post == null) {
	    	xmlHttpReadResponse.open('GET', url, true);
	    	xmlHttpReadResponse.send(null);
	  } else {
	    	xmlHttpReadResponse.open('POST', url, true);
	    	xmlHttpReadResponse.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	    	xmlHttpReadResponse.send(post);
	  }
}

//Process read reponse
function alertRead() {
  	try {
    		if ((xmlHttpReadResponse.readyState == 4 || xmlHttpReadResponse.readyState=='complete') && (xmlHttpReadResponse.status == 200)) {
	      		parse(xmlHttpReadResponse.responseText);
	      		clearTimeout(requestTimer); requestTimer = setTimeout('chatRead();', 2000); //good response read again in 2 seounds
	      	}else{
	      		clearTimeout(requestTimer); requestTimer = setTimeout('chatRead();', 4000); //bad response wait 4 seounds to read again
		}
	} catch(e) {
	    	
	    	clearTimeout(requestTimer); requestTimer = setTimeout('chatRead();', 4000); //bad response wait 4 seounds to read again
 	}
}

//Chat read
function chatRead(typing) {
  	if (typing == true){
  		httpRequestRead('chat_server.asp?p=' + chatPointer, 'SID=' + escape(sessionID.value) + '&t=true');
	}else if (hasFocus == false){
		httpRequestRead('chat_server.asp?p=' + chatPointer, 'SID=' + escape(sessionID.value) + '&a=false');
	}else{
  		httpRequestRead('chat_server.asp?p=' + chatPointer, 'SID=' + escape(sessionID.value));
  	}
  	clearTimeout(requestTimer);
	requestTimer = setTimeout('chatRead();', 3000); //reload in 3 secounds
}


//Output chat and member reponse to page
function responseWrite() {
  	html = '';
  	i = 0;
  	while ((i < lines) && (i < chatPointer)) {
    		h = chatPointer-i;
    		if (chatter[h]) html = chatter[h] + html;
    		i++;
  	}
  	chatBox.innerHTML = html;
  	chatBox.scrollTop = chatBox.scrollHeight;
  	
  	memhtml = '';
  	i = 0;
  	while (i < memberPointer) {
    		h = memberPointer-i;
    		if (members[h]) memhtml = members[h] + memhtml;
    		i++;
  	}
  	membersBox.innerHTML = memhtml;
}

//Setup chat display
function chat(msgID, uid, type, username, chatMsg) {
  	if (type == 'info') {
    		chatter[msgID] = '<span class="chatMessage">' + chatMsg + '</span><br />';
  	} else if (type == 'warn') {
    		chatter[msgID] = '<span class="chatAlert">' + chatMsg + '</span><br />';
  	} else {
    		if (username != '') {
      			username += ':';
      			spaces = 5 - username.length;
      			for (j = 0; j < spaces; j++) username += "&nbsp;";
      			username += ' ';
    		}
    		chatter[msgID] = '<span class="chatMember">' + username + '</span>' + chatMsg + '<br />'; ;
  	}
  	if (msgID > chatPointer) {
    		chatPointer = msgID;
  	}
}

//Parse srever response
function parse(serverRepsonse) {
  	if (serverRepsonse != '') {
    		serverRepsonse = unescape(serverRepsonse);
    		eval(serverRepsonse);
    		responseWrite();
  	}
}

//Setup member display
function online(aryPos, uid, joinTime, username, msgType) {
	username = username.replace(/ /g, '&nbsp;');
  	if (msgType == 'typing') {
    		members[aryPos] = '<img src="forum_images/chat_usr_typing.png" title="' + username + '" /> <a href="member_profile.asp?PF=' + uid + '&SID=' + escape(sessionID.value) + '" class="chatListMember">' + username + '</a><br />';
  	} else if (msgType == 'na') {
    		members[aryPos] = '<img src="forum_images/chat_usr_na.png" title="' + username + '" /> <a href="member_profile.asp?PF=' + uid + '&SID=' + escape(sessionID.value) + '" class="chatListMember" style="font-weight:normal;">' + username + '</a><br />';
  	} else if (msgType == 'left') {
    		members[aryPos] = '<img src="forum_images/chat_usr_na.png" title="' + username + '" /> <a href="member_profile.asp?PF=' + uid + '&SID=' + escape(sessionID.value) + '" class="chatListMember" style="font-weight:normal;">' + username + '</a><br />';
  	} else {
    		members[aryPos] = '<img src="forum_images/chat_usr.png" title="' + username + '" /> <a href="member_profile.asp?PF=' + uid + '&SID=' + escape(sessionID.value) + '" class="chatListMember">' + username + '</a><br />';
    	}
    		
    	memberPointer = aryPos;
}