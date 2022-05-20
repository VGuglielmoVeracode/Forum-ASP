/* *******************************************************
Software: Web Wiz Forums
Info: http://www.webwizforums.com
Copyright: (C)2001-2012 Web Wiz Ltd. All rights reserved
******************************************************* */


//Funtion to check or uncheck input boxes
function checkAll(boxName){

	changeState('input');
	
	//chkAllBoxState = 
	
	function changeState(tag){
	
		//Set up what we are looking for
		var classElements = new Array();
		var els = document.getElementsByTagName(tag);
		var elsLen = els.length;
		var pattern = new RegExp('(^|\\s)' + boxName + '(.*\)');

		//Loop through all input elemnets on page
		for (i = 0, j = 0; i < elsLen; i++) {
		
			//Use regular experssion to look for input elements to change
			if (pattern.test(els[i].name)){
				
				//Check the state of the check box
				if ((document.getElementById('chkAll' + boxName).checked) && (els[i].disabled == false)) {
					els[i].checked = true;
				}else{
					els[i].checked = false;
				}
				j++;
			}
		}
	}
}