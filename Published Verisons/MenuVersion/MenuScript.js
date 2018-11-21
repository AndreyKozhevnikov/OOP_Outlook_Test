var item;

var recipient = [{
    "displayName": "OOPIT Level 1",
    "emailAddress": "oopit@psu.edu"
}];

(function () {
    "use strict";
    

    Office.initialize = function () {
        item = Office.context.mailbox.item;
        $(document).ready(function () {
            
        });

        /*
         * Entries are written as jQuery statements. More recipients can be added to the recipients variable above.
         * 
         * Functionality can be added to any HTML in MessageRead.html so long as the element has a guid.
         * The jQuery functionality is copy-paste, and any html elements can be added in place.
         * 
        */
		
		//Useful Functions
		function hideAll(showCancel){
			$("#Menu_Main").hide();
			$("#Menu_Incident").hide();
			$("#Menu_Request").hide();
			$("#Menu_Question").hide();
			
			$("#Button_Cancel").hide();
			
			if(showCancel == true){
				$("#Button_Cancel").show();
			}
		};
		
		function composeMessage(Subject, Body){
			Office.context.mailbox.item.to.setAsync(recipient);
            Office.context.mailbox.item.subject.setAsync(Subject);
            Office.context.mailbox.item.body.setAsync(Body, { coercionType: Office.CoercionType.Html });
			
			//Once we have made a message, we can close the task pane
			Office.context.ui.closeContainer();
		};
		
		//Default build behavior
		hideAll();
		$("#Menu_Main").show();
		
		//Behavior of buttons in the main menu
		$("#Button_Incident").click(function (){
			hideAll(true);
			$("#Menu_Incident").show();
		});
		$("#Button_Request").click(function (){
			hideAll(true);
			$("#Menu_Request").show();
		});
		$("#Button_Question").click(function (){
			hideAll(true);
			$("#Menu_Question").show();
		});
		
		//Cancel button behavior
		$("#Button_Cancel").click(function () {
			hideAll(false)
			$("#Menu_Main").show();
		});
		
		//Behavior of buttons in the Incident menu
		$("#Button_Incident_Passwords").click(function (){
			composeMessage("I have a question regarding passwords", "Here is my question");
		});
		
		//Behavior of buttons in the Request menu
		
		//Behavior of buttons in the Question menu
		
		
		
    }
    
})();


