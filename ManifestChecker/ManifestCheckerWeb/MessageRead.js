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
         * This is some code!
         * 
         * Entries are written as jQuery statements. More recipients can be added to the recipients variable above.
         * 
         * Functionality can be added to any HTML in MessageRead.html so long as the element has a guid.
         * The jQuery functionality is copy-paste, and any html elements can be added in place.
         * 
        */

        $("#fauxButton").click(function () {
            Office.context.mailbox.item.to.setAsync(recipient);
            Office.context.mailbox.item.subject.setAsync("This is a subject!")
            Office.context.mailbox.item.body.setAsync("This is done from<br /> Read Initialize button test!", { coercionType: Office.CoercionType.Html });
        });

        $("#button2").click(function () {

            
            Office.context.mailbox.item.to.setAsync(recipient);
            Office.context.mailbox.item.subject.setAsync("This is a subject also!")
            Office.context.mailbox.item.body.setAsync("This is done from Read Initialize button 2 test!", { coercionType: Office.CoercionType.Html });
        });

        $("#realButton").click(function () {
            Office.context.mailbox.item.to.setAsync(recipient);
            Office.context.mailbox.item.subject.setAsync("This is a subject also but from a real button!")
            Office.context.mailbox.item.body.setAsync("This is done from Read Initialize button 1asgdfasdfasd test!", { coercionType: Office.CoercionType.Html });
        });

    }
    
})();


