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

        $("#fauxButton").click(function () {
            Office.context.mailbox.item.to.setAsync(recipient);
            Office.context.mailbox.item.subject.setAsync("This is a subject!")
            Office.context.mailbox.item.body.setAsync("This is done from Read Initialize button test!");
        });

        $("#button2").click(function () {
            Office.context.mailbox.item.to.setAsync(recipient);
            Office.context.mailbox.item.subject.setAsync("This is a subject also!")
            Office.context.mailbox.item.body.setAsync("This is done from Read Initialize button 2 test!");
        });
    }
    
})();


