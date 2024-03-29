//Uncommented: Message read displays, not bodies.
//Commented:   Message bodies displays, nothing from Read.
//*
Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel) {
        // Do Excel-specific initialization (for example, make add-in task pane's
        // appearance compatible with Excel "green").
    }
    if (info.platform === Office.PlatformType.PC) {
        // Make minor layout changes in the task pane.
    }
    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
    
});
Office.context.mailbox.item.body.prependAsync("Hello World Bodies! 1");
//*/
var item;


//Commented:   Message Read displays, but not bodies
//Uncommented: 
//*
(function () {
    "use strict";

    var messageBanner;


    Office.initialize = function () {
        item = Office.context.mailbox.item;
        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
            // Insert data in the top of the body of the composed 
            // item.
            //prependItemBody();

            
        });
        Office.context.mailbox.item.body.prependAsync("Hello World Bodies! 2");
        
    }

})();

//*/

function write(message) {
    //Office.context.mailbox.item.to.setAsync("This is to whom");

    Office.context.mailbox.item.body.setAsync("<b>Hello World Button!</b>", { coercionType: "html" });
}


var e1 = document.getElementById("button1");
if (e1.addEventListener)
    e1.addEventListener("click", write("This is a test of body changes"), false)
else if (e1.attachEvent)
    e1.attachEvent('onclick', write("This is a test of body changes"))

document.getElementById("button1").onclick = write("This is a test of body changes");


Office.context.mailbox.item.subject.setAsync("This is the subject line from bodies");