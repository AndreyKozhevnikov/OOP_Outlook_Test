var item;

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
        Office.context.mailbox.item.body.setSelectedDataAsync("<b>Hello World Bodies!</b>", { coercionType: "html" });
    }

})();

/*

// Get the body type of the composed item, and prepend data
// in the appropriate data type in the item body.
function prependItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Prepend data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of prependAsync.
                    item.body.prependAsync(
                        '<b>Greetings from message bodies!</b>',
                        {
                            coercionType: Office.CoercionType.Html,
                            asyncContext: { var3: 1, var4: 2 }
                        },
                        function (asyncResult) {
                            if (asyncResult.status ==
                                Office.AsyncResultStatus.Failed) {
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.prependAsync(
                        'Greetings from message read!',
                        {
                            coercionType: Office.CoercionType.Text,
                            asyncContext: { var3: 1, var4: 2 }
                        },
                        function (asyncResult) {
                            if (asyncResult.status ==
                                Office.AsyncResultStatus.Failed) {
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
            }
        });

}

*/

// Writes to a div with id='message' on the page.
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
