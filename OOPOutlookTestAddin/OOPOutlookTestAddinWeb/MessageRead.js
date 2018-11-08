'use strict';

(function () {

    var item;

    Office.initialize = function () {
        item = Office.context.mailbox.item;
        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
            // Insert data in the top of the body of the composed 
            // item.
            prependItemBody();
        });
    }

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
                            '<b>Greetings!</b>',
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
                            'Greetings!',
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

    // Writes to a div with id='message' on the page.
    function write(message) {
        document.getElementById('message').innerText += message;
    }

})();