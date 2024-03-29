﻿Office.initialize = function () {
}

// Helper function to add a status message to the info bar.
function statusUpdate(icon, text) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
    type: "informationalMessage",
    icon: icon,
    message: text,
    persistent: false
  });
}

function defaultStatus(event) {
  statusUpdate("icon16" , "Hello World!");
}

function prependItemBody(event) {
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
                        '<b>Greetings from the function file!</b>',
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