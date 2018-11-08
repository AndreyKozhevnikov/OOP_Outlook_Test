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
            //Office.context.mailbox.item.to.setAsync("This is to whom");

            
        });

    }
    
})();