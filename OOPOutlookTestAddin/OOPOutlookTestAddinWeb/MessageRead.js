﻿'use strict';

(function () {

    // Initialization will fail if Office.initialize
    // does not have a function assigned to it
    Office.initialize = function (reason) {
        $(document).ready(function () {
            loadItemProps(Office.context.mailbox.item);
        });
    };

    function loadItemProps(item) {
        // Get the table body element
        var tbody = $('.prop-table');

        // Add a row to the table for each message property
        tbody.append(makeTableRow("Id", item.itemId));
        tbody.append(makeTableRow("Subject", item.subject));
        tbody.append(makeTableRow("Message Id", item.internetMessageId));
        tbody.append(makeTableRow("From", item.from.displayName + " &lt;" +
            item.from.emailAddress + "&gt;"));
    }

    function makeTableRow(name, value) {
        return $("<tr><td><strong>" + name +
            "</strong></td><td class=\"prop-val\"><code>" +
            value + "</code></td></tr>");
    }

})();