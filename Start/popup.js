(function () {
"use strict";

    Office.initialize = function() {
        $(document).ready(function () {
            // Assign the handler to the OK button.
            $('#ok-button').click(sendStringToParentPage);
        });
    }

    // Create the OK button handler
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
}());