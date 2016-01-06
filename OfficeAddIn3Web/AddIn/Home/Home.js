/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            // page is now ready, initialize the calendar...
            $('#loadevents').click(getDataFromSelection);
                $('#calendar').fullCalendar({
                    // put your options and callbacks here
                    defaultView:'basicWeek'
                });
                

           
            
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        if (Office.context.document.getSelectedDataAsync) {
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
                function (result) {
                    for(var i=0;i<result.value.length;i++) {
                        var date = moment.fromOADate(result.value[i]);
                    }
                }
            );
        } else {
            app.showNotification('Error:', 'Reading selection data is not supported by this host application.');
        }
    }
})();