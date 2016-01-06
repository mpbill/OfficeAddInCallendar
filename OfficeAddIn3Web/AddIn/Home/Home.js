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
            Office.context.document.getData()
            Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
                function (result) {
                    var events = [];
                    for (var i = 0; i < result.value.length; i++) {
                        var row = result.value[i];
                        var dateField = row[0];
                        var timeField = row[1];
                        var date = moment.fromOADate(dateField);
                        var time = moment.fromOADate(timeField);
                        var timeArr = time.toArray();
                        date.add(timeArr[3], 'hours');
                        date.add(timeArr[4], 'minutes');
                        var s = date.toISOString();
                        events.push(s);

                    }
                    var unique = [];
                    var uniqueWithCount = [];
                    for (var i = 0; i < events.length;i++) {
                        var a = events[i];
                        if ($.inArray(a, unique)== -1) {
                            var count = 0;
                            for (var j = 0; j < events.length; j++) {
                                if (a == events[j]) {
                                    count++;
                                }
                            }
                            unique.push(a);
                            uniqueWithCount.push({
                                datetime: a,
                                count: count
                            });
                        }
                    }
                    for (var i = 0; i < uniqueWithCount.length; i++) {
                        var newEvent = new Object();
                        newEvent.title = "title";
                        newEvent.start = uniqueWithCount[i].datetime;
                        newEvent.allDay = false;
                        $('#calendar').fullCalendar('renderEvent', newEvent);

                    }
                }
            );
        } else {
            app.showNotification('Error:', 'Reading selection data is not supported by this host application.');
        }
    }
})();