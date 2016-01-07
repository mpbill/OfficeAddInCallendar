/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $('#loadevents').click(getDataFromSelection);
            $('#insertsampledata').click(insertSampleData);
            // page is now ready, initialize the calendar...
        });
    };
    function insertSampleData() {
        var sampleDataArr = [];
        var keys = [];
        for(var propName in sampleData[0]) {
            keys.push(propName);
        }
        sampleDataArr.push(keys);
        for(var i=1;i<sampleData.length;i++) {
            var row = [];
            for(var j=0;j<keys.length;j++) {
                row.push(sampleData[i][keys[j]]);
            }
            sampleDataArr.push(row);
        }
        Office.context.document.setSelectedDataAsync(sampleDataArr,{coercionType: Office.CoercionType.Matrix}, function (result) {
            var i = 0;
        });
        
    }

    function doNothing() {
        
    }
    function listListToDictList(arrArr) {
        var dictList = [];
        var header = arrArr[0];
        for (var i = 1; i < arrArr.length; i++) {
            var row = {}
            for(var j=0;j<header.length;j++) {
                row[header[j]] = arrArr[i][j];
            }
            dictList.push(row);
        }
        return dictList;
    }
    function findMin(dictList, key) {
        var min = dictList[0][key];
        for(var i=1;i<dictList.length;i++) {
            if(dictList[i][key]<min) {
                min = dictList[i][key];
            }
        }
        return min;
    }
    function findMax(dictList, key) {
        var max = dictList[0][key];
        for (var i = 1; i < dictList.length; i++) {
            if (dictList[i][key] > max) {
                max = dictList[i][key];
            }
        }
        return max;
    }
    function MakeRange(start, stop, step) {
        var arr = [];
        for(var i=start;i<stop;i+=step) {
            arr.push(i);
        }
        return arr;
    }
    function DateWithinDates(start, stop, test) {
        var range = moment.range(start, stop);
        return range.contains(test);
    }
    function MakeData(result) {
        var table = listListToDictList(result.value);
        var min = findMin(table, "DateTime");
        var max = findMax(table, "DateTime") + (3 / 24);
        var minDay = Math.floor(min);
        var maxDay = Math.floor(max);
        var dayDiff = maxDay - minDay;
        var dateArr = [];
        for (var i = 0; i <= dayDiff; i++) {
            dateArr.push(minDay + i);
        }
        

        var dateArrStrings = [];
        for (var i = 0; i < dateArr.length; i++) {
            dateArrStrings.push(moment.fromOADate(dateArr[i]).format("MM/DD/YYYY"));
        }
        var hrRange = MakeRange(8, 20, 1);
        var hrRangeStrings = [];
        hrRange.forEach(function (hour) {
            hrRangeStrings.push(moment({ hours: hour }).format("h a") + " - " + moment({hours:hour+1}).format("h a"));
        });
        var dataSet = [];
        for(var x=0;x<dateArr.length;x++) {
            for(var y=0;y<hrRange.length;y++) {
                var dataPoint = [x, y];
                var overlapCount = 0;
                var dateOne = dateArr[x] + hrRange[y] / 24;
                var dateTwo = dateOne + 1 / 24;
                for(var i=0;i<table.length;i++) {
                    var date3 = table[i]["DateTime"];
                    var date4 = date3 + table[i]["TimeSpan"] / 24;
                    var from = Math.max(dateOne, date3);
                    var to = Math.min(dateTwo, date4);
                    if (from <= to) {
                        overlapCount++;
                    }
                }
                dataPoint.push(overlapCount);
                dataSet.push(dataPoint);
                

                 
            }
        }
        var toReturn = {
            Data: dataSet,
            xAxis: dateArrStrings,
            yAxis: hrRangeStrings
        }
        return toReturn;



    }

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
                function (result) {
                    var dataSet = MakeData(result);
                    $('#container').highcharts({
                        chart: {
                            type: 'heatmap',
                            marginTop: 40,
                            marginBottom: 80,
                            plotBorderWidth: 1
                        },


                        title: {
                            text: 'Schedulings Per Day Per Hour'
                        },

                        xAxis: {
                            categories: dataSet.xAxis,
                            title:"Date"
                        },

                        yAxis: {
                            categories: dataSet.yAxis,
                            title: "Time"
                        },

                        colorAxis: {
                            min: 0,
                            minColor: '#FFFFFF',
                            maxColor: Highcharts.getOptions().colors[0]
                        },

                        legend: {
                            align: 'right',
                            layout: 'vertical',
                            margin: 0,
                            verticalAlign: 'top',
                            y: 25,
                            symbolHeight: 280
                        },

                        tooltip: {
                            formatter: function () {
                                return '<b>' + this.point.value + " Installs Occuring On <br/>" + moment(this.series.xAxis.categories[this.point.x]).add(this.series.yAxis.categories[this.point.y], 'hours').format("LLLL")+ "</b>";
                            }
                        },

                        series: [{
                            name: 'Sales per employee',
                            borderWidth: 1,
                            data: dataSet.Data,
                            dataLabels: {
                                enabled: true,
                                color: '#000000'
                            }
                        }]

                    });
                });
    }
})();