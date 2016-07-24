/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

/// <reference path="../App.js" />
var xhr;
var serviceRequest;

(function () {
    "use strict";

    // The Office initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            InitPage();
        });
    };
})();

function InitPage() {
    var itemId = Office.context.mailbox.item.itemId;
    var ewsId = Office.context.mailbox.convertToEwsId(itemId, Office.MailboxEnums.RestVersion.v2_0);
    $("#testLink").click(function () {
        Office.context.mailbox.displayMessageForm(ewsId);
    });
    ShowChart();
}

function ShowChart() {
    Chart.defaults.global.defaultFontSize = 11;
    var ctx = $("#dataChart");
    var myChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: ["Mar", "Apr", "May", "Jun", "Jul"], //Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec
            datasets: [{
                label: 'Kudos/Month',
                data: [5, 0, 10, 12, 7],
                backgroundColor: [
                    'rgba(255, 99, 132, 0.2)',
                    'rgba(54, 162, 235, 0.2)',
                    'rgba(255, 206, 86, 0.2)',
                    'rgba(75, 192, 192, 0.2)',
                    'rgba(153, 102, 255, 0.2)'
                ],
                borderColor: [
                    'rgba(255,99,132,1)',
                    'rgba(54, 162, 235, 1)',
                    'rgba(255, 206, 86, 1)',
                    'rgba(75, 192, 192, 1)',
                    'rgba(153, 102, 255, 1)'
                ],
                borderWidth: 1
            }]
        },
        options: {
            scales: {
                yAxes: [{
                    ticks: {
                        beginAtZero: true
                    }
                }]
            },
            legend: {
                labels: {
                    fontSize: 11,
                    boxWidth: 25
                }
            }
        }
    });
}
