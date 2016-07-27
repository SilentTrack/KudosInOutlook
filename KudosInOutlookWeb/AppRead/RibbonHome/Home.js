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

var months;
var kudosData;
var kudosInfos;

function InitPage() {
    $.ajax({
        //url: "https://localhost:44372/api/KudosService/5/?KudosReceiver=" + Office.context.mailbox.userProfile.emailAddress,
        url: "https://localhost:44372/api/KudosService/5/?KudosReceiver=" + "junxw@microsoft.com",
        success: function (result) {
            months = result.months;
            kudosData = result.kudosPerMonth;
            kudosInfos = result.kudosInfos;
            $(".data-my").html(kudosInfos.length);
            var itemId = Office.context.mailbox.item.itemId;
            var ewsId = Office.context.mailbox.convertToEwsId(itemId, Office.MailboxEnums.RestVersion.v2_0);
            $("#testLink").click(function () {
                Office.context.mailbox.displayMessageForm(ewsId);
            });

            $(".data-total").html(result.totalKudos);
            ShowChart();
            ShowHistory();
        },
        error: function () {
            months = result.months;
            kudosData = result.kudosPerMonth;
            totalKudos = result.totalKudos;
        }
    });
}

function ShowHistory()
{
    var list = document.getElementById("historyList");
    while (list.hasChildNodes()) {
        list.removeChild(list.firstChild);
    }
    var totalEntities = kudosInfos.length;
    if (totalEntities > 5)
    {
        totalEntities = 5;
    }

    var liOld = $(".templateLi");
    var liTemplate = liOld.clone();
    liOld.remove();
    for (var i = 0; i < totalEntities; ++i)
    {
        var li = liTemplate.clone();
        li.removeClass("templateLi").show().find(".data-sender").html(kudosInfos[i].senderName);
        li.find(".data-thread").html(kudosInfos[i].subject);
        li.find(".comment").html(kudosInfos[i].additionalMessage);
        li.find(".time").html(kudosInfos[i].sentDate);

        var itemID = kudosInfos[i].itemID;
        var ewsId = Office.context.mailbox.convertToEwsId(itemID, Office.MailboxEnums.RestVersion.v2_0);
        li.find(".data-thread").data("ewsId", ewsId);
        li.find(".data-thread").click(function () {
            var id = $(this).data("ewsId");
            Office.context.mailbox.displayMessageForm(id);
        });
        $(".data-ul").append(li);
    }
}

function ShowChart() {
    Chart.defaults.global.defaultFontSize = 11;
    var ctx = $("#dataChart");
    var myChart = new Chart(ctx, {
        type: 'bar',
        data: {
            //labels: ["Mar", "Apr", "May", "Jun", "Jul"], //Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec
            labels: months,
            datasets: [{
                label: 'Kudos/Month',
                //data: [5, 0, 10, 12, 7],
                data: kudosData,
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
