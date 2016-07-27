/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

/// <reference path="../App.js" />
var xhr;
var serviceRequest;
//var serviceBaseUrl = "https://kudosservice.azurewebsites.net";
var serviceBaseUrl = "https://localhost:44372";

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
    $("#footer").hide();
    $(".thumbnail-item-template").hide;
    QueryKudosRequest();
}

function QueryKudosRequest() {
    var itemID = Office.context.mailbox.item.itemId;
    $(".data-receiver").html(Office.context.mailbox.item.sender.displayName);
    $.ajax({
        url: serviceBaseUrl + "/api/KudosService?ItemID=" + encodeURIComponent(itemID),
        success: function (result) {
            var totalSenders = result.senders.length;
            if (totalSenders == 0) {
                $(".prompt-sec").show();
                $(".thumbnail-sec").hide();
            }
            else {
                $(".prompt-sec").hide();
                $(".thumbnail-sec").show();
                $(".thumbnail-item-template").hide();
                $(".data-count").html(totalSenders);
                $(".thumbnail-list").html();
                var liOld = $(".thumbnail-item-template");
                for (var i = 0; i < totalSenders; ++i) {
                    var li = liOld.clone();
                    li.removeClass("thumbnail-item-template").show().find(".data-img").attr("src", "data:image/jpeg;base64,"+result.thumbNails[i]);
                    li.find(".data-name").html(result.senderNames[i]);
                    $(".thumbnail-list").append(li);
                }

                for (var i = 0; i < totalSenders; ++i) {
                    if (result.senders[i] == Office.context.mailbox.userProfile.emailAddress) {
                        ChangeStatusToCantSendKudos();
                    }
                }
            }
        }
    });
}

function SendKudosRequest() {
    ChangeStatusToCantSendKudos();
    $.ajax({
        type: "POST",
        url: serviceBaseUrl + "/api/KudosService",
        contentType: "application/json; charset=utf-8",
        data: JSON.stringify(MakeSendKudosJson()),
        dataType: "json",
        success: function () {
            QueryKudosRequest();
        },
        error: function () {
        }
    });
};

function ChangeStatusToCantSendKudos() {
    $(".send-sec").hide();
}

function MakeSendKudosJson() {
    var item = Office.context.mailbox.item;
    var json = {
        "kudossender": Office.context.mailbox.userProfile.emailAddress,
        "kudossendername": Office.context.mailbox.userProfile.displayName,
        "kudosreceiver": item.sender.emailAddress,
        "kudosreceivername" : item.sender.displayName,
        "itemid": Office.context.mailbox.item.itemId,
        "subject": Office.context.mailbox.item.subject,
        "additionalmessage": document.getElementById("kudosComment").value,
        "senderemailaddress": Office.context.mailbox.userProfile.emailAddress
    };
    return json;
}

// Shows the service response.
function showResponse(response) {
    showToast("Service Response", "Attachments processed: " + response.attachmentsProcessed);
}

// Displays a message for 10 seconds.
function showToast(title, message) {

    var notice = document.getElementById("notice");
    var output = document.getElementById('output');

    notice.innerHTML = title;
    output.innerHTML = message;

    $("#footer").show("slow");

    window.setTimeout(function () { $("#footer").hide("slow") }, 10000);
};