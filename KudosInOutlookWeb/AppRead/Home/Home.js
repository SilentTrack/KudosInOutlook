/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

/// <reference path="../App.js" />
var xhr;
var serviceRequest;
//var serviceBaseUrl = "https://kudosservice.azurewebsites.net";
var serviceBaseUrl = "https://localhost:44372";
var totalSenders;
var sending;

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
    //$(".thumbnail-item-template").slideUp();
    $(".thumbnail-item-template").hide();
    sending = false;
    QueryKudosRequest();
}

function QueryKudosRequest() {
    var itemID = Office.context.mailbox.item.itemId;
    $(".data-receiver").html(Office.context.mailbox.item.sender.displayName);
    $.ajax({
        url: serviceBaseUrl + "/api/KudosService?ItemID=" + encodeURIComponent(itemID),
        success: function (result) {
            totalSenders = result.senders.length;
            if (totalSenders == 0) {
                $(".prompt-sec").fadeIn();
                $(".thumbnail-sec").fadeOut();
                $(".send-sec").fadeIn();
            }
            else {
                $(".prompt-sec").fadeOut();
                $(".thumbnail-sec").fadeIn();
                $(".data-count").html(totalSenders);
                $(".thumbnail-list").html();

                var flag = false;
                for (var i = 0; i < totalSenders; ++i) {
                    if (result.senders[i] == Office.context.mailbox.userProfile.emailAddress) {
                        flag = true;
                        break;
                    }
                }
                if (!flag) {
                    $(".send-sec").fadeIn();
                }

                var liOld = $(".thumbnail-item-template");
                for (var i = 0; i < totalSenders; ++i) {
                    var li = liOld.clone();
                    $(".thumbnail-list").append(li);
                    li.removeClass("thumbnail-item-template").find(".data-img").attr("src", "data:image/jpeg;base64," + result.thumbNails[i]);
                    li.find(".data-name").html(result.senderNames[i]);
                    li.fadeIn();
                }
            }
        }
    });
}

function SendKudosRequest() {
    if (sending) {
        return;
    }
    sending = true;
    $(".ms-Button-label").html("Sending Kudos...");
    $.ajax({
        type: "POST",
        url: serviceBaseUrl + "/api/KudosService",
        contentType: "application/json; charset=utf-8",
        data: JSON.stringify(MakeSendKudosJson()),
        dataType: "json",
        success: function (result) {
            $(".prompt-sec").hide();
            $(".send-sec").hide();
            $(".thumbnail-sec").fadeIn();
            ++totalSenders;
            $(".data-count").html(totalSenders);
            $(".thumbnail-list").html();

            var liOld = $(".thumbnail-item-template");
            var li = liOld.clone();
            $(".thumbnail-list").append(li);
            li.removeClass("thumbnail-item-template").find(".data-img").attr("src", "data:image/jpeg;base64," + result);
            li.find(".data-name").html(Office.context.mailbox.userProfile.displayName);
            li.fadeIn();
            $(".ms-Button-label").html("Send Kudos!");
            sending = false;
        },
        error: function () {
        }
    });
};

function MakeSendKudosJson() {
    var item = Office.context.mailbox.item;
    var comment = document.getElementById("kudosComment").value;
    if (comment == "")
    {
        comment = "Thanks for your contribution!";
    }
    var json = {
        "kudossender": Office.context.mailbox.userProfile.emailAddress,
        "kudossendername": Office.context.mailbox.userProfile.displayName,
        "kudosreceiver": item.sender.emailAddress,
        "kudosreceivername": item.sender.displayName,
        "itemid": Office.context.mailbox.item.itemId,
        "subject": Office.context.mailbox.item.subject,
        "additionalmessage": comment,
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