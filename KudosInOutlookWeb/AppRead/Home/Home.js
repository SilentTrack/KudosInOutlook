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
    $("#footer").hide();
    document.getElementById("label1").innerHTML = Office.context.mailbox.item.sender.displayName + "?";
    document.getElementById("label2").innerHTML = Office.context.mailbox.item.sender.displayName;
    QueryKudosRequest();
}

function QueryKudosRequest() {
    $.ajax({
        url: "https://localhost:44372/api/KudosService?InternetMessageID=" + Office.context.mailbox.item.internetMessageId,
        success: function (senders) {
            if (senders.length > 0) {
                for (var i = 0; i < senders.length; ++i)
                {
                    if (senders[i] == Office.context.mailbox.userProfile.displayName) {
                        ChangeStatusToCantSendKudos();
                    }
                }
                var displayString = senders[0];
                for (var i = 1; i < senders.length; ++i) {
                    displayString += "\r\n" + senders[i];
                }
                document.getElementById("kudosQueryResult").value = displayString;
            }
        },
    });
}

function SendKudosRequest() {
    $.ajax({
        type: "POST",
        url: "https://localhost:44372/api/KudosService",
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
    document.getElementById("sendKudos").innerHTML ="<span class=\"ms-Button-label\">You've already sent a kudos!</span>"
    document.getElementById("sendKudos").onclick = "";
}

function MakeSendKudosJson() {
    var item = Office.context.mailbox.item;
    var json = {
        "kudossender": Office.context.mailbox.userProfile.displayName,
        "kudosreceiver": item.sender.displayName,
        "internetmessageId": item.internetMessageId,
        "additionalmessage": document.getElementById("kudosComment").value
    };
    return json;
}

function MakeQueryKudosJson() {
    var item = Office.context.mailbox.item;
    var json = {
        "internetmessageId": item.internetMessageId,
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