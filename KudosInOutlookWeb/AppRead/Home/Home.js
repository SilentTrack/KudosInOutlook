/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/

/// <reference path="../App.js" />
var xhr;
var serviceRequest;
var serviceBaseUrl = "https://kudosservice.azurewebsites.net";
//var serviceBaseUrl = "https://localhost:44372";

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
    var itemID = Office.context.mailbox.item.itemId;
    $.ajax({
        url: serviceBaseUrl + "/api/KudosService?ItemID=" + encodeURIComponent(itemID),
        success: function (result) {
            //$("#thumbNail").html(result.senders.length);
            //var totalSenders = result.senders.length;
            //if (totalSenders > 0) {
            //    for (var i = 0; i < totalSenders; ++i) {
            //        if (result.senders[i] == Office.context.mailbox.userProfile.displayName) {
            //            ChangeStatusToCantSendKudos();
            //        }
            //    }

            //    var newRow;
            //    var table = document.getElementById("kudosQueryResult");
            //    for (var i = 1; i < table.rows.length; ++i) {
            //        table.deleteRow(i);
            //    }

            //    while (document.getElementById("thumbNail").hasChildNodes()) {
            //        document.getElementById("thumbNail").removeChild(document.getElementById("thumbNail").firstChild);
            //    }
            //    for (var i = 0; i < totalSenders; ++i) {
            //        newRow = table.insertRow(table.rows.length);
            //        newRow.insertCell(0).innerHTML = result.senders[i];
            //        newRow.insertCell(1).innerHTML = result.sentTime[i];
            //        var img = document.createElement("img");
            //        img.src = "data:image/jpeg;base64," + result.thumbNails[i];
            //        img.width = 48;
            //        img.height = 48;
            //        document.getElementById("thumbNail").appendChild(img);
            //    }
            //}
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
    document.getElementById("sendKudos").innerHTML = "<span class=\"ms-Button-label\">You've already sent a kudos!</span>"
    document.getElementById("sendKudos").onclick = "";
}

function MakeSendKudosJson() {
    var item = Office.context.mailbox.item;
    var json = {
        "kudossender": Office.context.mailbox.userProfile.displayName,
        "kudosreceiver": item.sender.displayName,
        "itemid": Office.context.mailbox.item.itemId,
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