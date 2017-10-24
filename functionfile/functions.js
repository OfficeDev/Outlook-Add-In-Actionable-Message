// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. see LICENSE in the project root for license information.

var sendDlg;
var btnEvent;

Office.initialize = function () {
};

function showError(error) {
  Office.context.mailbox.item.notificationMessages.replaceAsync("send-msg-error", {
    type: "errorMessage",
    message: error
  }, function(result){
  });
}

function sendActionableMessage(event) {
  // Save the event object so we can finish up later
  btnEvent = event;

  // Get the REST access token
  Office.context.mailbox.getCallbackTokenAsync(
    { isRest: true },
    function(result) {
      if (result.status == Office.AsyncResultStatus.Succeeded) {
        var token = result.value;
        
        var url = new URI("../dialog/dialog.html", window.location);

        url.addSearch("token", token);
        url.addSearch("restUrl", Office.context.mailbox.restUrl);
        url.addSearch("userEmail", Office.context.mailbox.userProfile.emailAddress);

        var dialogOptions = { width: 70, height: 80, displayInIframe: true };
      
        Office.context.ui.displayDialogAsync(url.toString(), dialogOptions, function(result) {
          sendDlg = result.value;
          sendDlg.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, receiveMessage);
          sendDlg.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogEventReceived, dialogClosed);
        });
      } else {
        showError("Could not retrieve access token: " + JSON.stringify(result.error));
        btnEvent.completed();
      }
    }
  );
}

function receiveMessage(success) {
  if (success) {
    sendDlg.close();

    Office.context.mailbox.item.notificationMessages.replaceAsync("send-msg-success", {
      type: "informationalMessage",
      icon: "icon-16",
      message: "Message sent successfully.",
      persistent: false
    }, function(result){
      btnEvent.completed();
      btnEvent = null;
    });
  }
}

function dialogClosed(message) {
  sendDlg = null;
  btnEvent.completed();
  btnEvent = null;
}