// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. see LICENSE in the project root for license information.
"use strict";

Office.initialize = function(reason) {
  $(document).ready(function(){
    // Set up event handler for ItemChanged
    // This will handle the scenario where the user
    // pins the taskpane open and selects a new message
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, 
      loadNewMessage);

    // Set up event handler for InitializationContextChanged
    // This will handle the scenario where the taskpane
    // is already opened before the user clicks the actionable
    // message button
    //Office.context.mailbox.item.addHandlerAsync

    loadNewMessage();
  })
}

function loadNewMessage(eventArgs) {
  var subject = Office.context.mailbox.item.subject;
  $("#subject").text(subject);

  resetPage();

  // Get the intialization context (if present)
  Office.context.mailbox.item.getInitializationContextAsync(
    function(asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
        if (asyncResult.value != null && asyncResult.value.length > 0) {
          // The value is a string, parse to an object
          var context = JSON.parse(asyncResult.value);
          $("#init-context").text(JSON.stringify(context, null, 2));
          $("#has-context").show();
        } else {
          $("#no-context").show();
        }
      } else {
        // Today OWA returns an error instead of empty string
        // when there is no context
        if (asyncResult.error.code == 9020) {
          $("#no-context").show();
        } else {
          // Show the error
          showError(JSON.stringify(asyncResult.error, null, 2));
        }
      }

      // Register for InitializationContextChanged
      Office.context.mailbox.item.addHandlerAsync(
        Office.EventType.InitializationContextChanged,
        loadNewInitContext
      );
    }
  );
}

function resetPage() {
  $("#error").hide();
  $("#error-msg").text("");
  $("#no-context").hide();
  $("#has-context").hide();
  $("#init-context").text("");
}

function loadNewInitContext(eventArgs) {
  $("#handler").text("[Event handler] Args: " + JSON.stringify(eventArgs, null, 2));
}

function showError(message) {
  $("#error-msg").text(message);
  $("#error").show();
}