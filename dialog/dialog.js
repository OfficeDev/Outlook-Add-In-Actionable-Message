// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. see LICENSE in the project root for license information.
"use strict";

var params;

Office.initialize = function(reason) {
  $(document).ready(function(){
    // Handle the onclick for the convert message button
    $("#send-msg").click(sendMessage)

    // Get the query params
    var uri = new URI(window.location);
    params = uri.search(true);

    var initContext = generateSampleContext();

    $("#init-context").val(JSON.stringify(initContext, null, 2));
  });
}

function generateSampleContext() {
  var context = {
    dateStamp: moment().format(),
    myStringProperty: "Hello world",
    myBooleanProperty: true
  };

  return context;
}

function sendMessage() {
  // Load the context and make sure it parses
  var context;
  try {
    context = JSON.parse($("#init-context").val());
  } catch (err) {
    showError("Could not parse the initialization context. Make sure it is valid JSON and try again.");
    return;
  }

  // Generate the proper HTML body including the card
  var htmlBody = generateActionableMessageBody(context);

  // Build a REST API message
  // Message resource doced at 
  // https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/resources/message
  // However, since the add-in API requires calling the Outlook REST endpoint directly,
  // need to convert property names to PascalCase
  // https://docs.microsoft.com/en-us/outlook/rest/compare-graph-outlook#translating-between-graph-and-outlook
  var message = {
    Subject: "ACTIVATE YOUR ADD-IN",
    Body: {
      ContentType: "HTML",
      Content: htmlBody
    },
    ToRecipients: [
      {
        EmailAddress: {
          Address: params.userEmail
        }
      }
    ]
  };

  var sendMailPayload = {
    Message: message,
    SaveToSentItems: true
  };

  // Use AJAX to POST a send mail request
  $.ajax({
    type: "POST",
    url: params.restUrl + "/v2.0/me/sendmail",
    dataType: "json",
    data: JSON.stringify(sendMailPayload),
    headers: {
      "Authorization": "Bearer " + params.token,
      "Content-Type": "application/json"
    }
  }).always(function(jqxhr){
    if (jqxhr.status == 202) {
      // Send message back signaling success
      Office.context.ui.messageParent(true);
    } else {
      // Report error
      showError("Could not send message because of the following error: " + JSON.stringify(jqxhr));
    }
  });
}

function generateActionableMessageBody(context) {
  // First build the card
  var actionCard = {
    "@type": "MessageCard",
    "@context": "http://schema.org/extensions",
    originator: "527104a1-f1a5-475a-9199-7a968161c870",
    hideOriginalBody: true,
    summary: "This is the summary property",
    sections: [
        {
            title: "Activate **Actionable Message Activation** add-in",
            startGroup: true,
            facts: convertContextToNameValuePairs(context),
            potentialAction: [
                {
                    "@type": "InvokeAddInCommand",
                    name: "Invoke \"View Initialization Context\"",
                    addInId: "527104a1-f1a5-475a-9199-7a968161c870",
                    desktopCommandId: "showInitContext",
                    initializationContext: context
                }
            ]
        },
        {
            title: "Activate **Message Header Analyzer**",
            text: "Click the button below to see \"on demand\" installation of an add-in from the Office Store.",
            startGroup: true,
            potentialAction: [
                {
                    "@type": "InvokeAddInCommand",
                    name: "Invoke MHA's \"View Headers\" command",
                    addInId: "62916641-fc48-44ae-a2a3-163811f1c945",
                    desktopCommandId: "mhaOpenPaneButton"
                }
            ]
        }
    ]
  };

  // Now build the HTML body
  var body = "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"><script type=\"application/ld+json\">" +
    JSON.stringify(actionCard) + "</script></head><body><p>If you don't see a message card above with clickable buttons, your email client doesn't support Actionable Messages. Please try viewing this mail in Outlook on the web for Office 365, or the latest version of Outlook 2016 for Windows.</p></body></html>";

  return body;
}

function convertContextToNameValuePairs(context) {
  var nvp = [];
  for (var propName in context) {
    nvp.push({ name: propName, value: context[propName].toString() });
  }

  return nvp;
}

function showError(message) {
  $("#error-msg").text(message);
  $("#error").show();
}