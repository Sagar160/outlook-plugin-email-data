/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady().then(body());

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };
  
 

  const msgCc = Office.context.mailbox.item.cc;
  console.log("Message copied to:");
  for (let i = 0; i < msgCc.length; i++) {
    console.log(msgCc[i].displayName + " (" + msgCc[i].emailAddress + ")");
  }
  console.log(`Subject: ${Office.context.mailbox.item.subject}`);

  console.log(1)
  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function body() {
  console.log('body')
  let emailBody = '';
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
    var ewsId = Office.context.mailbox.item.itemId;
    var token = result.value;

    // var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0); this does not work on API version 1.1
    var restId = ewsId.replaceAll("/", "-").replaceAll("+", "_"); // Convert ewsId to restId
    var getMessageUrl = (Office.context.mailbox.restUrl || 'https://outlook.office365.com/api') + '/v2.0/me/messages/' + restId;
    var xhr = new XMLHttpRequest();
    xhr.open('GET', getMessageUrl, false);
    xhr.setRequestHeader('Prefer', 'outlook.body-content-type="html"') // for retrieving body as HTML
    xhr.setRequestHeader("Authorization", "Bearer " + token);
    xhr.onload = (e) => {
      var json = JSON.parse(xhr.responseText);
      emailBody = json.Body.Content;
      console.log(emailBody)
    }
    xhr.send();
  });
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
