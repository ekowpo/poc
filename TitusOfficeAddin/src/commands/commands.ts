/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady((info) => {
  // If needed, Office.js is ready to be called
  if (info.host === Office.HostType.Outlook) {
    $(document).ready(function () {
      officeAddinonSend();
    });
  }
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
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

// eslint-disable-next-line no-unused-vars
function officeAddinonSend() {
  // eslint-disable-next-line no-undef
  console.log("This is office Add-in onsend");
}
const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;
g.officeAddinonSend = officeAddinonSend;
