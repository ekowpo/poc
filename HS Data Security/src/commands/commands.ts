/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

import sampleData from "./../JsonFiles/Sample1.json";

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
export function action(event: Office.AddinCommands.Event) {
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

const initOnSend = async (event) => {
  //window.alert('test');
  let sample = sampleData.Products.filter((x) => x.eventsList.filter((y) => y.event === "onSend"));
  let onsendEvent = [];
  for (let i = 0; i < sample.length; i++) {
    const OnSend = sample[i].eventsList.filter((x) => x.event === "onSend")[0];
    if (OnSend) {
      onsendEvent.push(OnSend);
    }
  }
  onsendEvent.sort((a, b) => {
    return a["eventOrder"] - b["eventOrder"];
  });
  for (let i = 0; i < onsendEvent.length; i++) {
    const url = onsendEvent[i]["eventUrl"];
    // eslint-disable-next-line no-unused-vars
    await $.getScript(url, function (data) {
    });
  }
  event.completed({ allowEvent: true });
};

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action;
g.initOnSend = initOnSend;
