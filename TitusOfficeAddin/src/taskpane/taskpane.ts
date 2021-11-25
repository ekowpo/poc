/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // eslint-disable-next-line no-undef
    $(document).ready(function () {
      document.getElementById("Apply").onclick = run;
    });
  }
});

export async function run() {
  let classification: HTMLSelectElement = document.getElementById("Classification") as HTMLSelectElement;
  const selectedValue = classification.value;
  Office.context.mailbox.item.internetHeaders.setAsync({ "x-hs-customheaders": selectedValue }, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      // eslint-disable-next-line no-undef
      console.log("Successfully set headers");
    } else {
      // eslint-disable-next-line no-undef
      console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
    }
  });
}
