/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest


/* global document, Office */

Office.onReady((info) => {
  // eslint-disable-next-line no-empty
  if (info.host === Office.HostType.Outlook) {
  }
});

export async function run() {
  let selectedValue = (document.querySelector('input[name="policy"]:checked') as HTMLSelectElement).value;
  Office.context.mailbox.item.internetHeaders.setAsync({ "x-hs-Veraheaders": selectedValue }, function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      // eslint-disable-next-line no-undef
      console.log("Successfully set headers");
    } else {
      // eslint-disable-next-line no-undef
      console.log("Error setting headers: " + JSON.stringify(asyncResult.error));
    }
  });
}
