/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global $, document, Office */

import { getGraphData } from "./../helpers/ssoauthhelper";
import sampleData from "./../JsonFiles/Sample1.json";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    $(document).ready(async function () {
      let result = {};
      buildUIReadPane(result);
      //getGraphData("ReadPane");
    });
  }
});

export function buildUIReadPane(result: object): void {
  let sample = sampleData.Products.filter((x) => x.controlsVisibility.includes("Read")).sort();
  let divControls = [];
  let product = document.createElement("div");
  let displayContainer = document.createElement("div");
  let listContainer = document.createElement("ul");
  displayContainer.className = "tab-content clear-fix";
  for (let i = 0; i < sample.length; i++) {
    let displayContent = document.createElement("div");
    listContainer.className = "nav nav-tabs";
    let liContainer = document.createElement("li");
    if (i == 0) {
      liContainer.className = "active";
      displayContent.className = "tab-pane active";
    } else {
      displayContent.className = "tab-pane";
    }
    let ahref = document.createElement("a");
    ahref.setAttribute("data-scr", sample[i].readControlsUrl);
    ahref.setAttribute("data-toggle", "tab");
    displayContent.id = sample[i].Name + "_Dispaly";
    ahref.href = "#" + sample[i].Name + "_Dispaly";
    ahref.innerText = sample[i].Name;
    liContainer.appendChild(ahref);
    listContainer.appendChild(liContainer);
    product.appendChild(listContainer);
    displayContainer.appendChild(displayContent);
    divControls.push({ controlId: "#" + sample[i].Name + "_Dispaly", url: sample[i].readControlsUrl });
  }

  document.getElementById("rootElement").appendChild(product);
  document.getElementById("rootElement").appendChild(displayContainer);
  for (let j = 0; j < divControls.length; j++) {
    $(divControls[j]["controlId"]).load(divControls[j]["url"]);
  }
}
