/* eslint-disable no-undef */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    $(document).ready(function () {
      getSelectedCustomHeaders();
    });
  }
});

// Get custom internet headers.
function getCallback(asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
    document.getElementById("classification").innerText =
      "Classification: " + asyncResult.value.match(/x-hs-customheaders:.*/gim)[0].slice(19);
  } else {
    console.log("Error getting preferences from header: " + JSON.stringify(asyncResult.error));
  }
}
// Get custom internet headers.
function getSelectedCustomHeaders() {
  Office.context.mailbox.item.getAllInternetHeadersAsync(getCallback);
}
