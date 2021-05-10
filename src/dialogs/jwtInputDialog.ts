import { alerts } from "../taskpane/alerts";
import { closeCanceledDialog, closeDialog, initializedDialog } from "../taskpane/common";

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension, Office */

Office.onReady(() => {
    console.log("P1");
  

  $(function () {
    console.log("P3");
    // NOTE: For demo - send data to dialog - part 2
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onMessageFromParent);

    $("#ok-btn").on("click", onOkClick);
    $("#close-btn").on("click", closeCanceledDialog);
    initializedDialog();
  });
});

function onOkClick() {
  const jwt = $("#jwt-txt").val() as string;
  const isValid = isValidJwt(jwt);
  if (isValid === true) {
    closeDialog(jwt);
  } else {
    alerts.error(<string>isValid);
  }
}

function isValidJwt(jwt: string): boolean | string {
  if (jwt) {
    return true;
  } else {
    return "JWT is required";
  }
}

// NOTE: For demo - send data to dialog - part 3
function onMessageFromParent(event) {
    console.log("P2");
  var messageFromParent = JSON.parse(event.message);
  $("#jwt-txt").val(messageFromParent.data);
}
