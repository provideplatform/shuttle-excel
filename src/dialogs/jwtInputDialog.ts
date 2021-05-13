import { alerts } from "../common/alerts";
import { closeCanceledDialog, closeDialog, initializedDialog } from "../common/common";
import { JwtInputData } from "./models/jwt-input-data";

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension, Office */

Office.onReady(() => {
  $(function () {
    // NOTE: For demo - send data to dialog - part 2
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onMessageFromParent);

    $("#ok-btn").on("click", onOkClick);
    $("#close-btn").on("click", closeCanceledDialog);
    initializedDialog();
  });
});

function onOkClick() {
  const $form = $("form");
  const formData = new JwtInputData($form);
  const isValid = formData.isValid();
  if (isValid === true) {
    closeDialog(formData);
  } else {
    alerts.error(<string>isValid);
  }
}

// NOTE: For demo - send data to dialog - part 3
function onMessageFromParent(event) {
  var messageFromParent = JSON.parse(event.message);
  $("#jwt-txt").val(messageFromParent.data);
}
