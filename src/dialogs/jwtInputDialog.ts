import { alerts } from "../common/alerts";
import { closeCanceledDialog, closeSuccessDialog, getDialogData } from "./dialogs";
import { JwtInputData } from "./models/jwt-input-data";

export const JwtInputDialogUrl = "https://localhost:3000/jwtInputDialog.html";

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension, Office */

Office.onReady(() => {
  $(function () {
    $("#ok-btn").on("click", onOkClick);
    $("#close-btn").on("click", onCancelClick);

    const data = getDialogData(JwtInputDialogUrl);
    if (data && data.data) {
      $("#jwt-txt").val(data.data);
    }
  });
});

function onOkClick() {
  const $form = $("form");
  const formData = new JwtInputData($form);
  const isValid = formData.isValid();
  if (isValid === true) {
    closeSuccessDialog(formData);
    formData.clean();
  } else {
    alerts.error(<string>isValid);
  }
}

function onCancelClick() {
  closeCanceledDialog();
}
