import { alerts } from "../common/alerts";
import { closeCanceledDialog, closeSuccessDialog, getDialogData } from "./dialogs";

export const JwtInputDialogV01Url = "https://localhost:3000/jwtInputDialog_V0.1.html";

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension, Office */

Office.onReady(() => {
  $(function () {
    $("#ok-btn").on("click", onOkClick);
    $("#close-btn").on("click", onCancelClick);

    const data = getDialogData(JwtInputDialogV01Url);
    if (data && data.data) {
      $("#jwt-txt").val(data.data);
    }
  });
});

function onOkClick() {
  const jwt = $("#jwt-txt").val() as string;
  const isValid = isValidJwt(jwt);
  if (isValid === true) {
    closeSuccessDialog(jwt);
  } else {
    alerts.error(<string>isValid);
  }
}

function onCancelClick() {
  closeCanceledDialog();
}

function isValidJwt(jwt: string): boolean | string {
  if (jwt) {
    return true;
  } else {
    return "JWT is required";
  }
}
