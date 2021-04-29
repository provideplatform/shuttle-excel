// eslint-disable-next-line no-unused-vars
import { IdentClient, authenticate, authenticateStub } from "./ident-client";
import { alerts } from "./alerts";

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import { LoginFormData } from "./login-form-data";

// eslint-disable-next-line no-unused-vars
/* global Excel, Office */

let identClient: IdentClient | null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    $(function () {
      initUi();
    });
  }
});

function initUi() {
  $("#sideload-msg").hide();
  $("#login-ui").show();
  $("#work-ui").hide();
  $("#app-body").show();
  $("#login-btn").on("click", login);
}

function setUiAfterLogin() {
  let $workUi = $("#work-ui");
  identClient.getUserFullName().then((userFullName) => {
    $("#user-name", $workUi).text(userFullName);
    $("#login-ui").hide();
    $workUi.show();
  });
}

function login() {
  var $form = $("#login-ui form");
  const loginFormData = new LoginFormData($form);
  const isValid = loginFormData.isValid();
  if (isValid !== true) {
    alerts.error(<string>isValid);
    return;
  }

  authenticateStub(loginFormData).then((client) => {
    identClient = client;
    setUiAfterLogin();
  }, onError);
}

function onError(reason) {
  let message = reason.toString();
  console.log(message);
  if (message.indexOf("Error: ") == 0) {
    message = message.substring("Error: ".length);
  }

  alerts.error(message);
}
