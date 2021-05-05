// eslint-disable-next-line no-unused-vars
import { IdentClient, authenticate, authenticateStub } from "./ident-client";
import { alerts } from "./alerts";

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import { LoginFormData } from "./login-form-data";
import { onError } from "./common";
import { excelWorker } from "./excel-worker";

const stubAuth = false;
const testAfterLogin = false;

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension, Office */

let identClient: IdentClient | null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    $(function () {
      initUi();
      setUiForLogin();  
    });
  }
});

function initUi() {
  $("#login-btn").on("click", login);

  $("#logout-btn").on("click", logout);

  $('#get-workgroups-btn').on('click', fillWorkgroups);
}

function setUiForLogin() {
  $("#sideload-msg").hide();
  $("#login-ui").show();
  $("#work-ui").hide();
  $("#app-body").show();
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

  const authenticateFn = stubAuth 
    ? authenticateStub
    : authenticate;

    authenticateFn(loginFormData).then((client) => {
    identClient = client;
    
    loginFormData.clean();
    setUiAfterLogin();

    if (testAfterLogin) {
      test();
    }
    
  }, onError);
}

function logout() {
  if (!identClient) {
    setUiForLogin();
    return;
  }
    
  identClient.logout().then(() => {
    setUiForLogin();
    identClient = null;
  }, onError);
}

function fillWorkgroups(): Promise<unknown> {
  if (!identClient) {
    setUiForLogin();
    return;
  }

  return identClient.getWorkgroups().then((apps) => {
    return excelWorker.showWorkgroups(apps);
  }, onError);
}

function test() {
  Excel.run((context) => {
    const cursheet = context.workbook.worksheets.getActiveWorksheet();
    const cellA1_A2 = cursheet.getRange("A1:A3");

    // const value = new Date(); // identClient.test_ExpiresAt();
    const value = identClient?.test_expiresAt;
    cellA1_A2.values = [[ value ], [ new Date() ], [ identClient?.isExpired ]];
    cellA1_A2.format.autofitColumns();

    return context.sync();
  })
  .catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
      onError(error.message);
    } else {
      onError(error);
    }
  })
}