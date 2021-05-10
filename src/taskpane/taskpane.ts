// eslint-disable-next-line no-unused-vars
import { IdentClient, authenticate, authenticateStub } from "./ident-client";
import { alerts } from "./alerts";
import { LoginFormData } from "./login-form-data";
import { DialogEvent, Jwtoken, onError } from "./common";
import { excelWorker } from "./excel-worker";
import { documentSettings, localStorageSettings, sessionStorageSettings } from "./settings/settings";

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

const stubAuth = true;

const JwtokenDialogUrl = "https://localhost:3000/jwtInputDialog.html";

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension, Office */

let identClient: IdentClient | null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    $(function () {
      initUi();
      setUiForLogin();

      showTestSettings();
    });
  }
});

function initUi() {
  $("#login-btn").on("click", onLogin);

  $("#logout-btn").on("click", onLogout);

  $("#get-workgroups-btn").on("click", onFillWorkgroups);

  $("#set-t-settings-btn").on("click", onSetTestSettings);
  $("#show-t-settings-btn").on("click", onShowTestSettings);

  $("#show-jwt-input-btn").on("click", onGetJwtokenDialog);
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

function onLogin() {
  var $form = $("#login-ui form");
  const loginFormData = new LoginFormData($form);
  const isValid = loginFormData.isValid();
  if (isValid !== true) {
    alerts.error(<string>isValid);
    return;
  }

  const authenticateFn = stubAuth ? authenticateStub : authenticate;

  authenticateFn(loginFormData).then((client) => {
    identClient = client;

    loginFormData.clean();
    setUiAfterLogin();
  }, onError);
}

function onLogout() {
  if (!identClient) {
    setUiForLogin();
    return;
  }

  identClient.logout().then(() => {
    setUiForLogin();
    identClient = null;
  }, onError);
}

function onFillWorkgroups(): Promise<unknown> {
  if (!identClient) {
    setUiForLogin();
    return;
  }

  return identClient.getWorkgroups().then((apps) => {
    return excelWorker.showWorkgroups(apps);
  }, onError);
}

function onSetTestSettings() {
  const object = { val: "Value!" };
  documentSettings.set("TestSettings", object);
  localStorageSettings.set("TestSettings", object);
  sessionStorageSettings.set("TestSettings", object);
}

function onShowTestSettings() {
  showTestSettings();
}

function showTestSettings() {
  var docSetsPromise = documentSettings.get("TestSettings");
  var locStgSetsPromise = localStorageSettings.get("TestSettings");
  var sessStgSetsPromise = sessionStorageSettings.get("TestSettings");

  Promise.all([docSetsPromise, locStgSetsPromise, sessStgSetsPromise]).then((values) => {
    const messages = values.map((x) => JSON.stringify(x));
    messages.unshift("Settings");
    alerts.warn(messages);
  });
}

function onGetJwtokenDialog() {
  getJwtokenDialog()
    .then((jwtoken) => {
      // alerts.success(["JWT", jwtoken]);
      return identClient.acceptWorkgroupInvitation(jwtoken);
    }, () => { return false; })
    .then((result) => {
      if (result !== false) {
        alerts.success("Invitation completed");
      }
    });
}

function getJwtokenDialog(): Promise<Jwtoken> {
  debugger;
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      JwtokenDialogUrl,
      { height: 37, width: 35, displayInIframe: true },
      (result: Office.AsyncResult<Office.Dialog>) => {
        const dialog = result.value;
        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (args: { message?: string | boolean; error?: number }) => {
            if (args.error) {          
              alerts.error(args.error + "");
              return;
            }

            const dialogResult = JSON.parse(args.message + "");
            switch (dialogResult.result) {
              case DialogEvent.Initialized: {
                // NOTE: For demo - send data to dialog - part 1
                console.log("A3");
                var messageToDialog = JSON.stringify({ data: "Test JWT" });
                dialog.messageChild(messageToDialog);
                break;
              }
              case DialogEvent.Ok: {
                dialog.close();
                const jwtoken = dialogResult.data as Jwtoken;
                resolve(jwtoken);
                break;
              }

              case DialogEvent.Cancel: {
                dialog.close();
                reject();
                break;
              }
            }
          }
        );
      }
    );
  });
}
