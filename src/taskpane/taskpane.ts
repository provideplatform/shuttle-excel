// eslint-disable-next-line no-unused-vars
import { ProvideClient, authenticate, authenticateStub, restore, restoreStub } from "../client/provide-client";
import { alerts, spinnerOff, spinnerOn } from "../common/alerts";
import { LoginFormData } from "../models/login-form-data";
import { DialogEvent, onError } from "../common/common";
import { excelWorker } from "./excel-worker";
import { settings } from "../settings/settings";
import { TokenStr } from "../models/common";
import { User } from "../models/user";
import { JwtInput } from "../dialogs/models/jwt-input-data";

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

const stubAuth = true;

const JwtokenDialogUrl = "https://localhost:3000/jwtInputDialog.html";

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension, Office */

let identClient: ProvideClient | null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    $(function () {
      initUi();

      tryRestoreAutorization();
    });
  }
});

function tryRestoreAutorization() {
  // debugger;
  settings.getRefreshToken();

  return Promise.all([settings.getRefreshToken(), settings.getUser()]).then((data) => {
    const refreshToken = data[0] as TokenStr;
    const user = data[1] as User;
    if (!refreshToken || !user) {
      setUiForLogin();
      spinnerOff();
      return;
    }

    const restoreFn = stubAuth ? restoreStub : restore;
    spinnerOn();
    restoreFn(refreshToken, user).then(
      (client) => {
        identClient = client;
        setUiAfterLogin();
        spinnerOff();
      },
      (reason) => {
        settings.removeTokenAndUser();
        setUiForLogin();
        onError(reason);
      }
    );
  });
}

function initUi() {
  $("#login-btn").on("click", onLogin);

  $("#logout-btn").on("click", onLogout);
  $("#get-workgroups-btn").on("click", onFillWorkgroups);
  $("#show-jwt-input-btn").on("click", onGetJwtokenDialog);
}

function setUiForLogin() {
  $("#sideload-msg").hide();
  $("#login-ui").show();
  $("#work-ui").hide();
  $("#app-body").show();
}

function setUiAfterLogin() {
  $("#sideload-msg").hide();
  let $workUi = $("#work-ui");
  const userName = (identClient.user || {}).name || "unknow";
  $("#user-name", $workUi).text(userName);
  $("#login-ui").hide();
  $workUi.show();
  $("#app-body").show();
}

function onLogin() {
  const $form = $("#login-ui form");
  const loginFormData = new LoginFormData($form);
  const isValid = loginFormData.isValid();
  if (isValid !== true) {
    alerts.error(<string>isValid);
    return;
  }

  const authenticateFn = stubAuth ? authenticateStub : authenticate;
  spinnerOn();
  authenticateFn(loginFormData).then((client) => {
    identClient = client;

    loginFormData.clean();
    setUiAfterLogin();

    const token: TokenStr = identClient.userRefreshToken;
    const user: User = { id: identClient.user.id, name: identClient.user.name, email: identClient.user.email };

    settings.setTokenAndUser(token, user);
    spinnerOff();
  }, onError);
}

function onLogout() {
  if (!identClient) {
    setUiForLogin();
    return;
  }

  identClient
    .logout()
    .then(() => {
      identClient = null;
      return settings.removeTokenAndUser();
    }, onError)
    .then(() => {
      setUiForLogin();
      spinnerOff();
    }, onError);
}

function onFillWorkgroups(): Promise<unknown> {
  if (!identClient) {
    setUiForLogin();
    return;
  }

  spinnerOn();
  return identClient.getWorkgroups().then((apps) => {
    spinnerOff();
    return excelWorker.showWorkgroups(apps);
  }, onError);
}

function onGetJwtokenDialog() {
  getJwtokenDialog().then((jwtInput) => {
    spinnerOn();
    // TODO: ??????
    // const organizationId: Uuid = "sdfgsdfgsdfg";
    return identClient.acceptWorkgroupInvitation(jwtInput.jwt, jwtInput.orgId).then(() => {
      spinnerOff();
      alerts.success("Invitation completed");
    }, onError);
  }, () => { });
}

function getJwtokenDialog(): Promise<JwtInput> {
  // debugger;
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      JwtokenDialogUrl,
      { height: 38, width: 35, displayInIframe: true },
      (result: Office.AsyncResult<Office.Dialog>) => {
        const dialog = result.value;
        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived, (args: {message: string | boolean}) => {
            const dialogResult = JSON.parse(args.message + "");
            switch (dialogResult.result) {
              case DialogEvent.Initialized: {
                // NOTE: For demo - send data to dialog - part 1
                var messageToDialog = JSON.stringify({ data: "Test JWT" });
                dialog.messageChild(messageToDialog);
                break;
              }
              case DialogEvent.Ok: {
                dialog.close();
                const jwtInput = dialogResult.data as JwtInput;
                resolve(jwtInput);
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
        dialog.addEventHandler(Office.EventType.DialogEventReceived,  (args: { error: number }) => {
          if (args.error === 12006 /*(dialog closed by user)*/) {
            return;
          }

          if (args.error) {
            alerts.error("Dialog error - " + (args.error + ""));
            reject();
          }
        });
      }
    );
  });
}
