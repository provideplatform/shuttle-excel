// eslint-disable-next-line no-unused-vars
import { IdentClient, authenticate, authenticateStub, restore, restoreStub } from "./ident-client";
import { alerts } from "./alerts";
import { LoginFormData } from "./models/login-form-data";
import { DialogEvent, Jwtoken, onError } from "./common";
import { excelWorker } from "./excel-worker";
import { settings } from "../settings/settings";

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import { TokenStr } from "./models/common";
import { User } from "./models/user";

const stubAuth = true;

const JwtokenDialogUrl = "https://localhost:3000/jwtInputDialog.html";

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension, Office */

let identClient: IdentClient | null;

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

  return Promise.all([settings.getRefreshToken(), settings.getUser()]).then((data) => {
    const refreshToken = data[0] as TokenStr;
    const user = data[1] as User;
    if (!refreshToken || !user) {
      setUiForLogin();
      return;
    }

    const restoreFn = stubAuth ? restoreStub : restore;

    restoreFn(refreshToken, user).then(
      (client) => {
        identClient = client;
        setUiAfterLogin();
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

    const token: TokenStr = identClient.refreshToken;
    const user: User = { id: identClient.user.id, name: identClient.user.name, email: identClient.user.email };

    settings.setTokenAndUser(token, user);
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

function onGetJwtokenDialog() {
  getJwtokenDialog()
    .then(
      (jwtoken) => {
        // alerts.success(["JWT", jwtoken]);
        return identClient.acceptWorkgroupInvitation(jwtoken);
      },
      () => {
        return false;
      }
    )
    .then((result) => {
      if (result !== false) {
        alerts.success("Invitation completed");
      }
    });
}

function getJwtokenDialog(): Promise<Jwtoken> {
  // debugger;
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
