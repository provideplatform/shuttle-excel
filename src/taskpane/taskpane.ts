import { ProvideClient, authenticate, authenticateStub, restore, restoreStub } from "../client/provide-client";
import { alerts, spinnerOff, spinnerOn } from "../common/alerts";
import { LoginFormData } from "../models/login-form-data";
import { onError } from "../common/common";
import { excelWorker } from "./excel-worker";
import { sessionSettings as session } from "../settings/settings";
import { TokenStr } from "../models/common";
import { User } from "../models/user";
import { showJwtInputDialog } from "../dialogs/dialogs-helpers";
import * as $ from "jquery";

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import "../../assets/logo-filled.png";

import "bootstrap/dist/css/bootstrap.min.css";
import "@fortawesome/fontawesome-free/js/fontawesome";
import "@fortawesome/fontawesome-free/js/solid";
import "@fortawesome/fontawesome-free/js/regular";
import "@fortawesome/fontawesome-free/js/brands";
import "./taskpane.css";

const stubAuth = false;

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension, Office */

let identClient: ProvideClient | null;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    $(function () {
      initUi();

      tryRestoreAutorization().then(fillMyWorkgroupSheet).then(startBaselining);
    });
  }
});

function tryRestoreAutorization() {
  return Promise.all([session.getRefreshToken(), session.getUser()]).then(([refreshToken, user]) => {
    if (!refreshToken || !user) {
      setUiForLogin();
      spinnerOff();
      return;
    }

    const restoreFn = stubAuth ? restoreStub : restore;
    spinnerOn();
    return restoreFn(refreshToken, user).then(
      (client) => {
        identClient = client;
        setUiAfterLogin();
        spinnerOff();
      },
      (reason) => {
        session.removeTokenAndUser();
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
  $("#start-baselining-btn").on("click", onSetupBaselining);
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

function onLogin(): Promise<void> {
  const $form = $("#login-ui form");
  const loginFormData = new LoginFormData($form);
  const isValid = loginFormData.isValid();
  if (isValid !== true) {
    alerts.error(<string>isValid);
    return;
  }

  const authenticateFn = stubAuth ? authenticateStub : authenticate;
  spinnerOn();
  return authenticateFn(loginFormData)
    .then((client) => {
      identClient = client;

      loginFormData.clean();
      setUiAfterLogin();

      const token: TokenStr = identClient.userRefreshToken;
      const user: User = { id: identClient.user.id, name: identClient.user.name, email: identClient.user.email };

      return session.setTokenAndUser(token, user).then(spinnerOff);
    }, onError)
    .then(fillMyWorkgroupSheet)
    .then(startBaselining);
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
      return session.removeTokenAndUser();
    }, onError)
    .then(() => {
      setUiForLogin();
      spinnerOff();
    }, onError);
}

function onFillWorkgroups(): Promise<unknown> {
  return fillMyWorkgroupSheet();
}

function onGetJwtokenDialog() {
  showJwtInputDialog().then(
    // NOTE: For demo - send data to dialog - part 1
    // showJwtInputDialog({ data: "Test JWT" }).then(
    (jwtInput) => {
      spinnerOn();
      return identClient.acceptWorkgroupInvitation(jwtInput.jwt, jwtInput.orgId).then(() => {
        spinnerOff();
        alerts.success("Invitation completed");
      }, onError);
    },
    () => {
      /* NOTE: On cancel - do nothing */
    }
  );
}

function fillMyWorkgroupSheet(): Promise<void> {
  if (!identClient) {
    setUiForLogin();
    return;
  }

  spinnerOn();
  return identClient.getWorkgroups().then((apps) => {
    return excelWorker.showWorkgroups("My Workgroups", apps, true).then(spinnerOff);
  }, onError);
}

function startBaselining(): Promise<void> {
  if (!identClient) {
    setUiForLogin();
    return;
  }
  return excelWorker.startBaselineService(identClient);
}

function onSetupBaselining(): Promise<unknown> {
  return initializeBaselining();
}

function initializeBaselining(): Promise<unknown> {
  if(!identClient) {
    setUiForLogin();
    return;
  }

  spinnerOn();
  return excelWorker.createInitialSetup().then(spinnerOff, onError);
}