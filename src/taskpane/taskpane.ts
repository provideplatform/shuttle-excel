import { IdentClient, authenticate } from "./identClient";

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";


// eslint-disable-next-line no-unused-vars
/* global window, console, alert, document, Excel, Office, $ */

Office.onReady((info) => {
  // console.log('P1');
  // debugger;
  // document.getElementById('aaaaa').textContent = info.host.toString();
  
  if (info.host === Office.HostType.Excel) {
    // if (!Office.context.requirements.isSetSupported('ExcelApi', '1.8')) {
    //   document.getElementById('bbbbb').textContent = 'Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.';
    // } else {
    //   document.getElementById('bbbbb').textContent = 'Support 1.8.';
    // }

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
  let workUi = $("#work-ui");
  identClient.getUserFullName().then((userFullName) => {
    $("#user-name", workUi).text(userFullName);
    $("#login-ui").hide();
    workUi.show();
  });
}

let identClient: IdentClient;
function login() {
  let email = <string>$("#email").val();
  let password = <string>$("#password").val();
  console.log("login", email, password);
  authenticate(email, password).then((client) => {
    identClient = client;
    setUiAfterLogin();
  });
}
