import { alerts, spinnerOff } from "./alerts";

export type Jwtoken = string;

/* global Office */

export function onError(reason: any) {
  let message = reason.toString();
  console.log(message);
  if (message.indexOf("Error: ") == 0) {
    message = message.substring("Error: ".length);
  }

  alerts.error(message);
  spinnerOff();
}

export enum DialogEvent { 
  // eslint-disable-next-line no-unused-vars
  Ok, Cancel, Initialized
}

export function closeCanceledDialog() {
  Office.context.ui.messageParent(
    JSON.stringify({
      result: DialogEvent.Cancel,
    })
  );
}

export function closeDialog(data: any) {
  Office.context.ui.messageParent(
    JSON.stringify({
      result: DialogEvent.Ok,
      data: data,
    })
  );
}

export function initializedDialog() {
  Office.context.ui.messageParent(
    JSON.stringify({
      result: DialogEvent.Initialized
    })
  );
}
