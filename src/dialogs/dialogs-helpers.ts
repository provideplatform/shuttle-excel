// NOTE: The wrapper functions of opening dialog for decrease "complexity" of opening :-)

import { showDialog } from "./dialogs";
import { JwtInputResult } from "./models/jwt-input-data";

export const JwtInputDialogUrl = "jwtInputDialog.html";
export function showJwtInputDialog(data?: any): Promise<JwtInputResult> {
  const url = getDialogUrl(JwtInputDialogUrl);
  return showDialog<JwtInputResult>(url, { height: 38, width: 35 }, data);
}

function getDialogUrl(dialogPage: string): string {
    const currentPage = getCurrentPage();
    const url = window.location.href.replace(currentPage, dialogPage);
    return removeSearchPart(url);
}

function removeSearchPart(url: string) {
    return url.substr(0, url.length - window.location.search.length);
}

function getCurrentPage(): string {
    return window.location.pathname.substr(window.location.pathname.lastIndexOf("/") + 1);
}