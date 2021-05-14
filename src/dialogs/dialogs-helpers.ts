// NOTE: The wrapper functions of opening dialog for decrease "complexity" of opening :-)

import { TokenStr } from "../models/common";
import { showDialog } from "./dialogs";
import { JwtInputResult } from "./models/jwt-input-data";

export const JwtInputDialogUrl = "https://localhost:3000/jwtInputDialog.html";
// NOTE: data - for demo only!
export function showJwtInputDialog(data: any): Promise<JwtInputResult> {
    return showDialog<JwtInputResult>(JwtInputDialogUrl, { height: 38, width: 35 }, data);
}

export const JwtInputDialogV01Url = "https://localhost:3000/jwtInputDialog_V0.1.html";
// NOTE: Demo function
export function showJwtInputDialoV01(data: any): Promise<TokenStr> {
    return showDialog<TokenStr>(JwtInputDialogV01Url, { height: 38, width: 35 }, data);
}