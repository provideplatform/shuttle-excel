// NOTE: The wrapper functions of opening dialog for decrease "complexity" of opening :-)

import { TokenStr } from "../models/common";
import { showDialog } from "./dialogs";
import { JwtInputDialogUrl } from "./jwtInputDialog";
import { JwtInputDialogV01Url } from "./jwtInputDialog_V0.1";
import { JwtInputResult } from "./models/jwt-input-data";


// NOTE: data - for demo only!
export function showJwtInputDialog(data: any): Promise<JwtInputResult> {
    return showDialog<JwtInputResult>(JwtInputDialogUrl, { height: 38, width: 35 }, data);
}

// NOTE: Demo function
export function showJwtInputDialoV01(data: any): Promise<TokenStr> {
    return showDialog<TokenStr>(JwtInputDialogV01Url, { height: 38, width: 35 }, data);
}