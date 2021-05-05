import { Guid, TokenStr } from "./common";

export interface Token {
    id: Guid;
    expires_in: number;
    access_token: TokenStr;
    refresh_token: TokenStr;
    scope: string;
    permissions: number;
}