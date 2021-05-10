import { DateStr, Email, Guid } from "./common";

export interface User {
    id: Guid;
    name: string;
    email: Email;
}

export interface ServerUser extends User {
    first_name: string;
    last_name: string;
    created_at: DateStr;
    
    permissions: number;
    privacy_policy_agreed_at: DateStr;
    terms_of_service_agreed_at: DateStr;
}

