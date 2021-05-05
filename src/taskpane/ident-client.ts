import { Ident } from "provide-js";
import { AuthenticationResponse } from "@provide/types";
import { Token as _Token } from "./models/token";
import { User as _User } from "./models/user";

export interface IdentClient {
  readonly test_expiresAt: Date;
  readonly isExpired: boolean;

  getToken(): Promise<string>;
  getUserFullName(): Promise<string>;
  logout(): Promise<void>;
}

class IdentClientImpl implements IdentClient {
  private token: _Token;
  private user: _User;

  private expiresAt: Date;

  constructor(token: _Token, user: _User) {
    this.token = token;
    this.user = user;

    this.initExpiresAt();
  }

  get test_expiresAt(): Date {
    return this.expiresAt;
  }

  get isExpired(): boolean {
    return new Date() > this.expiresAt;
  }

  getToken(): Promise<string> {
    if (this.isExpired) {
      return this.refresh().then(() => {
        return this.token.access_token;
      });
    } else {
      return Promise.resolve(this.token.access_token);
    }
  }

  getUserFullName(): Promise<string> {
    let fullName = this.user?.name ?? [this.user?.first_name, this.user?.last_name].join(" ");
    return Promise.resolve(fullName);
  }

  logout(): Promise<void> {
    this.token = null;
    this.user = null;
    return Promise.resolve();
  }

  private initExpiresAt() {
    const expires_in = this.token.expires_in;
    this.expiresAt = new Date();
    this.expiresAt.setSeconds(this.expiresAt.getSeconds() + expires_in - 60);
  }

  private refresh(): Promise<void> {
    let identService = new Ident(this.token.refresh_token);
    let params = { grant_type: "refresh_token" };
    return identService.createToken(params).then((token) => {
      this.token = (token as any) as _Token;
    });
  }
}

export interface AuthParams {
  email: string;
  password: string;
}

export function authenticateStub(authParams: AuthParams): Promise<IdentClient> {
  let token: _Token = {
    id: "sdfgsdfg",
    access_token: "qwertyuiop",
    refresh_token: "sdfgsdfg",
    expires_in: 86400,
    permissions: 7553,
    scope: "test",
  };
  token["expires_in"] = 86400;
  let user: _User = {
    id: "sdfgsdfg",
    first_name: "Test" + authParams.email,
    last_name: "User" + authParams.password,
    name: null,
    email: authParams.email,
    created_at: "2020-12-07T03:50:02.826Z",
    privacy_policy_agreed_at: "2020-12-07T03:50:02.826Z",
    terms_of_service_agreed_at: "2020-12-07T03:50:02.826Z",
    permissions: 7553,
  };

  return Promise.resolve(new IdentClientImpl(token, user));
}

export function authenticate(authParams: AuthParams): Promise<IdentClient> {
  let params = {
    scope: "offline_access",
  };
  params = Object.assign(params, authParams);
  // debugger;
  return Ident.authenticate(params).then((response: AuthenticationResponse) => {
    let token = (response.token as any) as _Token;
    let user = (response.user as any) as _User;
    return new IdentClientImpl(token, user);
  });
}
