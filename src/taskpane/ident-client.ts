import { Ident, identClientFactory } from "provide-js";
import { Application, AuthenticationResponse } from "@provide/types";
import { Token as _Token } from "./models/token";
import { User as _User } from "./models/user";

export interface IdentClient {
  readonly test_expiresAt: Date;
  readonly isExpired: boolean;

  getToken(): Promise<string>;
  getUserFullName(): Promise<string>;
  logout(): Promise<void>;

  getWorkgroups(): Promise<Application[]>;
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
    // return true;
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

  getWorkgroups(): Promise<Application[]> {
    return this.executeWithAccessToken((accessToken) => {
      let identService = identClientFactory(accessToken);
      return identService.fetchApplications({});
    })
  }

  private initExpiresAt() {
    const expires_in = this.token.expires_in;
    this.expiresAt = new Date();
    this.expiresAt.setSeconds(this.expiresAt.getSeconds() + expires_in - 60);
  }

  private refresh(): Promise<void> {
    let identService = identClientFactory(this.token.refresh_token);
    let params = { grant_type: "refresh_token" };
    return identService.createToken(params).then(token => {
      this.token.id = token.id;
      this.token.access_token = token.accessToken;
      this.token.expires_in = token["expiresIn"];
      this.token.permissions = token["permissions"];

      this.initExpiresAt();
    });
  }

  // eslint-disable-next-line no-unused-vars
  private async executeWithAccessToken<T>(action: (accessToken: string) => Promise<T>) {
    const accessToken = await this.getToken();
    return await action(accessToken);
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
