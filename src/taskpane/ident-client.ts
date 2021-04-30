import { Ident } from "provide-js";
import { AuthenticationResponse, Token, User, Model } from "@provide/types";

export interface IdentClient {
  readonly test_expiresAt: Date;
  readonly isExpired: boolean;

  getToken(): Promise<string>;
  getUserFullName(): Promise<string>;
  logout(): Promise<void>;
}

class IdentClientImpl implements IdentClient {
  private token: Token;
  private user: User;

  private expiresAt: Date;

  constructor(token: Token, user: User) {
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
        return this.token.accessToken;
      });
    } else {
      return Promise.resolve(this.token.accessToken);
    }
  }

  getUserFullName(): Promise<string> {
    let fullName = [this.user?.firstName, this.user?.lastName].join(" ");
    return Promise.resolve(fullName);
  }

  logout(): Promise<void> {
    let identService = new Ident(this.token.accessToken);
    return identService.deleteToken(this.token.id);
  }

  private initExpiresAt() {
    const expires_in = <number>this.token['expires_in'];
    this.expiresAt = new Date();
    this.expiresAt.setSeconds(this.expiresAt.getSeconds() + expires_in - 60);
  }

  private refresh(): Promise<void> {
    let identService = new Ident(this.token.refreshToken);
    let params = { grant_type: "refresh_token" };
    return identService.createToken(params).then((token) => {
      this.token = token;
    });
  }
}

export interface AuthParams {
    email: string;
    password: string;
}

export function authenticateStub(authParams: AuthParams): Promise<IdentClient> {
  let token: Token = {
      accessToken: 'qwertyuiop',
      marshal() { return 'marshal'; },
      // eslint-disable-next-line no-unused-vars
      unmarshal(json: string) {  },
  }
  token['expires_in'] = 86400;
  let user: User = {
      firstName: 'Test'+ authParams.email,
      lastName: 'User'+ authParams.password,
      email: authParams.email,
      marshal() { return 'marshal'; },
      // eslint-disable-next-line no-unused-vars
      unmarshal(json: string) {  }
  }

  let resp: AuthenticationResponse = {
      token: token,
      user: user
  };
  return Promise.resolve(new IdentClientImpl(resp.token, resp.user));
}

export function authenticate(authParams: AuthParams): Promise<IdentClient> {
  let params = {
    scope: "offline_access"
  };
  params = Object.assign(params, authParams);
  // debugger;
  return Ident.authenticate(params).then(
    (response: AuthenticationResponse) => {
      // debugger;
      // let t = new Model();
      // let ttt = response.token as Token;

      // response.token.expiresAt;
      // ttt.expiresAt;

      // response.token.unmarshal(JSON.stringify(response.token));

      // t.unmarshal(JSON.stringify(response.token));
      // let realToken = t as Token;
      // (realToken as any).issuedAt;
      return new IdentClientImpl(response.token, response.user);

    }
  );
}
