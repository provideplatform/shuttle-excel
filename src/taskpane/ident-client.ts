import { Ident } from "provide-js";
import { AuthenticationResponse, Token, User } from "@provide/types";

export interface IdentClient {
  getToken(): Promise<string>;
  getUserFullName(): Promise<string>;
  logout(): Promise<void>;
}

class IdentClientImpl implements IdentClient {
  private token: Token;
  private user: User;

  constructor(token: Token, user: User) {
    this.token = token;
    this.user = user;
  }

  getToken(): Promise<string> {
    if (this.isExpired()) {
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

  private isExpired() {
    // TODO:
    return false;
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
      unmarshal(json: string) {  }
  }
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
  return Ident.authenticate(authParams).then(
    (response: AuthenticationResponse) => {
      // debugger;
      return new IdentClientImpl(response.token, response.user);
    }
  );
}
