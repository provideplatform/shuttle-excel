import { Ident } from 'provide-js';
import { AuthenticationResponse, Token, User } from '@provide/types';

export interface IdentClient {
    getToken(): Promise<string>;
    getUserFullName(): Promise<string>;
}

class IdentClientImpl implements IdentClient {
    // eslint-disable-next-line no-unused-vars
    constructor(private token: Token, private user: User) {}

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
        let fullName = [this.user?.firstName, this.user?.lastName].join(' ');
        return Promise.resolve(fullName);
    }

    private isExpired() {
        // TODO: 
        return false;
    }

    private refresh(): Promise<void> {
        let identService = new Ident(this.token.refreshToken);
        let params = { grant_type: 'refresh_token' };
        return identService.createToken(params).then(token => {
            this.token = token;
        });
    }
}

export function authenticate(email: string, password: string): Promise<IdentClient> {
    // NOTE: To test 
    let token: Token = {
        accessToken: 'qwertyuiop',
        marshal() { return 'marshal'; },
        // eslint-disable-next-line no-unused-vars
        unmarshal(json: string) {  }
    }
    let user: User = {
        firstName: 'Test'+ email,
        lastName: 'User'+ password,
        email: email,
        marshal() { return 'marshal'; },
        // eslint-disable-next-line no-unused-vars
        unmarshal(json: string) {  }
    }

    let resp: AuthenticationResponse = {
        token: token,
        user: user
    };

    return Promise.resolve(new IdentClientImpl(resp.token, resp.user));

    // let params = { email, password};
    // return Ident.authenticate(params)
    //     .then((response: AuthenticationResponse) => {
    //         return new IdentClientImpl(response.token, response.user);
    //     });
}
