// import { Ident, identClientFactory, nchainClientFactory, vaultClientFactory } from "provide-js";
// import { Application } from "@provide/types";
// import { Token as _Token } from "../models/token";
// import { ServerUser as _User, User } from "../models/user";
// import { Jwtoken } from "../taskpane/common";
// import { Uuid, TokenStr } from "../models/common";
// import * as jwt from 'jsonwebtoken';
// import { AuthParams } from "../models/auth-params";

// export interface IdentClient {
//   readonly test_expiresAt: Date;

//   readonly user: User;
//   readonly isExpired: boolean;
//   readonly refreshToken: TokenStr;

//   getToken(): Promise<string>;
//   logout(): Promise<void>;

//   getWorkgroups(): Promise<Application[]>;

//   // eslint-disable-next-line no-unused-vars
//   acceptWorkgroupInvitation(jwt: Jwtoken): Promise<void>;


//   testA(): Promise<string>;
//   testB(): Promise<string>;
// }

// class IdentClientImpl implements IdentClient {
//   private _userToken: _Token;
//   private _user: User;

//   private expiresAt: Date;

//   constructor(token: _Token, user: User) {
//     this._userToken = token;
//     this._user = user;

//     this.initExpiresAt();
//   }

//   get test_expiresAt(): Date {
//     return this.expiresAt;
//   }

//   get user(): User {
//     return this._user;
//   }

//   get isExpired(): boolean {
//     // return true;
//     return new Date() > this.expiresAt;
//   }

//   get refreshToken(): TokenStr {
//     return this._userToken.refresh_token as TokenStr;
//   } 

//   getToken(): Promise<string> {
//     if (this.isExpired) {
//       return this.refresh().then(() => {
//         return this._userToken.access_token;
//       });
//     } else {
//       return Promise.resolve(this._userToken.access_token);
//     }
//   }

//   logout(): Promise<void> {
//     this._userToken = null;
//     this._user = null;
//     return Promise.resolve();
//   }

//   getWorkgroups(): Promise<Application[]> {
//     return this.executeWithUserAccessToken((accessToken) => {
//       let identService = identClientFactory(accessToken);
//       return identService.fetchApplications({});
//     })
//   }

//   acceptWorkgroupInvitation(inviteToken: Jwtoken): Promise<void> {
//     console.log("JWT:", inviteToken);

//     // debugger;

//     // // TODO: 
//     // const OrganizationId = 'asdfgasdfasdfasdfasd';

//     // // NOTE: Parse inviteToken
//     // const invite = jwt.decode(inviteToken) as { [key: string]: any };
//     // console.log(JSON.stringify(invite));

//     // const workgroupAndApplicationId = invite.baseline.workgroup_id;
//     // console.log("workgroup and application Id:", workgroupAndApplicationId);


//     // // // TODO: 
//     // // // NOTE: AuthorizeOrganizationContext
//     // // const userAccessToken = 'asdfsdfgsdfg';
//     // // let identService = identClientFactory(userAccessToken);
//     // // const params = {
//     // //   "scope": "offline_access",
//     // //   "organization_id": OrganizationId,
//     // // };

//     // const organizationAccessToken = identService.createToken(params);

//     // // // NOTE: authorizeApplicationContext
//     // // // NOTE: authorizeApplicationContext AuthorizeApplicationContext
//     // // const userAccessToken2 = 'asdfsdfgsdfg';
//     // // let identService2 = identClientFactory(userAccessToken2);
//     // // const params2 = {
//     // //   "scope":          "offline_access",
// 		// //   "application_id": workgroupAndApplicationId,
//     // // };
//     // // const applicationAccessToken = identService2.createToken(params2);

//     // // // NOTE: authorizeApplicationContext authorize in 
//     // // const nchainService = nchainClientFactory("applicationAccessToken");
//     // // const params3 = {
//     // //   "purpose": 44,
//     // // }
//     // // var notUsed = nchainService.createWallet(params3);

//     // // NOTE: common.RequireOrganizationVault()
//     // // NOTE: Читаем списко Vaults
//     // const vaultService = vaultClientFactory("organizationAccessToken");
//     // const params4 = {
//     //   "organization_id": OrganizationId,
//     // }
//     // const listVaults = vaultService.fetchVaults(params4);

//     // // NOTE: Если списко listVaults пустой - создаем новый
//     // const params5 = {
//     //   "name":        `vault for organization: ${OrganizationId}`,
// 		// 	"description": `identity/signing keystore for organization: ${OrganizationId}`
//     // };
//     // const vault = vaultService.createVault(params5);
//     // // NOTE: Если списко listVaults НЕ пустой
//     // const vault = listVaults as any[];

//     // const VaultId = vault.Id;

//     // // NOTE: requireOrganizationKeys


//     return Promise.resolve();
//   }

//   private initExpiresAt() {
//     const expires_in = this._userToken.expires_in;
//     this.expiresAt = new Date();
//     this.expiresAt.setSeconds(this.expiresAt.getSeconds() + expires_in - 60);
//   }

//   private refresh(): Promise<void> {
//     let identService = identClientFactory(this._userToken.refresh_token);
//     let params = { grant_type: "refresh_token" };
//     return identService.createToken(params).then(token => {
//       this._userToken.id = token.id;
//       this._userToken.access_token = token.accessToken;
//       this._userToken.expires_in = token["expiresIn"];
//       this._userToken.permissions = token["permissions"];

//       this.initExpiresAt();
//     });
//   }

//   // eslint-disable-next-line no-unused-vars
//   private async executeWithUserAccessToken<T>(action: (accessToken: string) => Promise<T>) {
//     const accessToken = await this.getToken();
//     return await action(accessToken);
//   }

//   private myTimeout(): Promise<void> {
//     return new Promise((resolve, reject) => {
//       setTimeout(() => { resolve(); }, 1000);
//     });
//   }
// }

// export function authenticateStub(authParams: AuthParams): Promise<IdentClient> {
//   let token: _Token = {
//     id: "sdfgsdfg",
//     access_token: "qwertyuiop",
//     refresh_token: "sdfgsdfg",
//     expires_in: 86400,
//     permissions: 7553,
//     scope: "test",
//   };
//   token["expires_in"] = 86400;
//   let user: _User = {
//     id: "sdfgsdfg",
//     first_name: "Test" + authParams.email,
//     last_name: "User" + authParams.password,
//     name: "Test" + authParams.email + " " + "User" + authParams.password,
//     email: authParams.email,
//     created_at: "2020-12-07T03:50:02.826Z",
//     privacy_policy_agreed_at: "2020-12-07T03:50:02.826Z",
//     terms_of_service_agreed_at: "2020-12-07T03:50:02.826Z",
//     permissions: 7553,
//   };

//   return Promise.resolve(new IdentClientImpl(token, user));
// }

// export async function authenticate(authParams: AuthParams): Promise<IdentClient> {
//   let params = {
//     scope: "offline_access",
//   };
//   params = Object.assign(params, authParams);
//   const response = await Ident.authenticate(params);
//   let token = (response.token as any) as _Token;
//   let user = (response.user as any) as _User;
//   return new IdentClientImpl(token, user);
// }

// // eslint-disable-next-line no-unused-vars
// export function restoreStub(refreshToken: TokenStr, user: User): Promise<IdentClient> {
//   const fakeauthParams = { email: "email", password: "password" };
//   return authenticateStub(fakeauthParams);
// }

// export async function restore(refreshToken: TokenStr, user: User): Promise<IdentClient> {
//   const token = { refresh_token: refreshToken, expires_in: 0 };
//   const identClient = new IdentClientImpl(token as _Token, user);
//   // NOTE: Call "getToken" to get accessToken by refreshToken.
//   await identClient.getToken();
//   return identClient;
// }
