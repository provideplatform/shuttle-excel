import { Application, Contract, Organization, Key, Wallet, Vault as ProvideVault, BaselineResponse, BusinessObject } from "@provide/types";
import { Ident, identClientFactory, nchainClientFactory, vaultClientFactory, baselineClientFactory } from "provide-js";
import { Uuid, TokenStr } from "../models/common";
import { AuthParams } from "../models/auth-params";
import * as jwt from "jsonwebtoken";
import { AuthContext } from "./auth-context";
import { AccessToken } from "./access-token";

import { Token as _Token } from "../models/token";
import { ServerUser as _User, User } from "../models/user";

export interface ProvideClient {
  readonly user: User;
  readonly userRefreshToken: TokenStr;

  logout(): Promise<void>;

  getWorkgroups(): Promise<Application[]>;

  // eslint-disable-next-line no-unused-vars
  sendCreateProtocolMessage(message: BusinessObject): Promise<BaselineResponse>;

  // eslint-disable-next-line no-unused-vars
  sendUpdateProtocolMessage(baselineID: string, message: BusinessObject): Promise<BaselineResponse>;

  // eslint-disable-next-line no-unused-vars
  acceptWorkgroupInvitation(inviteToken: TokenStr, organizationId: Uuid): Promise<void>;
}

class ProvideClientImpl implements ProvideClient {
  private _user: User;
  private _userAuthContext: AuthContext;
  private _appAuthContext: AuthContext;
  private _orgAuthContext: AuthContext;

  constructor(user: User, userAuthContext: AuthContext) {
    this._user = user;
    this._userAuthContext = userAuthContext;
  }

  get user(): User {
    return this._user;
  }

  get userRefreshToken(): TokenStr {
    return this._userAuthContext.refreshToken;
  }

  logout(): Promise<void> {
    // TODO: Maybe must call deleteToken?

    this._user = null;
    this._userAuthContext = null;
    return Promise.resolve();
  }

  async authorizeOrganization(organizationId: Uuid): Promise<void> {
    await this._userAuthContext.execute((accessToken) => {
      const identService = identClientFactory(accessToken);
      const params = {
        scope: "offline_access",
        organization_id: organizationId,
      };
      return identService.createToken(params).then((token) => {
        const expiresIn = parseInt(token["expiresIn"]);
        const accessToken = new AccessToken(token.id, token.accessToken, expiresIn);
        this._orgAuthContext = new AuthContext(token.refreshToken, accessToken);
      });
    });
  }

  async authorizeApplication(workgroupAndApplicationId: Uuid): Promise<void> {
    await this._userAuthContext.execute((accessToken) => {
      const identService = identClientFactory(accessToken);
      const params = {
        scope: "offline_access",
        application_id: workgroupAndApplicationId,
      };
      return identService.createToken(params).then((token) => {
        const expiresIn = parseInt(token["expiresIn"]);
        const accessToken = new AccessToken(token.id, token.accessToken, expiresIn);
        this._appAuthContext = new AuthContext(token.refreshToken, accessToken);
      });
    });
  }

  async getWorkgroups(): Promise<Application[]> {
    const retVal = await this._userAuthContext.get((accessToken) => {
      const identService = identClientFactory(accessToken);
      return identService.fetchApplications({});
    });
    return retVal;
  }

  async sendCreateProtocolMessage(message: BusinessObject): Promise<BaselineResponse> {
    const retVal = await this._orgAuthContext.get((accessToken) => {
      const baselineService = baselineClientFactory(accessToken, "https", "a4246b28e8d9.ngrok.io");
      return baselineService.createBusinessObject(message);
    });
    return retVal;
  }

  async sendUpdateProtocolMessage(baselineID: string, message: BusinessObject): Promise<BaselineResponse> {
    const retVal = await this._orgAuthContext.get((accessToken) => {
      const baselineService = baselineClientFactory(accessToken,"https", "a4246b28e8d9.ngrok.io");
      return baselineService.updateBusinessObject(baselineID, message);
    });
    return retVal;
  }
  async createWallet(): Promise<Wallet> {
    this.guardNotNullAppAuthContext();

    const retVal = await this._appAuthContext.get((accessToken) => {
      const nchainService = nchainClientFactory(accessToken);
      const params = {
        purpose: 44,
      };
      return nchainService.createWallet(params);
    });
    return retVal;
  }

  async getFirstOrCreateVault(organizationId: Uuid): Promise<ProvideVault> {
    const vaults = await this.getVaults(organizationId);
    if (vaults.length > 0) {
      return vaults[0];
    }

    const newVault = await this.createVault(organizationId);
    return newVault;
  }

  async getVaults(organizationId: Uuid): Promise<ProvideVault[]> {
    this.guardNotNullOrgAuthContext();

    const retVal = await this._orgAuthContext.get((accessToken) => {
      const vaultService = vaultClientFactory(accessToken);
      const params = {
        organization_id: organizationId,
      };
      return vaultService.fetchVaults(params);
    });
    return retVal;
  }

  async createVault(organizationId: Uuid): Promise<ProvideVault> {
    this.guardNotNullOrgAuthContext();

    const retVal = await this._orgAuthContext.get((accessToken) => {
      const vaultService = vaultClientFactory(accessToken);
      const params = {
        name: `vault for organization: ${organizationId}`,
        description: `identity/signing keystore for organization: ${organizationId}`,
      };
      return vaultService.createVault(params);
    });
    return retVal;
  }

  async getFirstOrCreateVaultKey(vaultId: Uuid, spec: string, organizationId: Uuid): Promise<Key> {
    const keys = await this.getVaultKeys(vaultId, spec);
    if (keys.length > 0) {
      return keys[0];
    }

    const newKey = await this.createVaultKey(vaultId, spec, organizationId);
    return newKey;
  }

  // TODO: Maybe it always only one.
  async getVaultKeys(vaultId: Uuid, spec: string): Promise<Key[]> {
    this.guardNotNullOrgAuthContext();

    const retVal = await this._orgAuthContext.get((accessToken) => {
      const vaultService = vaultClientFactory(accessToken);
      const params = {
        spec: spec,
      };
      return vaultService.fetchVaultKeys(vaultId, params);
    });
    return retVal;
  }

  async createVaultKey(vaultId: Uuid, spec: string, organizationId: Uuid): Promise<Key> {
    this.guardNotNullOrgAuthContext();

    const retVal = await this._orgAuthContext.get((accessToken) => {
      const vaultService = vaultClientFactory(accessToken);
      const params = {
        name: `${spec} key organization: ${organizationId}`,
        description: `${spec} key organization: ${organizationId}`,
        spec: spec,
        type: "asymmetric",
        usage: "sign/verify",
      };
      return vaultService.createVaultKey(vaultId, params);
    });
    return retVal;
  }

  async getContracts(contractType: string): Promise<Contract[]> {
    this.guardNotNullAppAuthContext();

    const retVal = await this._appAuthContext.get((accessToken) => {
      const nchainService = nchainClientFactory(accessToken);
      const params = {
        type: contractType,
      };
      return nchainService.fetchContracts(params);
    });
    return retVal;
  }

  async getOrCreateApplicationOrganization(
    workgroupAndApplicationId: Uuid,
    organizationId: Uuid
  ): Promise<Organization> {
    const applicationOrganizations = await this.getApplicationOrganizations(workgroupAndApplicationId, organizationId);

    // NOTE: I don't know why, but we must find organizationId... see RegisterWorkgroupOrganization in \provide-cli\cmd\common\baseline.go
    let applicationOrganization = applicationOrganizations.find((x) => x.id === organizationId);
    if (applicationOrganization) {
      return applicationOrganization;
    }

    applicationOrganization = await this.createApplicationOrganization(workgroupAndApplicationId, organizationId);
    return applicationOrganization;
  }

  async createOrGetApplicationOrganization(
    workgroupAndApplicationId: Uuid,
    organizationId: Uuid
  ): Promise<Organization> {
    try {
      const applicationOrganization = await this.createApplicationOrganization(
        workgroupAndApplicationId,
        organizationId
      );
      return applicationOrganization;
    } catch {
      const applicationOrganizations = await this.getApplicationOrganizations(
        workgroupAndApplicationId,
        organizationId
      );
      let applicationOrganization = applicationOrganizations.find((x) => x.id === organizationId);
      if (applicationOrganization) {
        return applicationOrganization;
      }

      throw "Organization not associated with workgroup";
    }
  }

  async createApplicationOrganization(workgroupAndApplicationId: Uuid, organizationId: Uuid): Promise<Organization> {
    this.guardNotNullAppAuthContext();

    const retVal = await this._appAuthContext.get((accessToken) => {
      const identService = identClientFactory(accessToken);
      const params = {
        organization_id: organizationId,
      };
      return identService.createApplicationOrganization(workgroupAndApplicationId, params);
    });
    return retVal;
  }

  async getApplicationOrganizations(workgroupAndApplicationId: Uuid, organizationId: Uuid): Promise<Organization[]> {
    this.guardNotNullAppAuthContext();

    const retVal = await this._appAuthContext.get((accessToken) => {
      const identService = identClientFactory(accessToken);
      const params = {
        organization_id: organizationId,
      };
      return identService.fetchApplicationOrganizations(workgroupAndApplicationId, params);
    });
    return retVal;
  }

  async acceptWorkgroupInvitation(inviteToken: TokenStr, organizationId: Uuid): Promise<void> {
    // TODO: requirePublicJWTVerifiers() from GO

    const inviteClaims = jwt.decode(inviteToken) as { [key: string]: any };
    if (!inviteClaims) {
      throw "Can't parse invite token";
    }
    const workgroupAndApplicationId = inviteClaims.baseline.workgroup_id as Uuid;

    await this.authorizeOrganization(organizationId);
    await this.authorizeApplication(workgroupAndApplicationId);
    // NOTE: I don't know for what, see: func authorizeApplicationContext() in \provide-cli\cmd\baseline\workgroups\workgroup_init.go
    await this.createWallet();

    const vault = await this.getFirstOrCreateVault(organizationId);

    /* NOTE: START - Where used? I don't know */
    // eslint-disable-next-line no-unused-vars
    const babyJubJubKey = await this.getFirstOrCreateVaultKey(vault.id, "babyJubJub", organizationId);
    // eslint-disable-next-line no-unused-vars
    const secp256k1Key = await this.getFirstOrCreateVaultKey(vault.id, "secp256k1", organizationId);
    // eslint-disable-next-line no-unused-vars
    const hdwalletKey = await this.getFirstOrCreateVaultKey(vault.id, "BIP39", organizationId);
    // eslint-disable-next-line no-unused-vars
    const rsa4096Key = await this.getFirstOrCreateVaultKey(vault.id, "RSA-4096", organizationId);
    /* NOTE: END */

    // common.RegisterWorkgroupOrganization(common.ApplicationID)
    const contracts = await this.getContracts("organization-registry");
    if (!contracts.length) {
      throw "Failed to initialize registry contract";
    }

    // NOTE: It's look like instance of applicationOrganization not required
    // eslint-disable-next-line no-unused-vars
    // const applicationOrganization = await this.getOrCreateApplicationOrganization(
    //   workgroupAndApplicationId,
    //   organizationId
    // );
    // eslint-disable-next-line no-unused-vars
    const applicationOrganization = await this.createOrGetApplicationOrganization(
      workgroupAndApplicationId,
      organizationId
    );

    // token := common.RequireAPIToken()
    // baseline.CreateWorkgroup(token, map[string]interface{}{
    // 	"token": jwt,
    // })
  }

  private guardNotNullAppAuthContext() {
    if (!this._appAuthContext) {
      throw "No application authorization";
    }
  }

  private guardNotNullOrgAuthContext() {
    if (!this._orgAuthContext) {
      throw "No organization authorization";
    }
  }
}

export function authenticateStub(authParams: AuthParams): Promise<ProvideClient> {
  const user: User = {
    id: "asdfgasdfgsdfgsdfg",
    email: authParams.email,
    name: "User Name " + authParams.password,
  };
  const userAccessToken = new AccessToken("id", "access_token", 86400);
  const userAuthContext = new AuthContext("refresh_token", userAccessToken);
  return Promise.resolve(new ProvideClientImpl(user, userAuthContext));
}

export async function authenticate(authParams: AuthParams): Promise<ProvideClient> {
  const params = Object.assign(
    {
      scope: "offline_access",
    },
    authParams
  );
  const response = await Ident.authenticate(params);
  let tokenResponse = (response.token as any) as _Token;
  let userResponse = (response.user as any) as _User;
  const user: User = userResponse as User;
  const userAccessToken = new AccessToken(tokenResponse.id, tokenResponse.access_token, tokenResponse.expires_in);
  const userAuthContext = new AuthContext(tokenResponse.refresh_token, userAccessToken);
  return new ProvideClientImpl(user, userAuthContext);
}

// eslint-disable-next-line no-unused-vars
export function restoreStub(refreshToken: TokenStr, user: User): Promise<ProvideClient> {
  const fakeauthParams = { email: "email", password: "password" };
  return authenticateStub(fakeauthParams);
}

export async function restore(refreshToken: TokenStr, user: User): Promise<ProvideClient> {
  const userAuthContext = new AuthContext(refreshToken);
  await userAuthContext.refresh();
  return new ProvideClientImpl(user, userAuthContext);
}
