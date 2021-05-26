// NOTE: Incapsulate working with storages

/* global Office */

import { TokenStr } from "../models/common";
import { User } from "../models/user";

export interface ISettingsStorage {
  // eslint-disable-next-line no-unused-vars
  set(key: string, value: any): Promise<void>;
  // eslint-disable-next-line no-unused-vars
  get(key: string): Promise<any>;
  // eslint-disable-next-line no-unused-vars
  remove(key: string): Promise<void>;
}

class DocumentSettings implements ISettingsStorage {
  set(key: string, value: any): Promise<void> {
    Office.context.document.settings.set(key, value);
    return this.save();
  }
  get(key: string): Promise<any> {
    const value = Office.context.document.settings.get(key);
    return Promise.resolve(value);
  }
  remove(key: string): Promise<void> {
    Office.context.document.settings.remove(key);
    return this.save();
  }

  private save(): Promise<void> {
    const promise = new Promise<void>((resolve, reject) => {
      Office.context.document.settings.saveAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
          // reject('Settings save failed. Error: ' + asyncResult.error.message);
          reject(asyncResult.error.message);
        } else {
          resolve();
        }
      });
    });

    return promise;
  }
}

abstract class StorageSettings implements ISettingsStorage {
  protected storage: Storage;

  constructor(storage: Storage) {
    this.storage = storage;
  }

  set(key: string, value: any): Promise<void> {
    this.storage.setItem(key, JSON.stringify(value));
    return Promise.resolve();
  }
  get(key: string): Promise<any> {
    const valueJson = this.storage.getItem(key);
    if (valueJson === "") {
      return Promise.resolve("");
    }
    if (valueJson) {
      const retVal = JSON.parse(valueJson);
      return Promise.resolve(retVal);
    }

    return Promise.resolve(null);
  }
  remove(key: string): Promise<void> {
    this.storage.removeItem(key);
    return Promise.resolve();
  }
}

class LocalStorageSettings extends StorageSettings {
  constructor() {
    super(window.localStorage);
  }
}

class SessionStorageSettings extends StorageSettings {
  constructor() {
    super(window.sessionStorage);
  }
}

const documentSettings: ISettingsStorage = new DocumentSettings();
// eslint-disable-next-line no-unused-vars
const localStorageSettings: ISettingsStorage = new LocalStorageSettings();
// eslint-disable-next-line no-unused-vars
const sessionStorageSettings: ISettingsStorage = new SessionStorageSettings();

class Settings {
  private readonly NAME = "__docSettings";

  async getRefreshToken(): Promise<TokenStr | null> {
    const value = await documentSettings.get(this.NAME);
    return (value && value["refreshToken"] as TokenStr) || null;
  }

  async getUser(): Promise<User | null> {
    const value = await documentSettings.get(this.NAME);
    return (value &&  value["user"] as User) || null;
  }

  async setTokenAndUser(token: TokenStr, user: User): Promise<void> {
    const settingObj = (await documentSettings.get(this.NAME)) || {};
    settingObj["refreshToken"] = token;
    settingObj["user"] = user;
    await documentSettings.set(this.NAME, settingObj);
  }

  async removeTokenAndUser(): Promise<void> {
    const settingObj = (await documentSettings.get(this.NAME)) || {};
    delete settingObj["refreshToken"];
    delete settingObj["user"];
    await documentSettings.set(this.NAME, settingObj);
  }
}

export const settings = new Settings();
