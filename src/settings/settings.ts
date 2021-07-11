// NOTE: Incapsulate working with storages

/* global Office */

import { TokenStr } from "../models/common";
import { User } from "../models/user";
import { Record } from "../models/record";

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

class IndexedDBSettings {
  protected db: IDBDatabase;
  private database: string;
  private version: number;

  constructor(database: string) {
    this.database = database;
  }
  async createObjectStore(tableName: string): Promise<void> {
    try {
      
      //Create or open the database
      this.version++;
      await this.db.close();
      const request = await indexedDB.open(this.database, this.version);

      //on upgrade needed, create object store
      request.onupgradeneeded = async (e) => {
        this.db = (<IDBOpenDBRequest>e.target).result;
        await this.db.createObjectStore(tableName+"Out", { keyPath: ["primaryKey", "columnName"] });
        await this.db.createObjectStore(tableName+"In", { keyPath: "baselineID" });
        if(!this.db.objectStoreNames.contains("tablePrimaryKeys")) {
          await this.db.createObjectStore("tablePrimaryKeys", { keyPath: "tableID" });}
        };

      //on success
      request.onsuccess = (e) => {
        this.db = (<IDBOpenDBRequest>e.target).result;
        this.version = this.db.version;
      };

      //on error
      request.onerror = (e) => {
        console.log((<IDBOpenDBRequest>e.target).error);
      };
    } catch (error) {
      return;
    }
  }

  async openDB(tableName: string): Promise<unknown> {
    try {
      //Open database
      //ONLY FOR TESTS
      //await indexedDB.deleteDatabase(this.database); 
      const request = await indexedDB.open(this.database);
      
      //on upgrade needed, create object store
      request.onupgradeneeded = async (e) => {
        this.db = (<IDBOpenDBRequest>e.target).result;
        if (!this.db.objectStoreNames.contains("tablePrimaryKeys")) {
          await this.db.createObjectStore("tablePrimaryKeys", { keyPath: "tableID" });
        }
        this.version = e.newVersion;
      };

      var tableExist = await new Promise((resolve, reject) => {
        //on success
        request.onsuccess = async (e) => {
          this.db = (<IDBOpenDBRequest>e.target).result;
          var bool = await this.tableExists(tableName);
          this.version = this.db.version;
          console.log("Database version: " + this.version);
          resolve(bool);
        };

        //on error
        request.onerror = (e) => {
          console.log((<IDBOpenDBRequest>e.target).error);
          reject(request.error)
        };
      });

      return tableExist;
      
    } catch (error) {
      return;
    }
  }
  async set(tableName: string, key: string[], value: string): Promise<void> {
    await this.setOutboundTable(tableName, key, value);
    await this.setInboundTable(tableName, value, key);
  }

  async setOutboundTable(tableName: string, key: string[], value: string): Promise<void> {
    const record: Record = {
      primaryKey: key[0],
      columnName: key[1],
      baselineID: value,
    };
    const tx = this.db.transaction(tableName+"Out", "readwrite");
    const store = tx.objectStore(tableName+"Out");
    store.put(record);
  }

  async setInboundTable(tableName: string, key: string, value: string[]): Promise<void> {
    const record: Record = {
      baselineID: key,
      primaryKey: value[0],
      columnName: value[1], 
    };
    const tx = this.db.transaction(tableName+"In", "readwrite");
    const store = tx.objectStore(tableName+"In");
    store.put(record);
  }

  async setPrimaryKey(key: string, value: string): Promise<void> {
    const record = {
      tableID : key,
      primaryKey : value
    }
    const tx = this.db.transaction("tablePrimaryKeys", "readwrite");
    const store = tx.objectStore("tablePrimaryKeys");
    store.put(record);
  }

  async get(tableName: string, key: string[]): Promise<any> {
    var record: Record = await new Promise((resolve, reject) => {
      const tx = this.db.transaction(tableName+"Out", "readonly");
      const store = tx.objectStore(tableName+"Out");
      const request = store.get(key);

      request.onsuccess = () => {
        resolve(request.result);
      };

      request.onerror = () => {
        reject(request.error);
      };
    });

    return record.baselineID;
  }

  async getKey(tableName: string, key: string): Promise<any> {
    var record: Record = await new Promise((resolve, reject) => {
      const tx = this.db.transaction(tableName+"In", "readonly");
      const store = tx.objectStore(tableName+"In");
      const request = store.get(key);

      request.onsuccess = () => {
        resolve(request.result);
      };

      request.onerror = () => {
        reject(request.error);
      };
    });

    return [record.primaryKey, record.columnName];
  }

  async getPrimaryKey(key: string) : Promise<any> {
    var primaryKey = await new Promise((resolve, reject) => {
      const tx = this.db.transaction("tablePrimaryKeys", "readonly"); 
      const store = tx.objectStore("tablePrimaryKeys");
      const request = store.get(key);

      request.onsuccess = () => {
        resolve(request.result);
      };

      request.onerror = () => {
        console.log(request.error);
        reject(request.error);
      };
    });
    
    return primaryKey; 
  }

  private async tableExists(tableName: string) : Promise<boolean> {
  
    var recordCount: number = await new Promise((resolve, reject) => {
      
      const tx = this.db.transaction("tablePrimaryKeys", "readonly"); 
      const store = tx.objectStore("tablePrimaryKeys");
      const request = store.count(tableName);

      request.onsuccess = () => {
        resolve(request.result);
      };

      request.onerror = () => {
        reject(request.error);
      };
    });

    if (recordCount > 0) {
      return true;
    } else {
      return false;
    }

  }

  async recordExists(tableName: string, key: string[]): Promise<boolean> {
    var recordCount: number = await new Promise((resolve, reject) => {
      const tx = this.db.transaction(tableName+"Out", "readonly");
      const store = tx.objectStore(tableName+"Out");
      const request = store.count(key);

      request.onsuccess = () => {
        resolve(request.result);
      };

      request.onerror = () => {
        reject(request.error);
      };
    });

    if (recordCount > 0) {
      return true;
    } else {
      return false;
    }
  }

  async remove(tableName: string, key: string[]): Promise<void> {
    const tx = this.db.transaction(tableName, "readwrite");
    const store = tx.objectStore(tableName);
    const result = await store.get(key);
    if (!result) {
      console.log("Key not found", key);
    }
    await store.delete(key);
    console.log("Data Deleted", key);
    return;
  }
}

const documentSettings: ISettingsStorage = new DocumentSettings();
// eslint-disable-next-line no-unused-vars
const localStorageSettings: ISettingsStorage = new LocalStorageSettings();

const sessionStorageSettings: ISettingsStorage = new SessionStorageSettings();

class Settings {
  private readonly NAME = "__docSettings";

  async getRefreshToken(): Promise<TokenStr | null> {
    const value = await documentSettings.get(this.NAME);
    return (value && (value["refreshToken"] as TokenStr)) || null;
  }

  async getUser(): Promise<User | null> {
    const value = await documentSettings.get(this.NAME);
    return (value && (value["user"] as User)) || null;
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

class SessionSettings {

  async getRefreshToken(): Promise<TokenStr | null> {
    const value = await sessionStorageSettings.get("refreshToken");
    return (value as TokenStr) || null;
  }

  async getUser(): Promise<User | null> {
    const value = await sessionStorageSettings.get("user");
    return (value as User) || null;
  }

  async setTokenAndUser(token: TokenStr, user: User): Promise<void> {
    await sessionStorageSettings.set("refreshToken", token);
    await sessionStorageSettings.set("user", user);
  }

  async removeTokenAndUser(): Promise<void> {
    await sessionStorageSettings.remove("refreshToken");
    await sessionStorageSettings.remove("user");
  } 
}

export const settings = new Settings();
export const indexedDatabase = new IndexedDBSettings("baselineDB");
export const sessionSettings = new SessionSettings();
