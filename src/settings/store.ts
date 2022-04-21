import { Record } from "src/models/record";
import { indexedDatabase as db } from "./settings";
import { TokenStr } from "src/models/common";
import { User } from "src/models/user";
import { crypto } from "./crypto";
import { alerts } from "../common/alerts";
export class Store {
  onDbOpen = false;

  //TODO --> WHEN to open
  async open(): Promise<void> {
    //Open DB and create tables
    this.onDbOpen = await db.open();
  }
  async close(): Promise<void> {
    await db.close();
  }

  async createTable(tableName: string, keypath: string, indexName?: string, indexPath?: string) {
    indexName
      ? await db.createObjectStore(tableName, keypath, indexName, indexPath)
      : await db.createObjectStore(tableName, keypath);
  }

  async setBaselineID(tableName: string, key: string, value: string) {
    const record = {};
    record["primaryKeyID"] = key;
    record["baselineID"] = value;

    await db.set(tableName, record);
  }

  async set(tableName, keyPath: string, key: string, value: any) {
    const record = {};
    record[keyPath] = key;
    record["value"] = value;

    await db.set(tableName, record);
  }

  async get(tableName, key): Promise<any> {
    const value = await db.get(tableName, key);
    return value;
  }

  async setPrimaryKeyColumnName(key: string, value: string): Promise<void> {
    const record = {
      tableID: key,
      primaryKey: value,
    };

    const tableName = "tablePrimaryKeys";
    await db.set(tableName, record);
  }

  async setTableID(key: string, value: string): Promise<void> {
    const record = {
      tableName: key,
      mappingTable: value,
    };

    const tableName = "tableNames";
    await db.set(tableName, record);
  }

  async getTableID(key: string): Promise<String> {
    var tableName: string = await db.get("tableNames", key);
    return tableName;
  }

  async getBaselineId(tableName: string, key: string): Promise<string> {
    var record: Record = await db.get(tableName, key);

    return record.baselineID;
  }

  async getPrimaryKeyId(tableName: string, key: string): Promise<any> {
    var record: Record = await db.getByIndex(tableName, "baselineID", key);

    return record.primaryKey;
  }

  async getPrimaryKeyColumnName(key: string): Promise<any> {
    const tableName = "tablePrimaryKeys";
    var record = await db.get(tableName, key);

    return record["primaryKey"];
  }

  async tableExists(dbObjectStoreName: string, tableName: string): Promise<boolean> {
    var tableExists = false;
    if (this.onDbOpen) {
      tableExists = await this.keyExists(dbObjectStoreName, tableName);
    }
    return tableExists;
  }

  async keyExists(tableName: string, key: any, indexName?: string): Promise<boolean> {
    var keyExists = indexName
      ? await this.checkRecord(tableName, key, indexName)
      : await this.checkRecord(tableName, key);

    return keyExists;
  }

  private async checkRecord(tableName: string, key?: any, indexName?: string): Promise<boolean> {
    var recordCount: number = indexName
      ? await db.countByIndex(tableName, indexName, key)
      : await db.count(tableName, key);
    if (recordCount > 0) {
      return true;
    } else {
      return false;
    }
  }

  async remove(tableName: string, key: string): Promise<void> {
    await db.remove(tableName, key);
  }

  async getRefreshToken(): Promise<TokenStr | null> {
    try {
      if (this.onDbOpen) {
        const encryptedUserInfo = await this.get("userInfo", "userInfo");
        const decryptedUserInfo: string = await crypto.decrypt(encryptedUserInfo);
        const value = JSON.parse(decryptedUserInfo).refreshToken;
        return (value as TokenStr) || null;
      }
      return null;
    } catch (e) {
      alerts.error(e);
    }
  }

  async getUser(): Promise<User | null> {
    try {
      if (this.onDbOpen) {
        const encryptedUserInfo = await this.get("userInfo", "userInfo");
        const decryptedUserInfo: string = await crypto.decrypt(encryptedUserInfo);
        const value = JSON.parse(decryptedUserInfo).user;
        return (value as User) || null;
      }
      return null;
    } catch (e) {
      alerts.error(e);
    }
  }

  async setTokenAndUser(token: TokenStr, user: User): Promise<void> {
    var userInfo = {
      refreshToken: token,
      user: user,
    };

    await crypto.setKey();
    var encryptedUserInfo = await crypto.encrypt(JSON.stringify(userInfo));
    await this.set("userInfo", "keyName", "userInfo", encryptedUserInfo);
  }

  async removeTokenAndUser(): Promise<void> {
    await store.remove("userInfo", "userInfo");
    await store.remove("userInfo", "cryptoKey");
    await store.remove("userInfo", "cryptoIv");
  }
}

export const store = new Store();
