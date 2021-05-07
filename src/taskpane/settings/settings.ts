
/* global Office */

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

export const documentSettings: ISettingsStorage = new DocumentSettings();
export const localStorageSettings: ISettingsStorage = new LocalStorageSettings();
export const sessionStorageSettings: ISettingsStorage = new SessionStorageSettings();