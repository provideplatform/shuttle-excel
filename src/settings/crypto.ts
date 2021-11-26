import { store } from "./store";
export class CryptoKeys {
  //Create Keys
  async generateKey() {
    const result = await window.crypto.subtle.generateKey(
      {
        name: "AES-GCM",
        length: 256,
      },
      false,
      ["encrypt", "decrypt"]
    );

    return result;
  }
  //Set Keys
  async setKey() {
    const key = await this.generateKey();
    store.set("userInfo", "keyName", "cryptoKey", key);
  }

  //Get Keys
  async getKey(): Promise<CryptoKey> {
    const key = await store.get("userInfo", "cryptoKey");
    return key;
  }

  //Encrypt
  async encrypt(data) {
    const key = await this.getKey();
    const iv = window.crypto.getRandomValues(new Uint8Array(12));
    await store.set("userInfo", "keyName", "cryptoIv", iv);
    return await window.crypto.subtle.encrypt(
      {
        name: "AES-GCM",
        iv: iv,
      },
      key,
      data
    );
  }

  //Decrypt
  async decrypt(encryptedData) {
    const key = await this.getKey();
    const iv: Uint8Array = await store.get("userInfo", "cryptoIv");
    return window.crypto.subtle.decrypt(
      {
        name: "AES-GCM",
        iv: iv,
      },
      key,
      encryptedData
    );
  }
}

export const crypto = new CryptoKeys();
