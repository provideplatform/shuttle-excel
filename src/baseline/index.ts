import { excelAPI } from "./api";
import { outboundMessage } from "./outbound";
import { onError } from "../common/common";
import { ProvideClient } from "src/client/provide-client";

// eslint-disable-next-line no-unused-vars
/* global Excel, Office, OfficeExtension */

export class Baseline {
  //Create a binding popup
  _websocket: WebSocket;
  _primaryKey: String;
  _identClient: ProvideClient;

  setBaselineServiceClient(identClient: ProvideClient): Promise<void> {
    try {
      this._identClient = identClient;
      return;
    } catch {
      this.catchError;
    }
  }
  
  async createListeners(): Promise<unknown> {
    return Excel.run((context: Excel.RequestContext) => {
      excelAPI.createExcelBinding(context);
      return context.sync().then(async () => {
        this._primaryKey = await excelAPI.getPrimaryKeyColumn();
      });
    }).catch(this.catchError);
  }

  private createWebSocket(): Promise<WebSocket> {
    return new Promise((resolve, reject) => {
      let webSocket = new WebSocket("wss://localhost:8080");

      if (webSocket) {
        resolve(webSocket);
      } else {
        reject(Error);
      }
    });
  }

  sendMessage(changedData: Excel.TableChangedEventArgs): void {
   
    Excel.run((context: Excel.RequestContext) => {
      outboundMessage.send(context, changedData, this._identClient, this._primaryKey);
      return context.sync(); 
    }).catch(this.catchError);
    
    
    
  }

  private catchError(error: any): void {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
      onError(error.message);
    } else {
      onError(error);
    }
  }
}

export const baseline = new Baseline();
