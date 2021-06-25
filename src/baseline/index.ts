import { excelAPI } from "./api";
import { outboundMessage } from "./outbound";
import { inboundMessage } from "./inbound";
import { onError } from "../common/common";
import { ProvideClient } from "src/client/provide-client";
import { NatsClientFacade as NatsClient } from "../client/nats-listener";

// eslint-disable-next-line no-unused-vars
/* global Excel, Office, OfficeExtension */

export class Baseline { 
  _primaryKeyColumn: String;
  _identClient: ProvideClient;
  _natsClient: NatsClient; 

  //Initialize baseline
  async createTableListeners(): Promise<unknown> {
    return Excel.run(async (context: Excel.RequestContext) => {
      await excelAPI.createExcelBinding(context);
      return context.sync().then(async () => {
        this._primaryKeyColumn = await excelAPI.getPrimaryKeyColumn();
      });
    }).catch(this.catchError);
  }

  //Start the Baseline Service after login
  async startToSendAndReceiveProtocolMessage(identClient: ProvideClient): Promise<void> {
    try {

      //Set Provide client for sending messages
      this._identClient = identClient;

      //Connect to Nats for receiving messages
     if (!this._natsClient) {
       await this._identClient.connectNatsClient();
       this._natsClient = this._identClient.natsClient;
     }
    
     //Subscribe
      this.receiveMessage();
      

      return;
    } catch {
      this.catchError;
    }
  } 

  sendMessage(changedData: Excel.TableChangedEventArgs): void {
   
    Excel.run((context: Excel.RequestContext) => {
      outboundMessage.send(context, changedData, this._identClient, this._primaryKeyColumn);
      return context.sync(); 
    }).catch(this.catchError); 
    
  }

  receiveMessage(): void {
   try {
     inboundMessage.primaryKeyColumn = this._primaryKeyColumn;
     this._natsClient.subscribe(">", inboundMessage.handler);
   } catch {
     this.catchError;
   }
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
