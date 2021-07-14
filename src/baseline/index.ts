import { excelAPI } from "./api";
import { outboundMessage } from "./outbound";
import { inboundMessage } from "./inbound";
import { onError } from "../common/common";
import { ProvideClient } from "src/client/provide-client";
import { NatsClientFacade as NatsClient } from "../client/nats-listener";

// eslint-disable-next-line no-unused-vars
/* global Excel, Office, OfficeExtension */

export class Baseline { 
  
  _identClient: ProvideClient;
  _natsClient: NatsClient;
  

  //Initialize baseline
  async createTableListeners(): Promise<unknown> {
    return Excel.run(async (context: Excel.RequestContext) => {
      var tableName = await excelAPI.createTableBinding(context);
      return context.sync().then(async () => {
        await excelAPI.getPrimaryKeyColumn(tableName);
      });
    }).catch(this.catchError);
  }

  //Start the Baseline Service after login
  async startToSendAndReceiveProtocolMessage(identClient: ProvideClient): Promise<void> {
    try {
      //Activate listeners on table
      await this.activateTableListeners();

      //Set Provide client for sending messages
      this._identClient = identClient;

      var data = {};
      data["Region"] = "Central";

      const test = {
        baselineID: "89455b85-6fec-4aec-a40d-551df0b13ca8",
        id: "574733300",
        type: "general_consistency",
        payload: {
          id: "574733300",
          data: data,
        },
      };
      inboundMessage.handler(test);

      //Connect to Nats for receiving messages
      //TODO : Check NATS client
      if (!this._natsClient) {
        //PROBLEM
        console.log("Connecting to Nats");
        await this._identClient.connectNatsClient();
        console.log("Connected to Nats");
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
      outboundMessage.send(context, changedData, this._identClient);
      return context.sync(); 
    }).catch(this.catchError); 
    
  }

  receiveMessage(): void {
   try {
     this._natsClient.subscribe("baseline-proxy.inbound", inboundMessage.handler);
           
   } catch {
     this.catchError;
   }
  }

 private async activateTableListeners() : Promise<unknown> {
    return Excel.run(async (context: Excel.RequestContext) => {
      await excelAPI.addTableListener(context);
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
