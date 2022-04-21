/* eslint-disable @typescript-eslint/no-unused-vars */
import { onError } from "../common/common";
import { Object, BaselineResponse } from "@provide/types";
import { ProvideClient } from "src/client/provide-client";
import { excelHandler } from "./excel-handler";
import { store } from "../settings/store";

// eslint-disable-next-line no-unused-vars
/* global Excel, Office, OfficeExtension */

export class OutBound {
  async send(
    context: Excel.RequestContext,
    changedData: Excel.TableChangedEventArgs | Excel.WorksheetChangedEventArgs,
    identClient: ProvideClient
  ): Promise<void> {
    try {
      //Get the primary key ID
      let tableName = await excelHandler.getSheetName(context);
      let tableID = (await store.getTableID(tableName)).toString();
      let primaryKeyColumnName = await store.getPrimaryKeyColumnName(tableID);
      let primaryKeyID = await this.getPrimaryKeyID(context, changedData, primaryKeyColumnName);

      //Check if baselineID exists
      let baselineIDExists = await store.keyExists(tableID, primaryKeyID);

      //Create the message
      let message = await this.createMessage(context, primaryKeyID, changedData);

      console.log(JSON.stringify(message));
      let baselineResponse: BaselineResponse;

      if (!baselineIDExists) {
        baselineResponse = await identClient.sendCreateProtocolMessage(message);
        console.log(baselineResponse);
        await store.setBaselineID(tableID, primaryKeyID, baselineResponse.baselineId);
      } else {
        let baselineId = await store.getBaselineId(tableID, primaryKeyID);
        baselineResponse = await identClient.sendUpdateProtocolMessage(baselineId, message);
        console.log("Baseline message : " + baselineResponse);
      }
    } catch {
      this.catchError;
    }
  }

  private async createMessage(
    context: Excel.RequestContext,
    primaryKeyID: string,
    changedData: Excel.TableChangedEventArgs | Excel.WorksheetChangedEventArgs
  ): Promise<Object> {
    //Get the refModelId (tableName --> refModelId)
    //Get the primary key column (refModelID --> primaryKeyColumn)
    //Get the primary key id
    //TODO: Check if the primary key is mapped to any baselineID
    //If yes, then resolve workstep ---> updateObject
    //If no, ask user to create

    let id = primaryKeyID;
    let dataColumnHeader = await excelHandler.getDataColumnHeader(context, changedData);

    let message: Object = {} as Object;
    message.type = "general_consistency";
    message.id = id;

    const data = {};
    data[dataColumnHeader] = changedData.details.valueAfter;

    let _payload = {
      id: id,
      data: data,
    };

    message.payload = _payload;

    return message;
  }

  private async getPrimaryKeyID(
    context: Excel.RequestContext,
    changedData: Excel.TableChangedEventArgs | Excel.WorksheetChangedEventArgs,
    primaryKeyColumnName: String
  ): Promise<string> {
    try {
      let primaryKeyColumnAddress = await excelHandler.getSheetColumnAddress(context, primaryKeyColumnName);
      let primaryKeyCell = primaryKeyColumnAddress + changedData.address.split(/\D+/)[1];
      let primaryKeyID = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange(primaryKeyCell + ":" + primaryKeyCell);

      primaryKeyID.load("values");
      await context.sync();

      console.log(primaryKeyID.values);
      return primaryKeyID.values[0][0].toString();
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

export const outboundMessage = new OutBound();
