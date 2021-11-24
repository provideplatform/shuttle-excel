import { onError } from "../common/common";
import { Object, BaselineResponse } from "@provide/types";
import { ProvideClient } from "src/client/provide-client";
import { indexedDatabase } from "../settings/settings";
import { excelHandler } from "./excel-handler";

// eslint-disable-next-line no-unused-vars
/* global Excel, Office, OfficeExtension */

export class OutBound {
  async send(
    context: Excel.RequestContext,
    changedData: Excel.TableChangedEventArgs,
    identClient: ProvideClient
  ): Promise<void> {
    try {
      let tableName = await excelHandler.getTableName(context);
      let message = await this.createMessage(context, tableName, changedData);

      //TO SECURE --> JsonSanitizer.sanitize(JSON.stringify(message))
      console.log(JSON.stringify(message));
      let baselineResponse: BaselineResponse;

      let recordExists = await indexedDatabase.keyExists(tableName, [message.payload.id, message.type], "Out");

      console.log(recordExists);

      if (!recordExists) {
        baselineResponse = await identClient.sendCreateProtocolMessage(message);
        console.log(baselineResponse);
        await indexedDatabase.set(tableName, [message.payload.id, message.type], baselineResponse.baselineId);
      } else {
        let baselineID = await indexedDatabase.get(tableName, [message.payload.id, message.type]);
        console.log("Baseline ID: " + baselineID);
        baselineResponse = await identClient.sendUpdateProtocolMessage(baselineID, message);
        console.log("Baseline message : " + baselineResponse);
      }
    } catch {
      this.catchError;
    }
  }

  private async createMessage(
    context: Excel.RequestContext,
    tableName: string,
    changedData: Excel.TableChangedEventArgs
  ): Promise<Object> {
    let primaryKey = await indexedDatabase.getPrimaryKeyField(tableName);
    console.log(primaryKey);
    let id = await this.getPrimaryKeyID(context, changedData, primaryKey);
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
    changedData: Excel.TableChangedEventArgs,
    primaryKeyColumnName: String
  ): Promise<string> {
    try {
      let primaryKeyColumnAddress = await excelHandler.getColumnAddress(context, primaryKeyColumnName);
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
