import { onError } from "../common/common";
import { BusinessObject, BaselineResponse } from "@provide/types";
import { ProvideClient } from "src/client/provide-client";
import { indexedDatabase } from "../settings/settings";

// eslint-disable-next-line no-unused-vars
/* global Excel, Office, OfficeExtension */

export class OutBound {
  async send(context: Excel.RequestContext, changedData: Excel.TableChangedEventArgs, identClient: ProvideClient, primaryKeyColumn: String): Promise<void> {
    try {
      let message = await this.createMessage(context, changedData, primaryKeyColumn);
      console.log(message);

      let baselineResponse: BaselineResponse;

      let recordCount: number = await indexedDatabase.recordCount(changedData.tableId, [message.payload.id, message.type]);

      if (recordCount < 1) {
        baselineResponse = await identClient.sendCreateProtocolMessage(message);
        console.log(baselineResponse);
        await indexedDatabase.set(changedData.tableId, [message.payload.id, message.type], baselineResponse.baselineId);
      } else {
        let baselineID = await indexedDatabase.get(changedData.tableId, [message.payload.id, message.type]);
        baselineResponse = await identClient.sendUpdateProtocolMessage(baselineID, message);
        console.log(baselineResponse);
      }
    } catch {
      this.catchError;
    }
  }

  private async createMessage(context: Excel.RequestContext, changedData: Excel.TableChangedEventArgs, primaryKeyColumn: String): Promise<BusinessObject> {
    let primaryKey = await this.getPrimaryKey(context, changedData, primaryKeyColumn);
    let dataColumnHeader = await this.getDataColumnHeader(context, changedData);

    let message: BusinessObject = {} as BusinessObject;
    message.type = dataColumnHeader;

    let _payload = {
      id: primaryKey.toString(),
      data: changedData.details.valueAfter,
    };

    message.payload = _payload;

    return message;
  }

  private getPrimaryKey(context: Excel.RequestContext, changedData: Excel.TableChangedEventArgs, primaryKeyColumn: String): Promise<string | number | boolean> {
    try {
      let dataCell = context.workbook.worksheets.getActiveWorksheet().getRange(changedData.address);
      let dataColumn = changedData.address.split(/\d+/)[0];
      let dataRange = context.workbook.worksheets.getActiveWorksheet().getRange(dataColumn + ":" + dataColumn);
      let primaryKeyRange = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange(primaryKeyColumn + ":" + primaryKeyColumn);

      let primaryKeyID = context.workbook.functions.lookup(dataCell, dataRange, primaryKeyRange);

      primaryKeyID.load("value");

      return context.sync().then(() => {
        return primaryKeyID.value;
      });
    } catch {
      this.catchError;
    }
  }

 private getDataColumnHeader(context: Excel.RequestContext, changedData: Excel.TableChangedEventArgs): Promise<string> {
    try {
      let dataColumn = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange(changedData.address.split(/\d+/)[0] + "1");
      dataColumn.load("values");
      return context.sync().then(() => {
        return dataColumn.values[0][0];
      });
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
