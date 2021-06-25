import { onError } from "../common/common";
import { indexedDatabase } from "../settings/settings";
import { Record } from "../models/record";

//TODO
import { ProtocolMessage } from "../models/protocolMessage";

// eslint-disable-next-line no-unused-vars
/* global Excel, Office, OfficeExtension */

export class InBound {
  primaryKeyColumn: String;

  handler(msg: ProtocolMessage) {
    Excel.run((context: Excel.RequestContext) => {
      this.updateExcelTable(context, msg);
      return context.sync();
    }).catch(this.catchError);
  }

  //TODO
  private async updateExcelTable(context: Excel.RequestContext, msg: ProtocolMessage): Promise<void> {
    var record = await this.getPrimaryKeyRecord(context, msg);
    //OPTION 2: Refer to the primary key and column Name already given in the protocol message

    //DO a Lookup funtion with primary key and column name --> reverse of outbound
    var address = await this.getDataCellAddress(context, record.primaryKey, record.columnName, this.primaryKeyColumn);

    var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
    range.values = [[msg.data]];
    range.format.autofitColumns();
    return context.sync();
  }

  private async getPrimaryKeyRecord(context: Excel.RequestContext, msg: ProtocolMessage): Promise<Record> {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getUsedRange();
    let table = range.getTables().getFirst();
    table.load("id");

    await context.sync();

    //OPTION 1: Use the baselineID received in protocol message and retrieve corresponding column Names and primary keys from IDB
    var record: Record = await indexedDatabase.getKey(table.id, msg.baselineID);
    return record;
  }

  private async getDataCellAddress(context: Excel.RequestContext, primaryKey: string, columnName: string, primaryKeyColumn: String): Promise<string> {
    //Get column Header Cell
    let table = context.workbook.worksheets.getActiveWorksheet().getUsedRange().getTables().getFirst();
    let headerRange = table.getHeaderRowRange();
    let headerCell = headerRange.findOrNullObject(columnName, { completeMatch: true });
    headerCell.load("address");

    //Get Primary Key Cell
    let primaryKeyRange = context.workbook.worksheets
      .getActiveWorksheet()
      .getRange(primaryKeyColumn + ":" + primaryKeyColumn);
    let primaryKeyCell = primaryKeyRange.findOrNullObject(primaryKey, { completeMatch: true });
    primaryKeyCell.load("address");

    return context.sync().then(() => {
      var column = headerCell.address.split(/\d+/)[0];
      var row = primaryKeyCell.address.split(/\D+/)[0];
      return column + row;
    });
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

export const inboundMessage = new InBound();
