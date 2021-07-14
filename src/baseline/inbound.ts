import { onError } from "../common/common";
import { indexedDatabase } from "../settings/settings";
import { Record } from "../models/record";

//TODO
import { ProtocolMessage } from "../models/protocolMessage";

// eslint-disable-next-line no-unused-vars
/* global Excel, Office, OfficeExtension */

export class InBound {
  tableID: string;

  handler(msg: ProtocolMessage) {
    Excel.run((context: Excel.RequestContext) => {
      this.updateExcelTable(context, msg);
      return context.sync();
    }).catch(this.catchError);
  }

  private async updateExcelTable(context: Excel.RequestContext, msg: ProtocolMessage): Promise<void> {
 
    //Disable event handler
    await this.disableTableListener(context);

    var tableName = await this.getTableName(context); 
    var primaryKeyColumn = await indexedDatabase.getPrimaryKeyField(tableName);
    var dataColumn = Object.keys(msg.payload.data)[0];
    var address = await this.getDataCellAddress(context, msg.id, dataColumn , primaryKeyColumn);

    var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
    
    range.values = [[msg.payload.data[dataColumn]]];
    range.format.autofitColumns();
    range.format.fill.color = "yellow";
    range.format.font.bold = true;
    await context.sync();

    //Enable event handler
    await this.enableTableListener(context);
  }

  private async getPrimaryKeyRecord(context: Excel.RequestContext, msg: ProtocolMessage): Promise<Record> {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getUsedRange();
    let table = range.getTables().getFirst();
    table.load("id");

    await context.sync();

    this.tableID = table.id;
    var record: Record = await indexedDatabase.getKey(table.id, msg.baselineID);
    return record;
  }

  private async getDataCellAddress(context: Excel.RequestContext, primaryKeyValue: string, columnName: string, primaryKeyColumn: string): Promise<string> {
    //Get column Header Cell
    let table = context.workbook.worksheets.getActiveWorksheet().getUsedRange().getTables().getFirst();
    let headerRange = table.getHeaderRowRange();
    let headerCell = headerRange.findOrNullObject(columnName, { completeMatch: true });
    headerCell.load("address");

    //Get Primary Key Cell
    let primaryKeyRange = context.workbook.worksheets
      .getActiveWorksheet()
      .getRange(primaryKeyColumn + ":" + primaryKeyColumn);
    let primaryKeyCell = primaryKeyRange.findOrNullObject(primaryKeyValue, { completeMatch: true });
    primaryKeyCell.load("address");

    return context.sync().then(() => {
      var address = headerCell.address.split("!")[1];
      var column = address.split(/\d+/)[0];
      var row = primaryKeyCell.address.split(/\D+/)[1];
      return column + row;
    });
  }

  private async getTableName(context: Excel.RequestContext): Promise<string> {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getUsedRange();
    let table = range.getTables().getFirst();
    table.load("name");

    await context.sync();

    return table.name; 
  }

  private async disableTableListener(context: Excel.RequestContext): Promise<void> {
    
   /* context.runtime.load("enableEvents");
    await context.sync();

    console.log("before disable" + context.runtime.enableEvents.toString());*/
    context.runtime.enableEvents = false;
    await context.sync();

   /* context.runtime.load("enableEvents");
    await context.sync();

    console.log("after disable" + context.runtime.enableEvents.toString());*/
  }

  private async enableTableListener(context: Excel.RequestContext): Promise<void> {
    /*context.runtime.load("enableEvents");
    await context.sync();

    console.log("before enable" + context.runtime.enableEvents.toString());*/
    context.runtime.enableEvents = true;
    await context.sync();

    /*context.runtime.load("enableEvents");
    await context.sync();

    console.log("after enable" + context.runtime.enableEvents);*/
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
