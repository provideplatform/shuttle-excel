/* eslint-disable @typescript-eslint/no-unused-vars */
import { onError } from "../common/common";
import { store } from "../settings/store";
import { excelHandler } from "./excel-handler";
import { localStore } from "../settings/settings";
import { getMyWorkgroups } from "../taskpane/taskpane";

//TODO
import { ProtocolMessage } from "../models/protocolMessage";

// eslint-disable-next-line no-unused-vars
/* global Excel, Office, OfficeExtension */

export class InBound {
  async updateExcelTable(context: Excel.RequestContext, msg: ProtocolMessage): Promise<void> {
    //Disable event handler
    await this.disableTableListener(context);

    var tableName = await excelHandler.getSheetName(context);
    let tableID = (await store.getTableID(tableName)).toString();
    let primaryKeyColumnName = await store.getPrimaryKeyColumnName(tableID);

    var dataColumn = Object.keys(msg.payload.data)[0];
    var address;
    var id;

    //var idExists = await store.keyExists(tableID, msg.baselineID, "baselineID");

    var primaryKeyColumnAddress = await excelHandler.getSheetColumnAddress(context, primaryKeyColumnName);
    var idExists = await excelHandler.cellDataExits(context, msg.id, primaryKeyColumnAddress);

    if (!idExists) {
      id = await this.generateNewPrimaryKeyID(context, primaryKeyColumnAddress);

      // //map it with baseline ID given in the message
      // //Get the workflowID from localStorage (Do While)
      // await getMyWorkgroups(true);
      // var workflowID = null;
      // do {
      //   workflowID = await localStore.getWorkflowID();
      // } while (workflowID == null);

      // //Set baselineID in table
      // await store.setBaselineIDAndWorkflowID(tableID, id, [msg.baselineID, workflowID]);

      //Add record in table
      await this.addNewIDToTable(context, id, primaryKeyColumnAddress);

      address = await excelHandler.getDataCellAddress(context, id, dataColumn, primaryKeyColumnAddress);
    } else {
      // id = await store.getPrimaryKeyId(tableID, msg.baselineID);
      id = msg.id;
      address = await excelHandler.getDataCellAddress(context, id, dataColumn, primaryKeyColumnAddress);
    }

    //Update Data
    var range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
    range.values = [[msg.payload.data[dataColumn]]];
    range.format.autofitColumns();
    range.format.fill.color = "yellow";
    range.format.font.bold = true;
    await context.sync();

    //Enable event handler
    await this.enableTableListener(context);
  }

  private async generateNewPrimaryKeyID(
    context: Excel.RequestContext,
    primaryKeyColumnAddress: string
  ): Promise<string> {
    //Get value in last cell
    let primaryKeyRange = context.workbook.worksheets
      .getActiveWorksheet()
      .getRange(primaryKeyColumnAddress + ":" + primaryKeyColumnAddress)
      .getUsedRange();
    let lastCell = primaryKeyRange.getLastCell();
    lastCell.load("values");

    await context.sync();

    //Increment that value
    var id = parseInt(lastCell.values[0][0]);
    var newID = id + 1;

    return newID.toString();
  }

  private async addNewIDToTable(context: Excel.RequestContext, newID: string, primaryKeyColumn: string): Promise<void> {
    let originalRange = context.workbook.worksheets
      .getActiveWorksheet()
      .getRange(primaryKeyColumn + ":" + primaryKeyColumn)
      .getUsedRange();

    let expandedRange = originalRange.getResizedRange(1, 0);
    expandedRange.copyFrom(originalRange.getLastCell(), Excel.RangeCopyType.formats);
    let lastCell = expandedRange.getLastCell();
    lastCell.values = [[newID]];

    lastCell.format.autofitRows();
    return await context.sync();
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
