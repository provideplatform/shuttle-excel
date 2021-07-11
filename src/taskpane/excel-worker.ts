// NOTE: Logic of working with Excel

import { Application } from "@provide/types";
import { onError } from "../common/common";
import { baseline } from "../baseline/index";
import { ProvideClient } from "src/client/provide-client";

/* global Excel, OfficeExtension */

export class ExcelWorker {
  showWorkgroups(sheetName: string, applications: Application[], active: boolean = false): Promise<unknown> {
    return Excel.run((context: Excel.RequestContext) => {
      var sheets = context.workbook.worksheets;
      sheets.load("items/name");

      return context.sync().then(() => {
        let sheet: Excel.Worksheet = sheets.items.find((x) => x.name === sheetName);
        if (!sheet) {
          sheet = sheets.add(sheetName);
        }

        //this.renderWorkgroups(sheet, applications);

        if (active) {
          //sheet.activate();
        }

        return context.sync();
      });
    }).catch(this.catchError);
  }

  private renderWorkgroups(sheet: Excel.Worksheet, applications: Application[]): void {
    let workgroupsTable: Excel.Table = sheet.tables.getItemOrNullObject("Workgroups");
    if (workgroupsTable) {
      workgroupsTable.delete();
    }

    workgroupsTable = sheet.tables.add("A1:F1", true /* hasHeaders */);
    workgroupsTable.name = "Workgroups";

    workgroupsTable.getHeaderRowRange().values = [["NetworkId", "UserId", "Name", "Description", "Type", "Hidden"]];
    const asRows = applications.map((app) => {
      return [app.networkId, app.userId, app.name, app.description, app.type, app.hidden];
    });
    workgroupsTable.rows.add(null /* add at the end */, asRows);

    // expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
    workgroupsTable.getRange().format.autofitColumns();
    workgroupsTable.getRange().format.autofitRows();
  }

  createInitialSetup(): Promise<unknown> {
    return baseline.createTableListeners();
  }

  startBaselineService(identClient: ProvideClient): Promise<void> {
    return baseline.startToSendAndReceiveProtocolMessage(identClient);
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

export const excelWorker = new ExcelWorker();

// function test() {
//   Excel.run((context) => {
//     const cursheet = context.workbook.worksheets.getActiveWorksheet();
//     const cellA1_A2 = cursheet.getRange("A1:A3");

//     // const value = new Date(); // identClient.test_ExpiresAt();
//     const value = identClient?.test_expiresAt;
//     cellA1_A2.values = [[ value ], [ new Date() ], [ identClient?.isExpired ]];
//     cellA1_A2.format.autofitColumns();

//     return context.sync();
//   })
//   .catch(function(error) {
//     console.log("Error: " + error);
//     if (error instanceof OfficeExtension.Error) {
//       console.log("Debug info: " + JSON.stringify(error.debugInfo));
//       onError(error.message);
//     } else {
//       onError(error);
//     }
//   })
// }
