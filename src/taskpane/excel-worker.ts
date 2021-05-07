import { Application } from "@provide/types";
import { onError } from "./common";

/* global Excel, OfficeExtension */

export class ExcelWorker {
  async showWorkgroups(applications: Application[]): Promise<unknown> {
    try {
      Excel.run((context: Excel.RequestContext) => {
        this.renderWorkgroups(context, applications);
        return context.sync();
      });
    } catch (error) {
      return this.catchError(error);
    }
  }

  private renderWorkgroups(context: Excel.RequestContext, applications: Application[]): void {
    const currentWorksheet: Excel.Worksheet = context.workbook.worksheets.getActiveWorksheet();

    let workgroupsTable: Excel.Table = currentWorksheet.tables.getItemOrNullObject("Workgroups");
    if (workgroupsTable) {
      workgroupsTable.delete();
    }

    workgroupsTable = currentWorksheet.tables.add("A1:F1", true /* hasHeaders */);
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