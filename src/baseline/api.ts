import { onError } from "../common/common";
import { baseline } from "./index";
import { showPrimaryKeyDialog } from "../dialogs/dialogs-helpers";
import { indexedDatabase } from "../settings/settings";

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension, Office */

export class ExcelAPI {
  tableToBaseline: string;

  async createTableBinding(context: Excel.RequestContext): Promise<string> {
    try {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let range = sheet.getUsedRange();
      let table = range.getTables().getFirst();
      table.load("name");

      table.onChanged.add(this.onChange);
      await context.sync();

      await indexedDatabase.createObjectStore(table.name);

      return table.name;
    } catch {
      this.catchError;
    }
  }

  //Read all the changed data.
  onChange(eventArgs: Excel.TableChangedEventArgs): Promise<unknown> {
    //Set global variables
    baseline.sendMessage(eventArgs);

    return new Promise((resolve, reject) => {
      if (eventArgs) {
        resolve(eventArgs);
      } else {
        return reject(Error);
      }
    });
  }

  //Add listener to table
  async addTableListener(context: Excel.RequestContext): Promise<void> {
    try {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let range = sheet.getUsedRange();
      let table = range.getTables().getFirst();
      table.load("name");

      await context.sync();

      var tableExists = await indexedDatabase.openDB(table.name);

     if (tableExists) {
        table.onChanged.add(this.onChange);
        await context.sync();
     }
    } catch {
      this.catchError;
    }
  }

  async getPrimaryKeyColumn(tableID: string): Promise<string> {
    let primaryKeyColumn: string;

    await showPrimaryKeyDialog().then(
      (primaryKeyInput) => {
        primaryKeyColumn = primaryKeyInput.primaryKey;
      },
      () => {
        /* NOTE: On cancel - do nothing */
      }
    );

    await indexedDatabase.setPrimaryKey(tableID, primaryKeyColumn);
    return primaryKeyColumn;
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

export const excelAPI = new ExcelAPI();
