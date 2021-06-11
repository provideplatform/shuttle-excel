import { onError } from "../common/common";
import { baseline } from "./index";
import { showPrimaryKeyDialog } from "../dialogs/dialogs-helpers";

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension, Office */

export class ExcelAPI {
  async createExcelBinding(context: Excel.RequestContext): Promise<void> {
    try {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let range = sheet.getUsedRange();
      let table = range.getTables().getFirst();

      table.onChanged.add(this.onChange);
      await context.sync();
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

  async getPrimaryKeyColumn(): Promise<String> {
    let primaryKeyColumn: String;

    await showPrimaryKeyDialog().then(
      (primaryKeyInput) => {
        primaryKeyColumn = primaryKeyInput.primaryKey;
      },
      () => {
        /* NOTE: On cancel - do nothing */
      }
    );
    return primaryKeyColumn;

    //TODO: Create a range binding
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
