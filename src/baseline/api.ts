import { onError } from "../common/common";
import { baseline } from "./index";
import { showPrimaryKeyDialog } from "../dialogs/dialogs-helpers";
import { indexedDatabase } from "../settings/settings";

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension, Office */

export class ExcelAPI {
  tableToBaseline: string ;

  async createTableBinding(context: Excel.RequestContext): Promise<string> {
    try {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let range = sheet.getUsedRange();
      let table = range.getTables().getFirst();
      table.load("name");

      table.onChanged.add(this.onChange);
      await context.sync();

      table.name = table.name + "B";
      await context.sync;

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
  //TODO: MAKE THIS CONDITIONAL ONLY IF THE BASELINING HAS BEEN ADDED FOR THIS TABLE
  async addTableListener(context: Excel.RequestContext) : Promise<void> {
    try {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let range = sheet.getUsedRange();
      let table = range.getTables().getFirst();
      table.load('name');

      await context.sync();

      await indexedDatabase.openDB();
     
     if(this.tableExists(table.name)){
        table.onChanged.add(this.onChange);
        await context.sync();
        console.log("Listeners activated");
     }
       
      
    } catch {
      this.catchError;
    } 
  }

 tableExists(tableName) : boolean {

  if(tableName.split("").slice(-1)[0] == "B"){
    return true;
  }

  return false;

 }





  

  async getPrimaryKeyColumn(tableID : string): Promise<string> {
    let primaryKeyColumn: string;

    await showPrimaryKeyDialog().then(
      (primaryKeyInput) => {
        primaryKeyColumn = primaryKeyInput.primaryKey;
      },
      () => {
        /* NOTE: On cancel - do nothing */
      }
    );

    await indexedDatabase.setPrimaryKey(tableID, primaryKeyColumn)
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
