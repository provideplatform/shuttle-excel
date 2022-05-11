/* eslint-disable @typescript-eslint/no-unused-vars */
import { onError } from "../common/common";
import { baseline } from "./index";
import { store } from "../settings/store";
import { ProvideClient } from "src/client/provide-client";
import * as $ from "jquery";
import { Mapping } from "@provide/types";

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension, Office */

export class ExcelAPI {
  tableToBaseline: string;

  async createTableListener(context: Excel.RequestContext): Promise<string> {
    try {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let range = sheet.getUsedRange();
      let table = range.getTables().getFirst();
      table.load("name");

      table.onChanged.add(this.onChange);
      await context.sync();

      return table.name;
    } catch {
      this.catchError;
    }
  }

  async createSheetListener(context: Excel.RequestContext): Promise<string> {
    try {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");

      sheet.onChanged.add(this.onWorksheetChange);
      await context.sync();

      return sheet.name;
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

  //Read all the changed data.
  onWorksheetChange(eventArgs: Excel.WorksheetChangedEventArgs): Promise<unknown> {
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

      var tableExists = await store.tableExists("tablePrimaryKeys", table.name);

      if (tableExists) {
        table.onChanged.add(this.onChange);
        await context.sync();
      }
    } catch {
      this.catchError;
    }
  }

  //Add listener to table
  async addSheetListener(context: Excel.RequestContext): Promise<void> {
    try {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");

      await context.sync();

      var sheetID = (await store.getTableID(sheet.name)).toString();
      var sheetExists = await store.tableExists("tablePrimaryKeys", sheetID);

      if (sheetExists) {
        sheet.onChanged.add(this.onWorksheetChange);
        await context.sync();
      }
    } catch {
      this.catchError;
    }
  }

  async createMappings(identClient: ProvideClient, mapping: Mapping): Promise<void> {
    var receivedSORModel = mapping.models[0];

    var mappingForm = $("#mapping-form form")
      .serializeArray()
      .reduce(function (obj, item) {
        obj[item.name] = item.value;
        return obj;
      }, {});

    var tableName = mappingForm["tableName"].toString();
    var primaryKey = mappingForm["primaryKey"].toString();

    //Create Store Name = refModelId, Stored Data = { primary key: baselineID}
    await store.close();
    await store.createTable(receivedSORModel.refModelId, "primaryKeyID", "baselineID", "baselineID");

    //Update Store Name = "TableNames", Stored Data = { tableName: refModelId}
    await store.setTableID(tableName, receivedSORModel.refModelId);

    //Update Store Name = "TablePrimaryKeys", Stored Data = { refModelId: primaryKeyColumn}
    await store.setPrimaryKeyColumnName(receivedSORModel.refModelId, primaryKey);

    var fields = [];

    //Add my fields
    fields = receivedSORModel.fields.map(async (field) => {
      var column = mappingForm[field.refFieldId];
      var columnToField = {
        name: column,
        type: field.type,
        refFieldId: field.refFieldId,
      };

      return columnToField;
    });

    var allFields = await Promise.all(fields);

    //Add my table
    var table = {
      type: tableName,
      primary_key: primaryKey,
      fields: allFields,
    };

    //Push my model to the exisiting mapping
    var models = [];
    models.push(receivedSORModel);
    models.push(table);

    var myMapping = mapping;
    myMapping.models = models;

    await identClient.updateWorkgroupMapping(mapping.id, myMapping);
  }

  //TODO
  async updateMappings(identClient: ProvideClient, mapping: Mapping): Promise<void> {
    // var excelModel = mapping.models[1];
    // var mappingForm = $("#mapping-form form")
    //   .serializeArray()
    //   .reduce(function (obj, item) {
    //     obj[item.name] = item.value;
    //     return obj;
    //   }, {});
    // var tableName = mappingForm["tableName"].toString();
    // var primaryKey = mappingForm["primaryKey"].toString();
    // var tableExists = await store.tableExists("tableNames", excelModel.type);
    // if (!tableExists) {
    //   await store.close();
    //   await store.createTable(tableName, "primaryKeyID", "baselineID", "baselineID");
    // }
    // await store.setTableName(excelModel.refModelId, tableName);
    // await store.setPrimaryKey(tableName, primaryKey);
    // var fields = [];
    // fields = excelModel.fields.map(async (field) => {
    //   var column = mappingForm[field.refFieldId];
    //   await store.setColumnMapping(tableName, column.toString(), field.refFieldId);
    //   var columnToField = {
    //     name: column,
    //     type: field.type,
    //     refFieldId: field.refFieldId,
    //   };
    //   return columnToField;
    // });
    // var allFields = await Promise.all(fields);
    // var table = {
    //   type: tableName,
    //   primary_key: primaryKey,
    //   fields: allFields,
    // };
    // var models = [];
    // models.push(excelModel);
    // models.push(table);
    // var myMapping = mapping;
    // myMapping.models = models;
    // await identClient.updateWorkgroupMapping(mapping.id, myMapping);
  }

  async changeButtonColor(): Promise<void> {
    var submitButton = document.getElementById("mapping-form-btn");
    submitButton.innerHTML = "Baselined";
    submitButton.style.backgroundColor = "Green";
  }

  async trim(str: string): Promise<string> {
    return str.replace(/\s/g, "");
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
