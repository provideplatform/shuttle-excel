// NOTE: Logic of working with Excel

import { Mapping } from "@provide/types";
import { store } from "../settings/store";
import { onError } from "../common/common";

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension */

export class MappingForm {
  tableName;
  primaryKey;
  columnNames: String[];
  tableExists: boolean;
  workgroupId: string;
  sheetName;
  sheetColumnNames: String[];
  sheetColumnDataType: String[];

  async showWorkgroupMappings(mappings: Mapping[]): Promise<void> {
    var excelSheetName;
    var excelSheetColumnNames;

    await Excel.run(async (context: Excel.RequestContext) => {
      excelSheetName = await this.getSheetName(context);
      excelSheetColumnNames = await this.getSheetColumnNames(context);
    }).catch(this.catchError);

    var mappingForm = document.getElementById("mapping-form-options");
    mappings.map((mapping) => {
      mapping.models.map(async (model) => {
        this.tableName = model.type;
        this.primaryKey = model.primaryKey;
        this.tableExists = await store.tableExists("tableNames", this.tableName);

        var tableID = await this.trim(model.type);
        var primaryKeyID = await this.trim(model.primaryKey);

        document.getElementById("mapping-form-header").innerHTML = "Confirm Mappings";
        var pkOptions = await this.addOptions(excelSheetColumnNames, model.primaryKey);
        //TO SECURE --> innerHTML https://newbedev.com/xss-prevention-and-innerhtml
        //TO SECURE --> Input validation
        mappingForm.innerHTML =
          `<div class="form-group">
						<label class="font-weight-normal h6" for="` +
          tableID +
          `">Table Name: ` +
          model.type +
          `</label>
						<input id="` +
          tableID +
          `" type="text" value="` +
          excelSheetName +
          `" class="form-control bg-transparent text-light shadow-none"\\>
						</div>
						<div class="form-group">	
						<label class="font-weight-normal h6" for="` +
          primaryKeyID +
          `">Primary Key Column: ` +
          model.primaryKey +
          `</label>
						<select id="` +
          primaryKeyID +
          `" class="form-control bg-transparent text-light shadow-none">` +
          pkOptions +
          `</select>
						</div>`;

        this.sheetColumnNames = [];

        model.fields.map(async (field) => {
          //field.name
          //field.type

          var columnID = await this.trim(field.name);
          this.sheetColumnNames.push(field.name);
          var options = await this.addOptions(excelSheetColumnNames, field.name);
          mappingForm.innerHTML +=
            `<div class="form-group container">
						<div class="row">
						<label class="col font-weight-normal h6" for="` +
            columnID +
            `">` +
            field.name +
            `<div class="text-muted font-weight-light">(` +
            field.type +
            `)</div></label>
						<select id="` +
            columnID +
            `" class="col form-control bg-transparent text-light shadow-none">` +
            options +
            `</select>
						</div>
						</div>`;
        });

        var submitButton = document.getElementById("mapping-form-btn");
        if (this.tableExists) {
          submitButton.innerHTML = "Baselined";
          submitButton.style.backgroundColor = "Green";
        } else {
          submitButton.innerHTML = "Start Baselining";
          submitButton.style.backgroundColor = "Red";
        }
      });
    });
  }

  async showUnmappedColumns(appId: string): Promise<void> {
    this.workgroupId = appId;
    this.sheetColumnDataType = [];

    await Excel.run(async (context: Excel.RequestContext) => {
      this.sheetName = await this.getSheetName(context);
      this.sheetColumnNames = await this.getSheetColumnNames(context);
      this.sheetColumnDataType = await this.getSheetColumnDataType(context, this.sheetColumnNames);
    }).catch(this.catchError);

    var mappingForm = document.getElementById("mapping-form-options");
    document.getElementById("mapping-form-header").innerHTML = "Create New Mapping";

    var pkOptions = await this.addOptions(this.sheetColumnNames);
    //TO SECURE --> innerHTML https://newbedev.com/xss-prevention-and-innerhtml
    //TO SECURE --> Input validation
    mappingForm.innerHTML =
      `<div class="form-group container">
				<div class="row">
				<label class="col" for="table-name"> Table Name: </label>
				<input id="table-name" type="text" value ="` +
      this.sheetName +
      `" class="col form-control bg-transparent text-light shadow-none" \\>
				</div>
				</div>
				<div class="form-group container">
				<div class="row">
				<label class="col" for="primary-key"> Primary Key Column: </label>
				<select id="primary-key" class="col form-control bg-transparent text-light shadow-none">` +
      pkOptions +
      `</select>
				</div>
				</div>`;

    this.sheetColumnNames.map(async (column, index) => {
      var columnID = await this.trim(column.toString());
      var columnDataType = await this.addOptions([this.sheetColumnDataType[index]]);

      //TO SECURE --> innerHTML https://newbedev.com/xss-prevention-and-innerhtml
      //TO SECURE --> Input validation
      mappingForm.innerHTML +=
        `<div class="form-group container float-right">
					<div class="row">
					<label class="col d-flex justify-content-end" for="` +
        columnID +
        `">` +
        column +
        `</label>
					<select id="` +
        columnID +
        `" class="col form-control bg-transparent text-light shadow-none">` +
        columnDataType +
        `</select>
					<!--<input id="` +
        column +
        `" type="checkbox" class="col form-control bg-transparent text-light shadow-none" style="margin-left:10px" \\>-->
					</div>
					</div>`;
    });

    var submitButton = document.getElementById("mapping-form-btn");
    submitButton.innerHTML = "Create Mapping";
  }

  private async getTableName(context: Excel.RequestContext): Promise<String> {
    var table = context.workbook.worksheets.getActiveWorksheet().getUsedRange().getTables().getFirst();
    table.load("name");
    await context.sync();
    return table.name;
  }

  private async getSheetName(context: Excel.RequestContext): Promise<String> {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("name");
    await context.sync();
    return sheet.name;
  }

  private async getColumnNames(context: Excel.RequestContext): Promise<String[]> {
    var table = context.workbook.worksheets.getActiveWorksheet().getUsedRange().getTables().getFirst();
    var columns = table.getHeaderRowRange();
    columns.load("values");
    await context.sync();
    return columns.values[0];
  }

  private async getSheetColumnNames(context: Excel.RequestContext): Promise<String[]> {
    var sheet = context.workbook.worksheets.getActiveWorksheet().getUsedRange();
    var columns = sheet.getRow(0);
    columns.load("values");
    await context.sync();
    return columns.values[0];
  }

  private async getSheetColumnDataType(context: Excel.RequestContext, columnNames: String[]): Promise<String[]> {
    var columnAddresses = await this.getSheetColumnAddress(context, columnNames);
    var columnDataType = [];
    columnAddresses.forEach((col) => {
      var colRange = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange(col + ":" + col)
        .getUsedRange();
      var lastCell = colRange.getLastCell();
      columnDataType.push(lastCell);
      lastCell.load("valueTypes");
    });

    await context.sync();

    columnDataType.forEach((col, i) => {
      columnDataType[i] = col.valueTypes[0][0].toString();
    });

    return columnDataType;
  }

  private async getSheetColumnAddress(context: Excel.RequestContext, columns: String[]): Promise<string[]> {
    var columnAddresses = [];

    var sheet = context.workbook.worksheets.getActiveWorksheet().getUsedRange();
    var headerRange = sheet.getRow(0);

    columns.forEach((column) => {
      let headerCell = headerRange.findOrNullObject(column.toString(), { completeMatch: true });
      columnAddresses.push(headerCell);
      headerCell.load("address");
    });

    await context.sync();

    columnAddresses.forEach((headerCell, i) => {
      var headerCellAddress = headerCell.address.split("!")[1];
      var columnAddress = headerCellAddress.split(/\d+/)[0];
      columnAddresses[i] = columnAddress;
    });

    return columnAddresses;
  }

  private async addOptions(options: String[], currentColumn?: string): Promise<String> {
    var str;
    if (this.tableExists) {
      options.map(async (column) => {
        //GET SAVED COLUMN NAMES
        var excelColumn = await store.getColumnMapping(this.tableName, currentColumn);
        //Add selected
        if (excelColumn == column) {
          str += `<option selected>` + column + "</option>";
        }
        str += `<option selected>` + column + "</option>";
      });
    }
    options.map(async (option) => {
      str += `<option selected>` + option + "</option>";
    });
    return str;
  }

  getFormTableName(): String {
    return this.tableName;
  }

  getFormSheetName(): String {
    return this.sheetName;
  }

  getFormPrimaryKey(): String {
    return this.primaryKey;
  }

  getFormColumnNames(): String[] {
    return this.columnNames;
  }
  getFormSheetColumnNames(): String[] {
    return this.sheetColumnNames;
  }

  getFormWorkgroupID(): String {
    return this.workgroupId;
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

export const mappingForm = new MappingForm();
