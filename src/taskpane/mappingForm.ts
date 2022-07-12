/* eslint-disable @typescript-eslint/no-unused-vars */
// NOTE: Logic of working with Excel

import { Mapping } from "@provide/types";
import { store } from "../settings/store";
import { onError } from "../common/common";
import { encodeForHTML } from "../common/validate";
import { excelWorker } from "./excel-worker";

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension */

const updateWorkgroupMappings = async (mapping: Mapping): Promise<void> => {
  var existingModel = mapping.models[1];
  var sheetName;
  var sheetColumnNames;

  await Excel.run(async (context: Excel.RequestContext) => {
    sheetName = await getSheetName(context);
    sheetColumnNames = await getSheetColumnNames(context);
  }).catch(catchError);

  var mappingForm = document.getElementById("mapping-form-options");
  document.getElementById("mapping-form-header").innerHTML = "Update Mapping";

  var pkOptions = await addOptions(sheetColumnNames);
  mappingForm.innerHTML =
    `<div class="form-group container">
				<div class="row">
				<label class="col" for="table-name"> Table Name: </label>
				<input pattern="[\\w\\d\\s-]+" id="table-name" name="tableName" type="text" value ="` +
    encodeForHTML(sheetName) +
    `" class="col form-control bg-transparent text-light shadow-none" \\>
				</div>
				</div>
				<div class="form-group container">
				<div class="row">
				<label class="col" for="primary-key"> Primary Key Column: ${existingModel.primaryKey}</label>
				<select id="primary-key" name="primaryKey" class="col form-control bg-transparent text-light shadow-none">` +
    pkOptions +
    `</select>
				</div>
				</div>`;

  //Iterate through exisiting model
  //Get field name and refFieldId --> select name="refID", value="myCOlumnName">Display=myCOlumnName</select>
  //Display column names
  existingModel.fields.map(async (field) => {
    var columnOptions = await addOptions(sheetColumnNames);
    mappingForm.innerHTML +=
      `<div class="form-group container float-right">
					<div class="row">
					<label class="col d-flex justify-content-end" for="` +
      encodeForHTML(field.refFieldId) +
      `">` +
      encodeForHTML(field.name) +
      `</label>
					<select id="` +
      encodeForHTML(field.refFieldId) +
      `name=` +
      encodeForHTML(field.refFieldId) +
      `" class="col form-control bg-transparent text-light shadow-none">` +
      field.type +
      `</select>` +
      columnOptions +
      `</div>
					</div>`;
  });

  var submitButton = document.getElementById("mapping-form-btn");
  submitButton.innerHTML = "Update Mapping";
  submitButton.onclick = function () {
    excelWorker.createInitialSetup(mapping);
  };
};

export const showUnmappedColumns = async (mapping: Mapping): Promise<void> => {
  var existingModel = mapping.models[0];
  var sheetName;
  var sheetColumnNames;

  await Excel.run(async (context: Excel.RequestContext) => {
    sheetName = await getSheetName(context);
    sheetColumnNames = await getSheetColumnNames(context);
  }).catch(catchError);

  var mappingForm = document.getElementById("mapping-form-options");
  document.getElementById("mapping-form-header").innerHTML = "Create New Mapping";

  var pkOptions = await addOptions(sheetColumnNames);
  mappingForm.innerHTML =
    `<div class="form-group container">
				<div class="row">
				<label class="col" for="table-name"> Table Name: ${existingModel.type} </label>
				<input pattern="[\\w\\d\\s-]+" id="table-name" name="tableName" type="text" value ="` +
    encodeForHTML(sheetName) +
    `" class="col form-control bg-transparent text-light shadow-none" \\>
				</div>
				</div>
				<div class="form-group container">
				<div class="row">
				<label class="col" for="primary-key"> Primary Key Column: ${existingModel.primaryKey}</label>
				<select id="primary-key" name="primaryKey" class="col form-control bg-transparent text-light shadow-none">` +
    pkOptions +
    `</select>
				</div>
				</div>`;

  //Iterate through exisiting model
  //Get field name and refFieldId --> select name="refID", value="myCOlumnName">Display=myCOlumnName</select>
  //Display column names
  existingModel.fields.map(async (field) => {
    var columnOptions = await addOptions(sheetColumnNames);
    mappingForm.innerHTML +=
      `<div class="form-group container float-right">
					<div class="row">
					<label class="col d-flex justify-content-end" for="` +
      encodeForHTML(field.refFieldId) +
      `">` +
      encodeForHTML(field.name) +
      `</label>
					<select id="` +
      encodeForHTML(field.refFieldId) +
      `name=` +
      encodeForHTML(field.refFieldId) +
      `" class="col form-control bg-transparent text-light shadow-none">` +
      field.type +
      `</select>` +
      columnOptions +
      `</div>
					</div>`;
  });

  var submitButton = document.getElementById("mapping-form-btn");
  submitButton.innerHTML = "Update Mapping";
  submitButton.onclick = function () {
    excelWorker.createInitialSetup(mapping);
  };
};

export const showMappedColumns = async (mapping: Mapping): Promise<void> => {
  var counterPartyModel = mapping.models[0];
  var excelModel = mapping.models[1];

  var mappingForm = document.getElementById("mapping-form-options");
  document.getElementById("mapping-form-header").innerHTML = "Baselined Mapping";

  mappingForm.innerHTML =
    `<div class="form-group container">
				<div class="row">
				<label class="col" for="table-name"> Table Name: </label>
				<input disabled pattern="[\\w\\d\\s-]+" id="table-name" name="tableName" type="text" value ="` +
    encodeForHTML(excelModel.type) +
    `" class="col form-control bg-transparent text-light shadow-none" \\>
				</div>
				</div>
				<div class="form-group container">
				<div class="row">
				<label class="col" for="primary-key"> Primary Key Column: ${excelModel.primaryKey}</label>
				</div>
				</div>`;

  excelModel.fields.map(async (excelModelfield) => {
    var counterPartyField = counterPartyModel.fields.find((field) => field.refFieldId === excelModelfield.refFieldId);

    mappingForm.innerHTML +=
      `<div class="form-group container float-right">
              <div class="row">
              <label class="col d-flex justify-content-end">` +
      encodeForHTML(counterPartyField.name) +
      `</label>
        <input disabled pattern="[\\w\\d\\s-]+" id="table-name" name="tableName" type="text" value ="` +
      encodeForHTML(excelModelfield.name) +
      `" class="col form-control bg-transparent text-light shadow-none" \\>
          </div>
        </div>
      </div>`;
  });

  var statusButton = document.getElementById("mapping-form-btn");

  statusButton.innerHTML = "Baselined";
  statusButton.style.backgroundColor = "Green";

  var updateButton = document.getElementById("mapping-update-btn");
  updateButton.setAttribute("style", "");
  updateButton.onclick = async function () {
    await updateWorkgroupMappings(mapping);
  };
};

const getSheetName = async (context: Excel.RequestContext): Promise<String> => {
  var sheet = context.workbook.worksheets.getActiveWorksheet();
  sheet.load("name");
  await context.sync();
  return sheet.name;
};

const getSheetColumnNames = async (context: Excel.RequestContext): Promise<String[]> => {
  var sheet = context.workbook.worksheets.getActiveWorksheet().getUsedRange();
  var columns = sheet.getRow(0);
  columns.load("values");
  await context.sync();
  return columns.values[0];
};

const addOptions = async (options: String[]): Promise<String> => {
  var str = "";
  options.map(async (option, i) => {
    if (i == 0) {
      str += `<option value=${option} selected>` + encodeForHTML(option) + "</option>";
    } else {
      str += `<option value=${option} >` + encodeForHTML(option) + "</option>";
    }
  });

  return str;
};

const getTableName = async (context: Excel.RequestContext): Promise<String> => {
  var table = context.workbook.worksheets.getActiveWorksheet().getUsedRange().getTables().getFirst();
  table.load("name");
  await context.sync();
  return table.name;
};

const getColumnNames = async (context: Excel.RequestContext): Promise<String[]> => {
  var table = context.workbook.worksheets.getActiveWorksheet().getUsedRange().getTables().getFirst();
  var columns = table.getHeaderRowRange();
  columns.load("values");
  await context.sync();
  return columns.values[0];
};

const catchError = (error: any): void => {
  console.log("Error: " + error);
  if (error instanceof OfficeExtension.Error) {
    console.log("Debug info: " + JSON.stringify(error.debugInfo));
    onError(error.message);
  } else {
    onError(error);
  }
};
