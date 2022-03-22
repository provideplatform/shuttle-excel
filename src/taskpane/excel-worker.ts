/* eslint-disable @typescript-eslint/no-unused-vars */
// NOTE: Logic of working with Excel

import { Application, Workflow, Workstep, Organization } from "@provide/types";
import { onError } from "../common/common";
import { baseline } from "../baseline/index";
import { ProvideClient } from "src/client/provide-client";
import { MappingForm } from "./mappingForm";
import { paginate } from "../common/paginate";

// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension */

export class ExcelWorker {
  identClient: ProvideClient | null;

  async showOrganizations(organizations: Organization[]): Promise<void> {
    paginate(organizations, "organizations-list");
  }

  async showWorkgroups(sheetName: string, applications: Application[]): Promise<void> {
    paginate(applications, "workgroups-list");
  }
  async showWorkflows(workflows: Workflow[]): Promise<void> {
    paginate(workflows, "workflows-list");
  }

  async showMappingButton(): Promise<void> {
    var completelist = document.getElementById("workgroup-mapping");
    completelist.innerHTML = "";

    //TO SECURE --> innerHTML https://newbedev.com/xss-prevention-and-innerhtml
    completelist.innerHTML += `<button type="button" class="btn btn-primary btn-sm float-right" id="mapping-btn">Mappings</button>`;
  }

  async showWorksteps(worksteps: Workstep[]): Promise<void> {
    paginate(worksteps, "worksteps-list");
  }

  async createInitialSetup(mappingForm: MappingForm): Promise<unknown> {
    //return baseline.createTableMappings(mappingForm);
    return baseline.createSheetMappings(mappingForm);
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
