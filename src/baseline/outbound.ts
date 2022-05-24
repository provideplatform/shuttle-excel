/* eslint-disable @typescript-eslint/no-unused-vars */
import { onError } from "../common/common";
import { Object, BaselineResponse, Workstep, ProtocolMessagePayload } from "@provide/types";
import { ProvideClient } from "src/client/provide-client";
import { excelHandler } from "./excel-handler";
import { store } from "../settings/store";
import { localStore } from "../settings/settings";
import { getMyWorkgroups } from "../taskpane/taskpane";

// eslint-disable-next-line no-unused-vars
/* global Excel, Office, OfficeExtension */

export class OutBound {
  async send(
    context: Excel.RequestContext,
    changedData: Excel.TableChangedEventArgs | Excel.WorksheetChangedEventArgs,
    identClient: ProvideClient
  ): Promise<void> {
    try {
      //Get the primary key ID
      let tableName = await excelHandler.getSheetName(context);
      let tableID = (await store.getTableID(tableName)).toString();
      let primaryKeyColumnName = await store.getPrimaryKeyColumnName(tableID);
      let primaryKeyID = await this.getPrimaryKeyID(context, changedData, primaryKeyColumnName);

      //Check if baselineID exists
      let baselineIDExists = await store.keyExists(tableID, primaryKeyID);

      //Create the message
      let message = await this.createMessage(context, primaryKeyID, changedData);

      console.log(JSON.stringify(message));
      let baselineResponse: BaselineResponse;
      var worksteps;
      var workstep: Workstep;
      var protocolMessage: ProtocolMessagePayload;

      if (!baselineIDExists) {
        //Get the workflowID from localStorage (Do While)
        await getMyWorkgroups(true);
        var workflowID = null;
        do {
          workflowID = await localStore.getWorkflowID();
        } while (workflowID == null);

        //RESOLVE workstep
        worksteps = await identClient.getWorksteps(workflowID);
        workstep = await identClient.resolveWorkstep(worksteps);

        //EXECUTE WORKSTEP
        protocolMessage = await identClient.executeWorkstep(workflowID, workstep.id, message);
        if (!protocolMessage) {
          await this.preventEdit(context, changedData);
        } else {
          baselineResponse = await identClient.sendCreateProtocolMessage(protocolMessage);

          console.log(baselineResponse);
          await store.setBaselineIDAndWorkflowID(tableID, primaryKeyID, [baselineResponse.baselineId, workflowID]);
          await localStore.removeWorkflowID();
        }
      } else {
        let [baselineId, workflowID] = await store.getBaselineIdAndWorkflowID(tableID, primaryKeyID);

        //RESOLVE workstep
        worksteps = await identClient.getWorksteps(workflowID);
        workstep = await identClient.resolveWorkstep(worksteps);

        //EXECUTE WORKSTEP
        protocolMessage = await identClient.executeWorkstep(workflowID, workstep.id, message);
        if (!protocolMessage) {
          await this.preventEdit(context, changedData);
        } else {
          baselineResponse = await identClient.sendCreateProtocolMessage(protocolMessage);

          baselineResponse = await identClient.sendUpdateProtocolMessage(baselineId, message);
          console.log("Baseline message : " + baselineResponse);
        }
      }
    } catch {
      this.catchError;
    }
  }

  private async createMessage(
    context: Excel.RequestContext,
    primaryKeyID: string,
    changedData: Excel.TableChangedEventArgs | Excel.WorksheetChangedEventArgs
  ): Promise<Object> {
    //Get the refModelId (tableName --> refModelId)
    //Get the primary key column (refModelID --> primaryKeyColumn)
    //Get the primary key id
    //TODO: Check if the primary key is mapped to any baselineID
    //If yes, then resolve workstep ---> updateObject
    //If no, ask user to create and store the workflow ID for each row.

    let id = primaryKeyID;
    let dataColumnHeader = await excelHandler.getDataColumnHeader(context, changedData);

    let message: Object = {} as Object;
    message.type = "general_consistency";
    message.id = id;

    const data = {};
    data[dataColumnHeader] = changedData.details.valueAfter;

    let _payload = {
      id: id,
      data: data,
    };

    message.payload = _payload;

    return message;
  }

  private async getPrimaryKeyID(
    context: Excel.RequestContext,
    changedData: Excel.TableChangedEventArgs | Excel.WorksheetChangedEventArgs,
    primaryKeyColumnName: String
  ): Promise<string> {
    try {
      let primaryKeyColumnAddress = await excelHandler.getSheetColumnAddress(context, primaryKeyColumnName);
      let primaryKeyCell = primaryKeyColumnAddress + changedData.address.split(/\D+/)[1];
      let primaryKeyID = context.workbook.worksheets
        .getActiveWorksheet()
        .getRange(primaryKeyCell + ":" + primaryKeyCell);

      primaryKeyID.load("values");
      await context.sync();

      console.log(primaryKeyID.values);
      return primaryKeyID.values[0][0].toString();
    } catch {
      this.catchError;
    }
  }

  private async preventEdit(
    context: Excel.RequestContext,
    changedData: Excel.TableChangedEventArgs | Excel.WorksheetChangedEventArgs
  ): Promise<void> {
    var range = context.workbook.worksheets.getActiveWorksheet().getRange();
    range.format.protection.locked = false;

    var cell = context.workbook.worksheets.getActiveWorksheet().getRange(changedData.address);
    cell.format.protection.locked = true;

    await context.sync();
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

export const outboundMessage = new OutBound();
