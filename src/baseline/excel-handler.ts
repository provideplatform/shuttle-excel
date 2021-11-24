// eslint-disable-next-line no-unused-vars
/* global Excel, OfficeExtension, Office */

export class ExcelHandler {
  async getTableName(context: Excel.RequestContext): Promise<string> {
    let table = context.workbook.worksheets.getActiveWorksheet().getUsedRange().getTables().getFirst();
    table.load("name");
    await context.sync();
    return table.name;
  }

  async getColumnAddress(context: Excel.RequestContext, column: String): Promise<string> {
    //Get column Header Cell
    let table = context.workbook.worksheets.getActiveWorksheet().getUsedRange().getTables().getFirst();
    let headerRange = table.getHeaderRowRange();
    let headerCell = headerRange.findOrNullObject(column.toString(), { completeMatch: true });
    headerCell.load("address");

    await context.sync();

    var headerCellAddress = headerCell.address.split("!")[1];
    var columnAddress = headerCellAddress.split(/\d+/)[0];

    return columnAddress;
  }

  async getDataColumnHeader(context: Excel.RequestContext, changedData: Excel.TableChangedEventArgs): Promise<string> {
    let dataColumn = context.workbook.worksheets
      .getActiveWorksheet()
      .getRange(changedData.address.split(/\d+/)[0] + "1");
    dataColumn.load("values");
    return context.sync().then(() => {
      return dataColumn.values[0][0];
    });
  }

  async getDataCellAddress(
    context: Excel.RequestContext,
    primaryKeyValue: string,
    columnName: string,
    primaryKeyColumnAddress: string
  ): Promise<string> {
    //Get column Header Cell
    var column = await this.getColumnAddress(context, columnName);

    //Get Primary Key Cell
    let columnAddress = primaryKeyColumnAddress;
    let primaryKeyRange = context.workbook.worksheets
      .getActiveWorksheet()
      .getRange(columnAddress + ":" + columnAddress);
    let primaryKeyCell = primaryKeyRange.findOrNullObject(primaryKeyValue, { completeMatch: true });
    primaryKeyCell.load("address");

    await context.sync();

    var row = primaryKeyCell.address.split(/\D+/)[1];
    return column + row;
  }
}

export const excelHandler = new ExcelHandler();
