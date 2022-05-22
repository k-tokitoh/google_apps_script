import { Utils } from "./utils";

export namespace Integration {
  const ALL_EXPENSES_COLUMNS = ["id", "isDebit", "日付", "内容", "金額"];

  export const integrate = (
    spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet
  ) => {
    const sheets = spreadSheet.getSheets();
    const debits = sheets.filter((sheet) =>
      /^debit_20\d{4}$/.test(sheet.getName())
    );
    const nonDebit = sheets.find((sheet) => sheet.getName() === "account");

    const integrated = spreadSheet.getSheetByName("integrated");
    integrated.clear();
    integrated
      .getRange(1, 1, 1, ALL_EXPENSES_COLUMNS.length)
      .setValues([ALL_EXPENSES_COLUMNS]);

    debits.forEach((debit) => new DebitCopier(debit, integrated).execute());

    new NonDebitCopier(nonDebit, integrated).execute();

    return integrated;
  };

  abstract class Copier {
    constructor(
      private readonly from: GoogleAppsScript.Spreadsheet.Sheet,
      protected readonly to: GoogleAppsScript.Spreadsheet.Sheet
    ) {}

    execute() {
      const toRowStart = this.to.getLastRow() + 1;
      const rows = this.from.getLastRow() - 1;
      const columns = this.from.getLastColumn();
      const rawValues = this.extract(this.from.getRange(1, 1, rows, columns));
      const values = rawValues.map((row) => [
        this.generateId(row),
        this.isDebit,
        ...row,
      ]);
      const rangeTo = this.to.getRange(
        toRowStart,
        1,
        values.length,
        values[0].length
      );
      rangeTo.setValues(values);
    }

    protected abstract isDebit: boolean;

    protected abstract extract(
      range: GoogleAppsScript.Spreadsheet.Range
    ): unknown[][];

    protected generateId(values: unknown[]) {
      return values
        .map((value) => (value instanceof Date ? value.getTime() : value))
        .join("-");
    }
  }

  class DebitCopier extends Copier {
    protected isDebit = true;

    protected extract(range: GoogleAppsScript.Spreadsheet.Range): unknown[][] {
      return Utils.getRows(range)
        .filter((row) => row["お取引日"] !== " ")
        .map((row) => [row["お取引日"], row["お取引内容"], row["お取引金額"]]);
    }
  }

  class NonDebitCopier extends Copier {
    protected isDebit = false;

    protected extract(range: GoogleAppsScript.Spreadsheet.Range): unknown[][] {
      return Utils.getRows(range)
        .filter(
          (row) => row["出金金額(円)"] && !/^デビット/.test(String(row["内容"]))
        )
        .map((row) => [row["日付"], row["内容"], row["出金金額(円)"]]);
    }
  }
}
