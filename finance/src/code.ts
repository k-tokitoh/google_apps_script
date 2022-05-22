import { Utils } from "./utils";
import { Integration } from "./integration";

const finance = SpreadsheetApp.openById(
  "1xdkKYQd9g1K1zGj9f41vCzQNTfgrdBIvC_CZTh8n3v8"
);
const expenses = finance.getSheetByName("expenses");

const expenseAnnotationRules = finance.getSheetByName(
  "expense_annotation_rules"
);

const reflectCsvs = () => {
  const rawExpenses = SpreadsheetApp.getActiveSpreadsheet();
  const integrated = Integration.integrate(rawExpenses);
  updateExpenses(integrated, expenses);
  updateExpenseAnnotationRules(expenses, expenseAnnotationRules);
};

const reflectExpenseAnnotationRules = () => {
  const base = expenseAnnotationRules;
  const target = expenses;

  const dict = Object.fromEntries(
    base.getRange(1, 1, base.getLastRow(), base.getLastColumn()).getValues()
  );

  const values = Utils.getRows(target.getDataRange()).map((row) => [
    ...row.slice(0, 5),
    dict[row["内容"]],
  ]);

  target
    .getRange(2, 1, target.getLastRow() - 1, target.getLastColumn() - 2)
    .setValues(values);
};

const updateExpenses = (
  base: GoogleAppsScript.Spreadsheet.Sheet,
  target: GoogleAppsScript.Spreadsheet.Sheet
) => {
  const targetIds: unknown[] = Utils.getRows(target.getDataRange()).map(
    (targetRow) => targetRow["id"]
  );

  // パフォーマンス改善余地あり
  const rowsToAppend = Utils.getRows(base.getDataRange()).filter(
    (row) => !targetIds.includes(row["id"])
  );

  if (rowsToAppend.length) {
    target
      .getRange(
        target.getLastRow() + 1,
        1,
        rowsToAppend.length,
        rowsToAppend[0].length
      )
      .setValues(rowsToAppend);
  }
};

const updateExpenseAnnotationRules = (
  base: GoogleAppsScript.Spreadsheet.Sheet,
  target: GoogleAppsScript.Spreadsheet.Sheet
) => {
  const targetTitles: unknown[] = Utils.getRows(target.getDataRange()).map(
    (targetRow) => targetRow["内容"]
  );

  // パフォーマンス改善余地あり
  const rowsToAppend = Utils.getRows(base.getDataRange())
    .filter((row) => {
      const title = row["内容"];
      const shouldAppend = !targetTitles.includes(title);
      if (shouldAppend) {
        targetTitles.push(title);
      }
      return shouldAppend;
    })
    .map((row) => [row["内容"]]);

  if (rowsToAppend.length) {
    target
      .getRange(
        target.getLastRow() + 1,
        1,
        rowsToAppend.length,
        rowsToAppend[0].length
      )
      .setValues(rowsToAppend);
  }
};
