export namespace Utils {
  export const getRows = (
    range: GoogleAppsScript.Spreadsheet.Range
  ): unknown[][] => {
    const [header, ...body] = range.getValues();
    return body.map((row) => {
      header.forEach((col, index) => {
        if (isNaN(Number(col))) {
          Object.defineProperty(row, col, { value: row[index] });
        }
      });
      return row;
    });
  };
}
