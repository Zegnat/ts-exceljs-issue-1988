#!/usr/bin/env -S npx ts-node

import { CellHyperlinkValue, CellValue, Workbook } from "exceljs";

// Type-guard to check if a CellValue is a Hyperlink.
const isHyperlink = (cellValue: CellValue): cellValue is CellHyperlinkValue =>
  typeof cellValue === "object" &&
  "text" in cellValue &&
  "hyperlink" in cellValue;

(async () => {
  const workbook = new Workbook();
  await workbook.xlsx.readFile(__dirname + "/example.xlsx");

  workbook.worksheets[0].getColumn(1).eachCell((cell) => {
    const cellValue = cell.value; // Type: CellValue
    if (!isHyperlink(cellValue)) return; // Narrowing the cellValue type.
    cellValue; // Type: CellHyperlinkValue
    const cellTextValue = cellValue.text; // Type: string

    console.log(typeof cellTextValue);
  });

  /** Console output:
   * string
   * object <-- Should not be possible according to typing
   */
})();
