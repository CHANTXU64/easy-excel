import { ExcelFile } from '../src/file';

let a = new ExcelFile();
a.load('./test/test.xlsx').then(workbook => {
  workbook.export('./test/export.xlsx');
})
