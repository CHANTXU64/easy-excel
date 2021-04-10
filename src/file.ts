import { Workbook } from './workbook';
import * as ExcelJS from 'exceljs';

export class ExcelFile {
  private workbook: Workbook;

  constructor () {
    this.workbook = new Workbook();
  }

  public load (fileName: string): Promise<Workbook> {
    return new Promise((resolve, rejects) => {
      let realWorkbook = new ExcelJS.Workbook();
      realWorkbook.xlsx.readFile(fileName).then(realWorkbook => {
        this.workbook = new Workbook(realWorkbook);
        resolve(this.workbook);
      }).catch(() => {
        rejects();
      });
    });
  }

  public save (fileName: string): Promise<unknown> {
    return new Promise((resolve, rejects) => {
      this.workbook.export(fileName).then(() => {
        resolve();
      }).catch(() => {
        rejects();
      });
    });
  }
}
