import { Workbook } from './workbook';
import * as ExcelJS from 'exceljs';

export class ExcelFile {
  private workbook: Workbook;

  constructor (workbook = new Workbook) {
    this.workbook = workbook;
  }

  public load (fileName: string): Promise<Workbook> {
    return new Promise((resolve, rejects) => {
      let realWorkbook = new ExcelJS.Workbook();
      realWorkbook.xlsx.readFile(fileName).then(realWorkbook => {
        this.workbook = new Workbook(realWorkbook);
        if (resolve) {
          resolve(this.workbook);
        }
      }).catch(() => {
        if (rejects) {
          rejects();
        }
      });
    });
  }

  public save (fileName: string): Promise<void> {
    return new Promise((resolve, rejects) => {
      this.workbook.export(fileName).then(() => {
        if (resolve) {
          resolve();
        }
      }).catch(() => {
        if (rejects) {
          rejects();
        }
      });
    });
  }
}
