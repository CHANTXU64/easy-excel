import * as ExcelJS from 'exceljs';
import { Image } from './type';
import { Worksheet } from './worksheet';
import { copyObject } from './copy';

export class Workbook {
  private readonly realWorkbook: ExcelJS.Workbook;
  private worksheets: Worksheet[] = [];

  constructor (realWorkbook: ExcelJS.Workbook = new ExcelJS.Workbook()) {
    this.realWorkbook = realWorkbook;
    this.realWorkbook.eachSheet(realWorksheet => {
      const worksheet = this.transWorksheet(realWorksheet);
      this.worksheets.push(worksheet);
    });
  }

  get date1904 (): boolean {
    return this.realWorkbook.properties.date1904;
  }

  set date1904 (bool: boolean) {
    this.realWorkbook.properties.date1904 = bool;
  }

  get fullCalcOnLoad (): boolean {
    return this.realWorkbook.calcProperties.fullCalcOnLoad;
  }

  set fullCalcOnLoad (bool: boolean) {
    this.realWorkbook.calcProperties.fullCalcOnLoad = bool;
  }

  public addWorksheet (name?: string): Worksheet {
    const realWorksheet = this.realWorkbook.addWorksheet(name);
    const worksheet = this.transWorksheet(realWorksheet);
    this.worksheets.push(worksheet);
    return worksheet;
  }

  public removeWorksheet (name: string): void {
    this.realWorkbook.removeWorksheet(name);
    const i = this.worksheets.findIndex(worksheet => worksheet?.name == name);
    delete this.worksheets[i];
  }

  public getWorksheet (name: string): Worksheet | undefined {
    return this.worksheets.find(worksheet => worksheet?.name == name);
  }

  public eachSheet (callback: (worksheet: Worksheet) => void): void {
    this.worksheets.forEach(sheet => {
      callback(sheet);
    });
  }

  public addImage (img: Image): number {
    return this.realWorkbook.addImage(img);
  }

  public clone (): Workbook {
    let newBook = new Workbook();
    this.eachSheet(sourceSheet => {
      let targetSheet = newBook.addWorksheet(sourceSheet.name);
      sourceSheet.copy(targetSheet);
    });
    newBook.realWorkbook.properties = copyObject(this.realWorkbook.properties);
    return newBook;
  }

  public export (fileName: string): Promise<void> {
    return new Promise((resolve, rejects) => {
      this.realWorkbook.xlsx.writeFile(fileName).then(() => {
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

  private transWorksheet (realWorksheet: ExcelJS.Worksheet): Worksheet {
    return new Worksheet(this, realWorksheet);
  }
}
