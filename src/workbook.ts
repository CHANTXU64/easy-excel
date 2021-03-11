import * as ExcelJS from 'exceljs';
import { Workbook, Worksheet, Image } from './index';
import { __Worksheet__ } from './worksheet';
import { copyObject } from './copy';

export class __Workbook__ implements Workbook {
  private realWorkbook: ExcelJS.Workbook;
  private worksheets: Worksheet[] = [];

  constructor (realWorkbook: ExcelJS.Workbook = new ExcelJS.Workbook()) {
    this.realWorkbook = realWorkbook;
    this.realWorkbook.eachSheet(realWorksheet => {
      const worksheet = this.transWorksheet(realWorksheet);
      this.worksheets.push(worksheet);
    });
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
    let newBook = new __Workbook__();
    this.eachSheet(sourceSheet => {
      let newSheet = newBook.addWorksheet();
      newSheet = sourceSheet.clone();
    });
    newBook.realWorkbook.properties = copyObject(this.realWorkbook.properties);
    return newBook;
  }

  private transWorksheet (realWorksheet: ExcelJS.Worksheet): Worksheet {
    return new __Worksheet__(this, realWorksheet);
  }
}
