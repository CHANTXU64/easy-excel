import * as ExcelJS from 'exceljs';
import { ImagePosition } from './type';
import { Workbook } from './workbook';
import { Row } from './row';
import { Cell } from './cell';
import { copyObject } from './copy';
import { Address } from './address';
import { Column } from './column';

export class Worksheet {
  public readonly workbook: Workbook;

  private readonly realWorksheet: ExcelJS.Worksheet;
  private rows: Row[] = [];

  constructor (workbook: Workbook, realWorksheet: ExcelJS.Worksheet) {
    this.workbook = workbook;
    this.realWorksheet = realWorksheet;
    this.realWorksheet.eachRow({ includeEmpty: true }, realRow => {
      const row = this.transRow(realRow);
      this.rows.push(row);
    });
  }

  get rowCount (): number {
    return this.realWorksheet.rowCount;
  }

  get columnCount (): number {
    return this.realWorksheet.columnCount;
  }

  get lastRow (): Row | undefined {
    let rowCount = this.rowCount;
    if (rowCount) {
      return this.rows[rowCount];
    } else {
      return undefined;
    }
  }

  public findRow (rowNumber: number): Row | undefined {
    return this.rows[rowNumber - 1];
  }

  public getRow (rowNumber: number): Row {
    if (!this.rows[rowNumber - 1]) {
      const newRealRow = this.realWorksheet.getRow(rowNumber);
      this.rows[rowNumber - 1] = this.transRow(newRealRow);
    }
    return this.rows[rowNumber - 1];
  }

  public getColumn (colNumber: number): Column {
    return new Column(this, colNumber);
  }

  set name (newName: string) {
    this.realWorksheet.name = newName;
  }

  get name (): string {
    return this.realWorksheet.name;
  }

  set state (newState: 'visible' | 'hidden' | 'veryHidden') {
    this.realWorksheet.state = newState;
  }

  get printArea (): string {
    return this.realWorksheet.pageSetup.printArea;
  }

  set printArea (area: string) {
    this.realWorksheet.pageSetup.printArea = area;
  }

  get state (): 'visible' | 'hidden' | 'veryHidden' {
    return this.realWorksheet.state;
  }

  public eachRow (callback: (row: Row, rowNumber: number) => void): void {
    const l = this.rows.length;
    for (let i = 1; i <= l; ++i) {
      callback(this.getRow(i), i);
    }
  }

  public eachColumn (callback: (column: Column, colNumber: number) => void): void {
    const l = this.columnCount;
    for (let i = 1; i <= l; ++i) {
      callback(this.getColumn(i), i);
    }
  }

  public getCell (row: number, col: number | string): Cell;
  public getCell (address: string): Cell;

  public getCell (a: number | string, b?: number | string): Cell {
    let cell: Cell;
    if (typeof b == 'undefined' && typeof a == 'string') {
      cell = this.getCellEx(a);
    } else {
      let address: string;
      if (typeof b == 'number' && typeof a == 'number') {
        address = Address.rc2address({c: b, r: a});
      } else {
        address = '' + b + a;
      }
      cell = this.getCellEx(address);
    }
    return cell;
  }

  private getCellEx (address: string): Cell {
    let rowcol = Address.address2rc(address);
    let row = this.getRow(rowcol.r);
    return row.getCell(rowcol.c);
  }

  public getCellByName (cellName: string): Cell | undefined {
    this.eachRow(row => {
      row.eachCell(cell => {
        if (cell.name.indexOf(cellName) != -1) {
          return cell;
        }
      });
    });
    return undefined;
  }
  
  public addImage (imageId: number, range: ImagePosition): void {
    this.realWorksheet.addImage(imageId, range);
  }

  public copy (targetSheet: Worksheet): void {
    this.eachRow((row, rowNumber) => {
      let targetRow = targetSheet.getRow(rowNumber);
      row.copy(targetRow);
    });
    this.copyDefineNames(targetSheet);
    this.copyDefineNames(targetSheet);
    this.copyColumnsWidth(targetSheet);
    this.copyPageProperties(targetSheet);
    this.copyHeaderFooter(targetSheet);
  }

  public copyDefineNames (targetSheet: Worksheet): void {
    this.eachRow((row, rowNumber) => {
      let targetRow = targetSheet.getRow(rowNumber);
      this.copyRowDefineNames(row, targetRow);
    });
  }

  public copyPageProperties (targetSheet: Worksheet): void {
    targetSheet.realWorksheet.properties = copyObject(this.realWorksheet.properties);
    targetSheet.realWorksheet.pageSetup = copyObject(this.realWorksheet.pageSetup);
  }

  public copyHeaderFooter (targetSheet: Worksheet): void {
    targetSheet.realWorksheet.headerFooter = copyObject(this.realWorksheet.headerFooter);
  }

  public copyColumnsWidth (targetSheet: Worksheet): void {
    let sourceColumns = this.realWorksheet.columns;
    if (sourceColumns != null) {
      let targetRealSheet = targetSheet.realWorksheet;
      sourceColumns.forEach((column, index) => {
        if (column.isCustomWidth) {
          targetRealSheet.getColumn(index + 1).width = column.width;
        }
      });
    }
  }

  private copyRowDefineNames (sourceRow: Row, targetRow: Row): void {
    sourceRow.eachCell((cell, colNumber) => {
      let cellNames = cell.name;
      if (cellNames) {
        let targetCell = targetRow.getCell(colNumber);
        cellNames.forEach(name => targetCell.addName(name));
      }
    });
  }

  private transRow (realRow: ExcelJS.Row): Row {
    return new Row(this, realRow);
  }
}
