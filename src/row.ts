import * as ExcelJS from 'exceljs';
import { Worksheet } from './worksheet';
import { Cell } from './cell';
import {Address} from './address';

export class Row {
  public readonly worksheet: Worksheet;
  private realRow: ExcelJS.Row;
  private cells: Cell[] = [];

  constructor (worksheet: Worksheet, realRow: ExcelJS.Row) {
    this.worksheet = worksheet;
    this.realRow = realRow;
    realRow.eachCell({includeEmpty: true}, realCell => {
      this.cells.push(this.transCell(realCell));
    });
  }

  get number (): number {
    return this.realRow.number;
  }

  get cellCount (): number {
    return this.realRow.cellCount;
  }

  public getCell (col: number | string): Cell {
    let realCell = this.realRow.getCell(col);
    let colNumber = Address.address2rc(realCell.address).c;
    let cell = this.cells[colNumber - 1];
    if (!cell) {
      cell = this.transCell(realCell);
      this.cells[colNumber - 1] = cell;
    }
    return cell;
  }

  public eachCell (callback: (cell: Cell, colNumber: number) => void): void {
    const l = this.cells.length;
    for (let i = 1; i <= l; ++i) {
      callback(this.getCell(i), i);
    }
  }

  public addPageBreak (): void {
    this.realRow.addPageBreak();
  }

  public copy (targetRow: Row): void {
    targetRow.realRow.height = this.realRow.height;
    this.eachCell((cell, colNumber) => {
      let targetCell = targetRow.getCell(colNumber);
      cell.copy(targetCell);
    });
  }

  private transCell (realCell: ExcelJS.Cell): Cell {
    let cell = new Cell(this, realCell);
    return cell;
  }
}
