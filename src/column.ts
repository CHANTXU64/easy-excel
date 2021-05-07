import { Cell } from './cell';
import { Worksheet } from './worksheet';

export class Column {
  public readonly worksheet: Worksheet;
  public readonly number: number;

  constructor (worksheet: Worksheet, colNumber: number) {
    this.worksheet = worksheet;
    this.number = colNumber;
  }

  get cellCount (): number {
    return this.worksheet.rowCount;
  }
  
  public getCell (rowNumber: number): Cell {
    return this.worksheet.getCell(rowNumber, this.number);
  }

  public eachCell (callback: (cell: Cell, rowNumber: number) => void): void {
    const l = this.cellCount;
    for (let i = 1; i <= l; ++i) {
      callback(this.getCell(i), i);
    }
  }
}
