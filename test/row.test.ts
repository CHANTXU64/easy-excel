import * as ExcelJS from 'exceljs';
import { Cell } from '../src/cell';
import { Row } from '../src/row';
import { Worksheet } from '../src/worksheet';

class TESTRow extends Row {
  public readonly _realRow: ExcelJS.Row;
  public _cells: Cell[] = [];

  constructor (worksheet: Worksheet, realRow: ExcelJS.Row) {
    super(worksheet, realRow);
    this._realRow = this.realRow;
    this._cells = this.cells;
  }

  public _transCell (realCell: ExcelJS.Cell): Cell {
    return this.transCell(realCell);
  }
}

describe("private cells 更新测试", () => {
})
