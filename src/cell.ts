import * as ExcelJS from 'exceljs';
import { Row } from './row';
import { CellValue } from './type';
import { copyObject } from './copy';
import { Address } from './address';

export class Cell {
  private realCell: ExcelJS.Cell;
  public readonly row: Row;

  constructor (row: Row, realCell: ExcelJS.Cell) {
    this.realCell = realCell;
    this.row = row;
  }

  get name (): string[] {
    return this.realCell.names.map(name =>
      name.replace(/__rpc__/g, "）").replace(/__lpc__/g, "（")
          .replace(/__rpe__/g, ")").replace(/__lpe__/g, "(")
    );
  }

  public addName (name: string): string[] {
    name = name.replace(/）/g, "__rpc__").replace(/（/g, "__lpc__")
               .replace(/\)/g, "__rpe__").replace(/\(/g, "__lpe__");
    this.realCell.addName(name)
    return this.name;
  }

  get value (): CellValue {
    let value = this.getValueFromRealCell(this.realCell);
    return value;
  }

  private getValueFromRealCell (realCell: ExcelJS.Cell): CellValue {
    let ValueType = ExcelJS.ValueType;
    let cellValue: any;
    let value: CellValue;
    switch (realCell.type) {
      case ValueType.Date:
        cellValue = realCell.value;
        value = cellValue;
        break;
      case ValueType.Boolean:
        cellValue = realCell.value;
        value = cellValue;
        break;
      case ValueType.Formula:
        value = realCell.result;
        break;
      case ValueType.Merge:
        value = this.getValueFromRealCell(realCell.master);
        break;
      case ValueType.Number:
        cellValue = realCell.value;
        value = cellValue;
        break;
      case ValueType.String:
        cellValue = realCell.value;
        value = cellValue;
        break;
      case ValueType.RichText:
        cellValue = realCell.value;
        let richText_arr: ExcelJS.CellRichTextValue = cellValue;
        value = "";
        richText_arr.richText.forEach(richText => {
          value += richText.text;
        });
        break;
      case ValueType.Hyperlink:
        cellValue = realCell.value;
        let hyperLinx: ExcelJS.CellHyperlinkValue = cellValue;
        value = hyperLinx.text;
        break;
      default:
        value = null;
        break;
    }
    return value;
  }

  set value (newValue: CellValue) {
    if (typeof newValue == null) {
      this.realCell.value = "";
    } else if (typeof newValue == "number" || typeof newValue == "boolean"
               || newValue instanceof Date) {
      this.realCell.value = newValue;
    } else {
      if (newValue?.[0] == "=") {
        this.setFormulaValue(newValue);
      } else {
        this.realCell.value = newValue;
      }
    }
  }

  private setFormulaValue (formula_str: string): void {
    let formula: ExcelJS.CellFormulaValue;
    formula_str = formula_str.replace(/）/g, "__rpc__")
                             .replace(/（/g, "__lpc__");
    let date1904 = this.realCell.workbook.properties.date1904;
    formula = {formula: formula_str, date1904: date1904};
    this.realCell.value = formula;
  }

  set note (note: string) {
    this.realCell.note = note;
  }

  get note (): string {
    let origNote = this.realCell.note;
    let note = "";
    if (typeof origNote != 'string') {
      origNote?.texts.forEach(text => note += text.text);
    } else {
      note = origNote;
    }
    return note;
  }

  public copy (targetCell: Cell): void {
    let targetRealCell = targetCell.realCell;
    let thisRealCell = this.realCell;
    targetRealCell.style = copyObject(thisRealCell.style);
    targetRealCell.numFmt = thisRealCell.numFmt;
    this.copyCellMerge(targetCell);
    if (thisRealCell.value && typeof thisRealCell.value == 'object') {
      targetRealCell.value = copyObject(thisRealCell.value);
    } else {
      targetRealCell.value = thisRealCell.value;
    }
  }

  private copyCellMerge (targetCell: Cell): void {
    let targetRealCell = targetCell.realCell;
    let thisRealCell = this.realCell;
    if (thisRealCell.isMerged && thisRealCell.model.hasOwnProperty("master")) {
      let thisMasterPos = Address.address2rc(thisRealCell.model.master);
      let thisCellPos = Address.address2rc(thisRealCell.model.address.address);
      let relativePos = Address.calcRelativePos(thisCellPos, thisMasterPos);
      let targetCellPos = Address.address2rc(
        targetRealCell.model.address.address);
      let targetMasterPos = Address.calcTargetPos(targetCellPos, relativePos);
      targetRealCell.model.master = Address.rc2address(targetMasterPos);
      let targetWorksheet = targetRealCell.worksheet;
      let targetMasterCell = targetWorksheet.getCell(
        targetMasterPos.r, targetMasterPos.c);
      targetRealCell.merge(targetMasterCell);
    }
  }
}

