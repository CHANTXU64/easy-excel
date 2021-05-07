import { Workbook } from './workbook';
import { Worksheet } from './worksheet';
import { Row } from './row';
import { Column } from './column';
import { Cell } from './cell';
import { ExcelFile } from './file';

export interface ExcelData {[id: string]: any};

type dataType = "number" | "boolean" | "string" | "Date";
type ohead = {key: string, type: dataType};

export class utils {
  public static openFiles (filesName: string[]): Promise<Workbook[]> {
    return this.open1file(filesName, 0, []);
  }

  public static saveFiles (workbooks: Workbook[], filesName: string[]): Promise<void> {
    if (workbooks.length != filesName.length) {
      throw new Error("save files error 01.");
    }
    return this.save1file(workbooks, filesName, 0);
  }

  public static getData (worksheet: Worksheet): ExcelData[] {
    let flag = String(worksheet.getCell(1, 1).value);
    if (flag == "ROW") {
      let head = this.getHead(worksheet.getRow(1));
      let data: ExcelData[] = [];
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber != 1 && String(row.getCell(1).value) == ".data") {
          data.push(this.getOneData(head, row));
        }
      });
      return data;
    } else if (flag == "COLUMN") {
      let head = this.getHead(worksheet.getColumn(1));
      let data: ExcelData[] = [];
      worksheet.eachColumn((column, colNumber) => {
        if (colNumber != 1 && String(column.getCell(1).value) == ".data") {
          data.push(this.getOneData(head, column));
        }
      })
      return data;
    } else {
      return [];
    }
  }

  private static open1file (filesName: string[], index: number, workbooks: Workbook[]): any {
    if (index == filesName.length) {
      return Promise.resolve(workbooks);
    } else if (index < filesName.length) {
      const file = new ExcelFile();
      return file.load(filesName[index])
        .then(workbook => {
          workbooks.push(workbook);
          return this.open1file(filesName, index + 1, workbooks);
        });
    } else {
      throw new Error("open 1 file error.");
    }
  }

  private static save1file (workbooks: Workbook[], filesName: string[], index: number): any {
    if (index == filesName.length) {
      return Promise.resolve();
    } else if (index < filesName.length) {
      const file = new ExcelFile(workbooks[index]);
      return file.save(filesName[index])
        .then(() => {
          return this.save1file(workbooks, filesName, index + 1);
        });
    } else {
      throw new Error("save 1 file error.");
    }
  }

  private static getHead (headGroup: Row | Column): ohead[] {
    let head: {key: string, type: dataType}[] = [];
    headGroup.eachCell((cell, number) => {
      if (number != 1 && cell.value != null) {
        head[number] = this.getHeadKeyFromCell(cell);
      }
    });
    return head;
  }

  private static getHeadKeyFromCell (cell: Cell): ohead {
    let cellValue = cell.value == null ? "" : String(cell.value);
    let headKey = cellValue.split("/")[0];
    let typeKey = cellValue.split("/")[1];
    let type: dataType;
    if (typeKey == "n") {
      type = "number";
    } else if (typeKey == "b") {
      type = "boolean";
    } else if (typeKey == "d") {
      type = "Date";
    } else {
      type= "string";
    }
    return {key: headKey, type: type};
  }

  private static getOneData (head: ohead[], group: Row | Column): {[id: string]: any} {
    let oData: {[id: string]: any} = {};
    head.forEach((oHead, index) => {
      const key = oHead.key;
      let value = group.getCell(index).value;
      if (oHead.type == "number") {
        value = Number(value);
      } else if (oHead.type == "Date") {
        value = new Date(String(value));
      } else if (oHead.type == "boolean") {
        value = Boolean(value);
      } else {
        value = String(value == null ? "" : value);
      }
      oData[key] = value;
    });
    return oData;
  }
}

