export type CellValue = null | number | string | boolean | Date;

export interface Cell {
  readonly row: Row;
  readonly name: string[];
  value: CellValue;
  note: string;
  addName (name: string): string[];
  copy (targetCell: Cell): void;
}

export interface Row {
  readonly worksheet: Worksheet;
  readonly number: number;
  readonly cellCount: number;

  getCell (col: number | string): Cell;
  eachCell (callback: (cell: Cell, colNumber: number) => void): void;
  addPageBreak (): void;
  copy (targetRow: Row): void;
}

export type worksheetState = 'visible' | 'hidden' | 'veryHidden';

export interface Worksheet {
  readonly workbook: Workbook;
  readonly rowCount: number;
  readonly lastRow: Row | undefined;
  name: string;
  state: 'visible' | 'hidden' | 'veryHidden';

	/**
	 * Tries to find and return row for row no, else undefined
	 * 
	 * @param rowNumber The 1-index row number
	 */
  findRow (rowNumber: number): Row | undefined;

	/**
	 * Get or create rows by 1-based index
	 */
	getRow (index: number): Row;

	/**
	 * Iterate over all rows (including empty rows) in a worksheet
	 */
	eachRow (callback: (row: Row, rowNumber: number) => void): void;

	/**
	 * returns the cell at [r,c] or address given by r. If not found, return undefined
	 */
	// findCell (r: number | string, c: number | string): Cell | undefined;

	/**
	 * Get or create cell
	 */
	getCell (row: number, col: number | string): Cell;
  getCell (address: string): Cell;

  getCellByName (cellName: string): Cell | undefined;

	/**
	 * Using the image id from `Workbook.addImage`,
	 * embed an image within the worksheet to cover a range
	 */
	addImage (imageId: number, range: ImagePosition): void;

  copy (targetSheet: Worksheet): void;
}

export interface Workbook {
	/**
	 * Add a new worksheet and return a reference to it
	 */
	addWorksheet (name?: string): Worksheet;

	removeWorksheet (name: string): void;

	/**
	 * fetch sheet by name
	 */
	getWorksheet (name: string): Worksheet | undefined;

	/**
	 * Iterate over all sheets.
	 *
	 * Note: `workbook.worksheets.forEach` will still work but this is better.
	 */
	eachSheet (callback: (worksheet: Worksheet) => void): void;

	/**
	 * Add Image to Workbook and return the id
	 */
	addImage (img: Image): number;

  clone (): Workbook;

  export (fileName: string): Promise<void>;
}

export interface Image {
  extension: 'jpeg' | 'png' | 'gif';
  filename: string;
}

export interface ImagePosition {
	tl: { col: number; row: number };
	ext: { width: number; height: number };
}

export class ExcelFile {
  public load (fileName: string): Promise<Workbook>;
  public save (filename: string): Promise<void>;
}
