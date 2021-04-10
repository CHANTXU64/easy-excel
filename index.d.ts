export type CellValue = null | number | string | boolean | Date;

export interface Cell {
  readonly workbook: Workbook;
  readonly worksheet: Worksheet;
  readonly master: Cell;

  name (addNames?: string | string[]): string[];
  value (newValue?: CellValue): CellValue;
  note (newNote?: string): string;
}

export interface Row {
  readonly worksheet: Worksheet;

  readonly number: number;
  readonly cellCount: number;

  eachCell(callback: (cell: Cell, colNumber: number) => void): void;
  addPageBreak(): void;

  copy(): Row;
}

export type worksheetState = 'visible' | 'hidden' | 'veryHidden';

export interface Worksheet {
  readonly workbook: Workbook;

  readonly rowCount: number;
  readonly lastRow: Row | undefined;

	/**
	 * Tries to find and return row for row no, else undefined
	 * 
	 * @param rowNumber The 1-index row number
	 */
  findRow(rowNumber: number): Row | undefined;

	/**
	 * Get or create rows by 1-based index
	 */
	getRow(index: number): Row;

  name: string;

  state: 'visible' | 'hidden' | 'veryHidden';

	/**
	 * Iterate over all rows (including empty rows) in a worksheet
	 */
	eachRow(callback: (row: Row, rowNumber: number) => void): void;

	/**
	 * returns the cell at [r,c] or address given by r. If not found, return undefined
	 */
	// findCell(r: number | string, c: number | string): Cell | undefined;

	/**
	 * Get or create cell
	 */
	// getCell(r: number | string, c: number | string): Cell;

  getCell(cellName: string): Cell | undefined;

	/**
	 * Using the image id from `Workbook.addImage`,
	 * embed an image within the worksheet to cover a range
	 */
	addImage(imageId: number, range: ImagePosition): void;

  clone(): Worksheet;
  copy(): Worksheet;
}

export interface Workbook {
	/**
	 * Add a new worksheet and return a reference to it
	 */
	addWorksheet(name?: string): Worksheet;

	removeWorksheet(name: string): void;

	/**
	 * fetch sheet by name
	 */
	getWorksheet(name: string): Worksheet;

	/**
	 * Iterate over all sheets.
	 *
	 * Note: `workbook.worksheets.forEach` will still work but this is better.
	 */
	eachSheet(callback: (worksheet: Worksheet) => void): void;

	/**
	 * Add Image to Workbook and return the id
	 */
	addImage(img: Image): number;

  clone(): Workbook;
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
