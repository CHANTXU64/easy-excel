export type CellValue = null | number | string | boolean | Date;

export interface Image {
  extension: 'jpeg' | 'png' | 'gif';
  filename: string;
}

export interface ImagePosition {
	tl: { col: number; row: number };
	ext: { width: number; height: number };
}
