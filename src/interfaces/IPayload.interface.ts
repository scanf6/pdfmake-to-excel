export interface IPayload {
	title?: string;
	campaign?: string;
	situation?: string;
	logo?: string;
	data: ICell[][] | ISheetData[]
}

export interface ICell {
	text:string;
	type?:string;
	rowSpan?:number | null | undefined;
	colSpan?:number | null | undefined;
	border?: number[]
}

export interface ISheetData {
	sheetName:string;
	sheetData: ICell[][]
}