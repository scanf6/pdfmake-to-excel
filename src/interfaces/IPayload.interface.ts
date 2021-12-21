export interface IPayload {
	title?: string;
	campaign?: string;
	situation?: string;
	logo?: string;
	data: ICell[][]
}

export interface ICell {
	text:string;
	rowSpan?:number | null | undefined;
	colSpan?:number | null | undefined;
	border?: number[]
}