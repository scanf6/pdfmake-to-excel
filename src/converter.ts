import { saveAs } from "file-saver";
import {Buffer} from 'exceljs';
import * as ExcelJS from 'exceljs';
//const ExcelJS = require('exceljs');
import { IPayload } from './interfaces/IPayload.interface';
import { SheetDefaultOptions } from './types/sheetOptions.type';
import renderFunction from "./renderFunction";

export class ExcelConverter {
	constructor(
		private filename:String,
		private payload:IPayload,
		private sheetDefaultOptions:SheetDefaultOptions = { defaultColWidth: 20}
	) {}

	downloadExcel() {
		const workbook = new ExcelJS.Workbook();

		renderFunction(workbook, this.payload, this.sheetDefaultOptions).xlsx.writeBuffer().then((data:Buffer) => {
			var blob = new Blob([data], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
			saveAs(blob, this.filename + '.xlsx');
		});
	}
}