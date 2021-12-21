import { saveAs } from "file-saver";
import {Buffer} from 'exceljs';
import * as ExcelJS from 'exceljs';
import { IPayload } from './interfaces/IPayload.interface';
import { IDefaultOptions } from './interfaces/IDefaultOptions.interface';
import renderFunction from "./renderFunction";

export class ExcelConverter {
	constructor(
		private filename:String,
		private payload:IPayload,
		private options:IDefaultOptions = {
			protection: undefined,
			defaultOptions: {defaultColWidth: 20}
		},

	) {}

	async downloadExcel() {
		const workbook = new ExcelJS.Workbook();
		let renderer = await renderFunction(workbook, this.payload, this.options);


		renderer.xlsx.writeBuffer().then((data:Buffer) => {
			var blob = new Blob([data], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
			saveAs(blob, this.filename + '.xlsx');
		});
	}
}