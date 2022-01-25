import { saveAs } from "file-saver";
import {Buffer} from 'exceljs';
import * as ExcelJS from 'exceljs';
import { IPayload } from './interfaces/IPayload.interface';
import { IDefaultOptions } from './interfaces/IDefaultOptions.interface';
import renderFunction from "./renderFunction";

// Testing streams with ExcelJS
const Stream = require('stream');


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
			let blob = new Blob([data], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
			saveAs(blob, this.filename + '.xlsx');
		});
	}

	async streamExcel() {
		console.log('Streaming')
		const stream = new Stream.PassThrough();
		const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ stream })
		await renderFunction(workbook, this.payload, this.options);
		workbook.commit();
	}
}