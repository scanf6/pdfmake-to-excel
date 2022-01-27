import { saveAs } from "file-saver";
import {Buffer} from 'exceljs';
import * as ExcelJS from 'exceljs';
import { IPayload } from './interfaces/IPayload.interface';
import { IDefaultOptions } from './interfaces/IDefaultOptions.interface';
import renderFunction from "./renderFunction";
const {Readable} = require('stream');

export class ExcelConverter {
	constructor(
		private filename:String,
		private payload:IPayload,
		private options:IDefaultOptions = {
			protection: undefined,
			defaultOptions: {defaultColWidth: 20}
		},

	) {}

	/**
	 * Front-End purposes: Create the Excel File and starts the download
	 */
	async downloadExcel():Promise<void> {
		const workbook = new ExcelJS.Workbook();
		let renderer = await renderFunction(workbook, this.payload, this.options);


		renderer.xlsx.writeBuffer().then((data:Buffer) => {
			let blob = new Blob([data], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
			saveAs(blob, this.filename + '.xlsx');
		});
	}

	/**
	 * Back-End purposes: Create a readable stream of data that you can pipe to a response request
	 */
	async getStream(response=null) {
		console.log("response", response)
		const workbook = new ExcelJS.Workbook();
		let renderer = await renderFunction(workbook, this.payload, this.options);
		console.log("rendered", renderer)

		renderer.xlsx.writeBuffer().then((data:Buffer) => {
			console.log("data", data);
			const stream = Readable.from(data);
			if(response) {
				stream.pipe(response);
				return null;
			}
			else return stream;
		});
	}
}