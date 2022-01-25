import { saveAs } from "file-saver";
import {Buffer} from 'exceljs';
import * as ExcelJS from 'exceljs';
import { IPayload } from './interfaces/IPayload.interface';
import { IDefaultOptions } from './interfaces/IDefaultOptions.interface';
import renderFunction from "./renderFunction";
const Blob = require('node-blob');
const toStream = require('blob-to-stream');
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

	async downloadExcel() {
		const workbook = new ExcelJS.Workbook();
		let renderer = await renderFunction(workbook, this.payload, this.options);


		renderer.xlsx.writeBuffer().then((data:Buffer) => {
			let blob = new Blob([data], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
			saveAs(blob, this.filename + '.xlsx');
		});
	}

	async getBlob() {
		const workbook = new ExcelJS.Workbook();
		let renderer = await renderFunction(workbook, this.payload, this.options);
		const data = await renderer.xlsx.writeBuffer();

		// Returning the data as a Blob and to a stream
		const stream = Readable.from(data);
		return stream;
		//let myBlob = new Blob([data], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
		//return myBlob;
		//return new Blob([data], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
	}
}