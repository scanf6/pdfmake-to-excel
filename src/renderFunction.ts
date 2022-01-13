import { Workbook } from 'exceljs';
import { IPayload, ICell, ISheetData } from './interfaces/IPayload.interface';
import { IDefaultOptions } from './interfaces/IDefaultOptions.interface';

/**
 * Function that take a number and returns the according character
 * @param {Number} colNumber The number of the lettre. Ex: 1 -> A, 2 -> B
 * @returns {String} The letter from the received number. Ex: 1 -> A, 2 -> B
 */
 function excelColumns(colNumber:number):string {
	const start = 65;
	const end = 90;

	if(colNumber <= 0) return '';

	let charCode = (start + colNumber) - 1;

	if(charCode > end) {
		let subCode = (start + (charCode - end)) - 1;
		return `A${String.fromCharCode(subCode)}`
	}

	return String.fromCharCode(charCode);
}

/**
 * This function return the final cell number after the row and col spans are applied
 * @param {Object} payload
 * @param {String} cellNumber
 * @param {Number} letterColumnIndex
 * @param {Number} rowSpan
 * @param {Number} colSpan
 * @param {String} cellText
 * @param {Number} lineIndex
 * @param {Number} columnIndex
 * @returns {String} The final cell number for the merge
 */
function afterMerge(
	payload: ICell[][],
	cellNumber:string,
	letterColumnIndex:number,
	rowSpan:number | null | undefined=null,
	colSpan:number | null | undefined=null,
	cellText:string,
	lineIndex:number,
	columnIndex:number
) {
    let numberPart = Number(cellNumber.split('').filter(char => !Number.isNaN(Number(char))).join(''));
    let stringPart = cellNumber.split('').filter(char => Number.isNaN(Number(char))).join('');

    if(rowSpan) {
        numberPart = numberPart + (rowSpan - 1);
        payload[lineIndex + (rowSpan - 1)][columnIndex] = {...payload[lineIndex + (rowSpan - 1)][columnIndex], text: cellText, colSpan: null, rowSpan: null};
    }
    if(colSpan) {
        stringPart = excelColumns(letterColumnIndex + (colSpan - 1));
        payload[lineIndex][columnIndex + (colSpan - 1)] = {...payload[lineIndex][columnIndex + (colSpan - 1)], text: cellText, colSpan: null, rowSpan: null}
    }

    return `${stringPart}${numberPart};`
}

function isICell(object: any): object is ICell {
    return 'text' in object;
}

function isISheetData(object: any): object is ISheetData {
    return 'sheetName' in object;
}

async function sheetBuilding(workbook:Workbook, title:string | undefined, campaign:string | undefined, situation:string | undefined, logo:string | undefined, data:ICell[][], options:IDefaultOptions, sheetName:string) {
    /* BUILDING PROCESS */
    let startingLine = 0;

    let titlePositionning = excelColumns(Math.round((data[0].length) / 2));
    const {protection, defaultOptions = {defaultColWidth: 20}} = options;

    const worksheet = workbook.addWorksheet(sheetName, { properties: defaultOptions});

    if(protection) await worksheet.protect(protection, {});

    if(logo) {
        startingLine = 8;
        const image = workbook.addImage({ base64: logo, extension: 'png' });
        worksheet.addImage(image, 'A1:B3');
    }

    if(campaign) worksheet.getCell('A5').value = campaign;
    if(situation) worksheet.getCell('A6').value = situation;
    if(title) worksheet.getCell(`${titlePositionning}7`).value = title;

    for(let i=startingLine; i < data.length; i++) {
        const line = data[i];

        for(let j=0; j < line.length; j++) {
            let finalCellNumber = null;
            const cell = line[j];
            const cellNumber = `${excelColumns(j+1)}${i+1}`;

            if(cell.rowSpan || cell.colSpan) {
                finalCellNumber = afterMerge(data, cellNumber, j+1, cell.rowSpan, cell.colSpan, cell.text, i, j);
                worksheet.mergeCells(`${cellNumber}`, `${finalCellNumber}`);
            }

            worksheet.getCell(cellNumber).value = cell.text;

            worksheet.getCell(cellNumber).font = {
                name: 'Calibri',
                family: 1,
                size: 14,
            };

            worksheet.getCell(cellNumber).border = {
                top: {style:'thin'},
                left: {style:'thin'},
                bottom: {style:'thin'},
                right: {style:'thin'}
            };

            worksheet.getCell(cellNumber).alignment = {
                wrapText: true,
                shrinkToFit: false,
                vertical: 'middle',
                horizontal: 'center'
            };
        }
    }

    return workbook;
}

/**
 * Function to build the Excel file
 * @param {Workbook} workbook Empty workbook
 * @param {SheetDefaultOptions} worksheetOptions Worksheet Global Default Options
 * @returns {Workbook} A Workbook containing the provided data
 */
export default async (workbook:Workbook, sheetData:IPayload, options:IDefaultOptions) => {
    /* METADATA */
    workbook.creator = ''; //
    workbook.lastModifiedBy = '';
    workbook.created = new Date();
    workbook.modified = new Date();
    workbook.lastPrinted = new Date();

    let finalWorkbook = workbook;

    if(Array.isArray(sheetData.data)) {
        // The data is a list of sheets
        if(isISheetData(sheetData.data[0])) {
            let {title, campaign, situation, logo, data} = sheetData;
            let dataCasted = data as ISheetData[];

            let loopWorkbook = workbook
            dataCasted.forEach(({sheetName, sheetData}) => {
                sheetBuilding(loopWorkbook, title, campaign, situation, logo, sheetData, options, sheetName).then(wb => {
                    loopWorkbook = wb;
                })
            })
            finalWorkbook = loopWorkbook;
        }

        // The data is a single table content definition
        else {
            if(isICell(sheetData.data[0][0])) {
                /* BUILDING PROCESS */
                let startingLine = 0;

                let titlePositionning = excelColumns(Math.round((sheetData.data[0].length) / 2));
                const {protection, defaultOptions = {defaultColWidth: 20}} = options;

                const worksheet = workbook.addWorksheet('Feuille Excel', { properties: defaultOptions});

                if(protection) await worksheet.protect(protection, {});

                let {title, campaign, situation, logo, data} = sheetData;

                let dataCasted = data as ICell[][];

                if(logo) {
                    startingLine = 8;
                    const image = workbook.addImage({ base64: logo, extension: 'png' });
                    worksheet.addImage(image, 'A1:B3');
                }

                if(campaign) worksheet.getCell('A5').value = campaign;
                if(situation) worksheet.getCell('A6').value = situation;
                if(title) worksheet.getCell(`${titlePositionning}7`).value = title;

                for(let i=startingLine; i < dataCasted.length; i++) {
                    const line = dataCasted[i];

                    for(let j=0; j < line.length; j++) {
                        let finalCellNumber = null;
                        const cell = line[j];
                        const cellNumber = `${excelColumns(j+1)}${i+1}`;

                        if(cell.rowSpan || cell.colSpan) {
                            finalCellNumber = afterMerge(dataCasted, cellNumber, j+1, cell.rowSpan, cell.colSpan, cell.text, i, j);
                            worksheet.mergeCells(`${cellNumber}`, `${finalCellNumber}`);
                        }

                        worksheet.getCell(cellNumber).value = cell.text;

                        worksheet.getCell(cellNumber).font = {
                            name: 'Calibri',
                            family: 1,
                            size: 14,
                        };

                        worksheet.getCell(cellNumber).border = {
                            top: {style:'thin'},
                            left: {style:'thin'},
                            bottom: {style:'thin'},
                            right: {style:'thin'}
                        };

                        worksheet.getCell(cellNumber).alignment = {
                            wrapText: true,
                            shrinkToFit: false,
                            vertical: 'middle',
                            horizontal: 'center'
                        };
                    }
                }
            }
            finalWorkbook = workbook;
        }
    }

    return finalWorkbook;
}