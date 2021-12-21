import { Workbook } from 'exceljs';
import { IPayload, ICell } from './interfaces/IPayload.interface';
import { SheetDefaultOptions } from './types/sheetOptions.type';

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
        payload[lineIndex + (rowSpan - 1)][columnIndex] = {...payload[lineIndex + (rowSpan - 1)][columnIndex], text: cellText};
    }
    if(colSpan) {
        stringPart = excelColumns(letterColumnIndex + (colSpan - 1));
        payload[lineIndex][columnIndex + (colSpan - 1)] = {...payload[lineIndex][columnIndex + (colSpan - 1)], text: cellText}
    }

    return `${stringPart}${numberPart};`
}

/**
 * Fonction de contruction du fichier a partir des operation de base
 * @param {ExcelJS.Workbook} workbook Le classeur vide
 * @param {any} data Les donnees a manipuler pour construire le fichier Excel
 * @param {Object} worksheetOptions Worksheet Global Options
 * @returns {ExcelJS.Workbook} La fonction retourne le classeur contenant les lignes et colonnes apres construction
 */
export default (workbook:Workbook, sheetData:IPayload, worksheetOptions:SheetDefaultOptions) => {

    /* META DONNEES DU CLASSEUR */
    workbook.creator = ''; //
    workbook.lastModifiedBy = '';
    workbook.created = new Date();
    workbook.modified = new Date();
    workbook.lastPrinted = new Date();

    /* CONSTRUCTION DU CONTENU DU CLASSEUR */
    const worksheet = workbook.addWorksheet('My New Sheet', { properties: worksheetOptions}); // Ajout d'une feuille au classeur

    // ============= LARGE TABLE ====================//
    const {title, campaign, situation, logo, data} = sheetData;

    for(let i=0; i < data.length; i++) {
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