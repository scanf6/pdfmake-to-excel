
# PDFMake Table to Excel

pdfmake-to-excel is a package that helps build an Excel file from a table content definition object from the pdfmake library.


## Installation

Install pdfmake-to-excel with npm

```bash
  npm install pdfmake-to-excel
```

## Usage/Examples
Import the ExcelConverter class from pdfmake-to-excel, instanciate it with the name of your excel file and your pdfmake table content definition object, then call the downloadExcel() method.

Pass the following arguments to the constructor

- The Excel filename
- The PDFMake table content definition object
- A configuration object including
    * A sheet protection password [OPTIONAL]
    * A default Options Excel configuration [OPTIONAL]
```javascript
import {ExcelConverter} from 'pdfmake-to-excel';

function downloadFile() {
    const exporter = new ExcelConverter(
        'Export test',
        contentDefinition,
        {
            protection?: 'p@ssw0rd',
            defaultOptions?: {defaultColWidth: 20}
        }
    );
    exporter.downloadExcel();
}
```


## Content Definition Object Format
Here is how you should format your table content definition object

```javascript
{
    "title": "Title displayed on your Excel file", //OPTIONAL
    "logo": "base64 of your image here" //OPTIONAL
    "data": [
        [ // LINE 01
            {
                "text": "Cell 01", // CELL 01 spanned accross 2 rows
                "rowSpan": 2
            },
            {
                "text": "Cell 02", // CELL 02 Spanned accross 2 cells
                "colSpan": 2
            },
            {
                "text": ""
            },
        ],
        [ // Empty line from the first line rowSpan
            {
                "text": "" // Empty cell from the first line rowSpan
            }
        ]
    ]
}
```

## Multiple Sheets
To generate an Excel file with multiple sheets and a table on each sheet, all you have to do is to provide the ExcelConverter Class with a content definition object where the data attribute is an array of sheets, each sheets being an object with the name (sheetName property) and the table content definition (sheetData property)
```
const exporter = new ExcelConverter(
  'File_name',
  {
    data: [
      { sheetName: 'Sheet_name 01', sheetData: contentDefinition1 },
      { sheetName: 'Sheet_name 02', sheetData: contentDefinition2 },
      { sheetName: 'Sheet_name 03', sheetData: contentDefinition3 },
    ]
  }
);
```