
# PDFMake to Excel

pdfmake-to-excel is a package that helps build an Excel file from a content definition object from the pdfmake library.


## Installation

Install pdfmake-to-excel with npm

```bash
  npm install pdfmake-to-excel
```

## Usage/Examples
Import the ExcelConverter class from pdfmake-to-excel, instanciate it with the name of your excel file and your content definition object, then call the downloadExcel() method.
```javascript
import {ExcelConverter} from 'pdfmake-to-excel';

function downloadFile() {
    const exporter = new ExcelConverter('Export test', contentDefinition);
    exporter.downloadExcel();
}
```


## Content Definition Object Format
Here is how you should format your content definition object

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