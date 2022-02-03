<h1 align="center" style="text-align: center;">PDFMake-To-Excel</h1>

<p align="center">
  <a href="mailto:scanf555@gmail.com">
    <img src="https://img.shields.io/badge/Ask%20me-anything-1abc9c.svg" alt="NPM package" />
  </a>

  <a href="#">
      <img src="https://img.shields.io/npm/dt/pdfmake-to-excel" alt="NPM package" />
  </a>

  <a href="https://github.com/scanf6/pdfmake-to-excel/">
      <img src="https://img.shields.io/badge/GitHub-100000?&logo=github&logoColor=white" alt="NPM package" />
  </a>

  <a href="https://www.buymeacoffee.com/scanf6">
      <img src="https://img.shields.io/badge/-buy_me_a%C2%A0coffee-gray?logo=buy-me-a-coffee" />
  </a>
</p>

<p align="center">
  <b>Excel generation from Pdfmake tables</b></br>
  <sub>Made by <a href="https://patrickhermann.netlify.app">Patrick</a> </sub>
</p>

# Main Purpose

The purpose of this package was to easily generate an Excel file from the [Pdfmake library](http://pdfmake.org/#/).
Therefore, by providing the payload used to generate PDFs using pdfmake, you should be able to get an Excel file,
without malformed cells or incorrectly structured cols and rows.

Here is the documentation to build pdfmake payloads: [PDFMake Playground](http://pdfmake.org/playground.html)

**This library don't require pdfmake.**

Here is what this library brings:

- Merge rows and cols
- Sheet protection by password
- Front-end excel download
- Server-side excel download by streaming

## ❯ Installation

Install pdfmake-to-excel with npm

```bash
  npm install pdfmake-to-excel
```

## ❯ Usage/Examples

Import the ExcelConverter class from pdfmake-to-excel, instanciate it with the name of your excel file and your pdfmake
table content definition object, then call the downloadExcel() method.

Pass the following arguments to the constructor

- The Excel filename
- The PDFMake table content definition object
- An optional configuration object including
    * A sheet protection password
    * A default Options Excel configuration

```javascript
import {ExcelConverter} from 'pdfmake-to-excel';

function downloadFile() {
    const exporter = new ExcelConverter(
        'Export test',
        contentDefinition,
        {
            protection: 'p@ssw0rd',
            defaultOptions: {defaultColWidth: 20}
        }
    );
    exporter.downloadExcel();
}
```

## ❯ Content Definition Object Format

Here is the documentation to build pdfmake payloads: [PDFMake Playground](http://pdfmake.org/playground.html). Here is
how you should format your table content definition object

```json
{
  "title": "Title displayed on your Excel file", //OPTIONAL
  "logo": "base64 of your image here" //OPTIONAL
  "data": [
    [ // LINE 01
      {
        "text": "Cell 01",  // CELL 01 spanned accross 2 rows
        "rowSpan": 2
      },
      {
        "text": "Cell 02", // CELL 02 Spanned accross 2 cells
        "colSpan": 2
      },
      {
        "text": ""
      }
    ],
    [ // Empty line from the first line rowSpan
      {
        "text": "" // Empty cell from the first line rowSpan
      }
    ]
  ]
}
```

## ❯ Multiple Sheets

To generate an Excel file with multiple sheets and a table on each sheet, all you have to do is to provide the
ExcelConverter Class with a content definition object where the data attribute is an array of sheets, each sheets being
an object with the name (sheetName property) and the table content definition's "data" (See content definition object
format) property (sheetData property)

```javascript
const exporter = new ExcelConverter(
    'File_name',
    {
        data: [
            {sheetName: 'Sheet_name 01', sheetData: contentDefinitionData1},
            {sheetName: 'Sheet_name 02', sheetData: contentDefinitionData2},
            {sheetName: 'Sheet_name 03', sheetData: contentDefinitionData3},
        ]
    }
);
```

## ❯ Streaming

In case you want to export your Excel file server-side, pdfmake-to-excel provides the getStream() method which takes in
an optionnal argument which is your response as the first argument.

- When the response argument is provided, the excel file is created and directly piped to your response.
- When the response argument is not provided, you'll get the stream itself. Up to you to pipe it wherever you want

## ❯ Example Using NestJS

```javascript
@Get('/export-excel-file')
async exportReportExcel(@Res() response:Response):Promise < any > {
    const exporter = new ExcelConverter('FileTest', contentDefinition);

    // Automatic pipe to response
    await exporter.getStream(response);


    // Get the stream
    const stream = await exporter.getStream();

    // Then pipe it if you want
    stream.pipe(response);
}
```

## ❯ Example Using AdonisJS

```javascript
const ExportService = use('App/Services/ExportService');

const {ExcelConverter} = require('pdfmake-to-excel')

class ExportController{
  async exportAction({request, response}){

    const executor = ExportService(request.all());
    const excelConverter = new ExcelConverter('test-filename.png', await executor.getData());

    response.implicitEnd = false;
    response.header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    response.header('Content-Transfer-Encoding', 'binary');
    response.header('Content-Disposition', `attachment; filename="fichier.xlsx"`)

    excelConverter.getStream().pipe(response.response);
  }
}
module.exports = ExportController;
```