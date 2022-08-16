
const path = require('path');
const Excel = require('../lib/exceljs.nodejs.js');

const HrStopwatch = require('./utils/hr-stopwatch');

const filename = path.resolve(__dirname, 'data/hyperlink.xlsx');
const filenameOut = path.resolve(__dirname, 'data/hyperlink-out.xlsx');

const test = async () =>{

  const wb = new Excel.Workbook();
  
  await wb.xlsx.readFile(filename);
  const ws = wb.addWorksheet('Foo');
  
  ws.getCell('A1').value = {
    hyperlink: 'https://www.npmjs.com/package/exceljs',
    text: 'ExcelJS',
    tooltip: 'https://www.npmjs.com/package/exceljs',
  };
  ws.getCell('B1').value = {
    location: 'Sheet1!A1',
    text: 'Sheet1',
    tooltip: 'Go To Sheet1',
  };
  const ws1 = wb.getWorksheet('Sheet1');
  console.log('================over===========');
  console.log(ws1.getCell(1,1).isHyperlink, ws1.getCell(1,1).hyperlink);
  console.log(ws1.getCell(1,1).value);
  console.log(ws1.getCell(1,2).isHyperlink, ws1.getCell(1,2).value);
  console.log(ws1.getCell(1,2).model);
  console.log(ws1.hyperlinks);
  wb.xlsx.writeFile(filenameOut);
 
  
};
test();