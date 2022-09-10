
const path = require('path');
const ExcelJS = require('../lib/exceljs.nodejs');

const TEST_XLSX_FILE_NAME = path.resolve(__dirname, '../data/hyperlink.xlsx');
const filenameOut = path.resolve(__dirname, '../data/hyperlink-out.xlsx');
const options = {
  filename: path.resolve(__dirname, '../data/hyperlink-stream.xlsx'),
  useStyles: true,
};

const wb = new ExcelJS.stream.xlsx.WorkbookWriter(options);
const ws = wb.addWorksheet('Sheet1');
ws.getCell('A1').value = {
  hyperlink: 'https://www.npmjs.com/package/exceljs',
  text: 'ExcelJS',
  tooltip: 'https://www.npmjs.com/package/exceljs',
};
ws.commit();
wb.commit();
// ws.getCell('B1').value = {
//   location: 'Sheet1!A1',
//   text: 'Sheet1',
//   tooltip: 'Go To Sheet1',
// };
// ws.getCell('B1').value = {
//   location: 'Sheet2!A1',
//   richText: [{
//                 text: 'TTTTTest te',
//                 font: {
//                     underline: true,
//                 },
//             }],
//   // text: 'Sheet1',
//   tooltip: 'TTTTTest',
// };
// ws.getCell('C1').style = {
//   font: {
//     underlin: true,
//   },
// };


const wb2 = new ExcelJS.Workbook();
wb2.xlsx.readFile(TEST_XLSX_FILE_NAME).then(() => {

  const ws2 = wb2.getWorksheet('Sheet2');
  ws2.getCell('B1').value = {
    location: 'Sheet1!A1',
    text: 'Sheet1',
    tooltip: 'Go To Sheet1',
  };
  ws2.getCell('C1').value = {
    location: 'Sheet2!A1',
    richText: [{
      text: 'TTTTTest te',
      font: {
        underline: true,
      },
    }],
    // text: 'Sheet1',
    tooltip: 'TTTTTest',
  };
  ws2.getCell('D1').style = {
    font: {
      underline: true,
    },
  };

  wb2.xlsx.writeFile(filenameOut);
});