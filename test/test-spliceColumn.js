
const path = require('path');
const ExcelJS = require('../lib/exceljs.nodejs');

const TEST_XLSX_FILE_NAME = path.resolve(__dirname, '../data/splice.xlsx');
const filenameOut = path.resolve(__dirname, '../data/splice-out.xlsx');
const options = {
    filename: TEST_XLSX_FILE_NAME,
    useStyles: true,
};
const wb = new ExcelJS.Workbook(options);

wb.xlsx.readFile(TEST_XLSX_FILE_NAME).then(async () => {

    const ws = wb.getWorksheet('Sheet1');
    // ws.spliceColumns(4, 2);
    ws.spliceColumns(7, 1);
    ws.spliceColumns(6, 1);

    ws.spliceColumns(6, 0, [undefined, 1, 2, 3]);

    await wb.xlsx.writeFile(filenameOut);
});