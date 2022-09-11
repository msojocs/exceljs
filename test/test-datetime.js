
const path = require('path');
const ExcelJS = require('../lib/exceljs.nodejs');

const TEST_XLSX_FILE_NAME = path.resolve(__dirname, './data/datetime.xlsx');
const filenameOut = path.resolve(__dirname, './data/datetime-out.xlsx');
const options = {
    filename: TEST_XLSX_FILE_NAME,
    useStyles: true,
};
const wb = new ExcelJS.Workbook(options);

wb.xlsx.readFile(TEST_XLSX_FILE_NAME).then(async () => {

    const ws = wb.getWorksheet('Sheet1');
    const a1 = ws.getCell('A1');
    console.log(a1);
    const a3 = ws.getCell('A3');
    console.log(a3);

    await wb.xlsx.writeFile(filenameOut);
});