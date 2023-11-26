const ExcelJS = require('exceljs');
const fs = require('fs').promises;

async function convertExcelToJson() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile('./src/assets/總統-A05-1-候選人得票數一覽表(中央).xlsx');
  const worksheet = workbook.getWorksheet(1); // 取用第一張表單
  const json = [];

  worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
    const rowValue = row.values;
    // 從指定 row 開始取資料
    if (rowNumber > 4) {
      const id = rowValue[1].toString().replace(/\s+/g, ''); // 取得標識符所在的 row ，並刪除空白字元
      json.push({ id: id, data: rowValue.slice(2) }); // 忽略第一筆資料，剩餘的放入 array
    }
  });

  await fs.writeFile('./src/output.json', JSON.stringify(json, null, 2), 'utf8');
  console.log('Excel file has been converted to JSON');
}

convertExcelToJson().catch((error) => console.error(error));
