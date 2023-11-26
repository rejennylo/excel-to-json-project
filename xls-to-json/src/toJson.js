const node_xj = require('xls-to-json');
const fs = require('fs');

node_xj(
  {
    input: './src/assets/總統-A05-1-候選人得票數一覽表(中央).xls',
    output: './src/output.json',
    sheet: '中 央', // specific sheetname
    rowsToSkip: 1, // number of rows to skip at the top of the sheet; defaults to 0
    allowEmptyKey: false, // avoids empty keys in the output, example: {"": "something"}; default: true
  },
  function (err, result) {
    if (err) {
      console.error(err);
    } else {
      const cleanedResult = result.map((row) => {
        return row;
      });
      // 將處理後的結果寫入到JSON檔案
      fs.writeFile(
        './src/output.json',
        JSON.stringify(cleanedResult, null, 2),
        (writeErr) => {
          if (writeErr) {
            console.error('寫入檔案時發生錯誤:', writeErr);
          } else {
            console.log('JSON檔案已成功寫入.');
          }
        }
      );
    }
  }
);
