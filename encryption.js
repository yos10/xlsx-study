'use strict';

const XlsxPopulate = require('xlsx-populate');

// 空のワークブックを作成する
XlsxPopulate.fromBlankAsync().then(workbook => {
  workbook.sheet(0).cell('A1').value('機密情報');  // A1 セルに文章を書き込む

  // パスワード保護された Excel の書き出し
  workbook.toFileAsync('./encryption.xlsx', { password: 'S3cret!' });  // パスワードを S3cret! に設定  
});



// パスワードで保護された Excel の開き方
// XlsxPopulate.fromFileAsync('./encryption.xlsx', { password: 'S3cret!' }).then(workbook => {
//   // なにか処理
// });