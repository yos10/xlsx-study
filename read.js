'use strict';

const XlsxPopulate = require('xlsx-populate');

// 既存の Book1.xlsx ワークブックを読み込む
XlsxPopulate.fromFileAsync('./Book1.xlsx').then(workbook => {
  // Sheet1 の A1 セルの値を取得する
  const value = workbook.sheet('Sheet1').cell('A1').value()
  
  // 取得した値をログに出力する
  console.log(value);
});