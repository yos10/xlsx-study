'use strict';

const XlsxPopulate = require('xlsx-populate');

// 空のワークブックを作成する
XlsxPopulate.fromBlankAsync().then(workbook => {
  // 特定のセルに書き込む
  workbook.sheet(0).cell('A1').value('得点表');
  workbook.sheet(0).range('B1:D1').value(
    [
      ['英語', '国語', '数学']
    ]
  );

  workbook.sheet(0).cell('A2').value(
    [
      ['a くん'],
      ['b くん'],
      ['c くん'],
      ['d くん'],
      ['e くん'],
    ]
  );

  const range = workbook.sheet(0).range('B2:D6');
  range.value((cell, ri, ci, range) => Math.floor(Math.random() * 101));

  // Excel ファイルの書き出し
  return workbook.toFileAsync('./points.xlsx');
});