'use strict';

const XlsxPopulate = require('xlsx-populate');

// 空のワークブックを作成する
XlsxPopulate.fromBlankAsync().then(workbook => {
  // ワークブックの Sheet1 の A1 セルに文章を入れる
  workbook.sheet('Sheet1').cell('A1').value('新しく作った Excel ');
  workbook.sheet('Sheet1').column('A').width(25);
  workbook.sheet('Sheet1').row(1).height(30);

  // Excel ファイルの書き出し
  return workbook.toFileAsync('./Book1.xlsx');
});