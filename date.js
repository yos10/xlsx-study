'use strict';

const XlsxPopulate = require('xlsx-populate');

// 空のワークブックを作成する
XlsxPopulate.fromBlankAsync().then(workbook => {
  // A1 セルに日付を設定
  workbook.sheet(0).cell('A1').value(new Date(2020, 8, 22)).style('numberFormat', 'yyyy年 m月 dd日');

  // A2 セルに算出した日付を設定
  const date = XlsxPopulate.numberToDate(42788);  // Wed Feb 22 2017 00:00:00: GMT-0500
  workbook.sheet(0).cell('A2').value(date).style('numberFormat', 'dddd, mmmm dd, yyyy');

  // シートの微調節
  workbook.sheet(0).column('A').width(30);

  // Excel ファイルの書き出し
  return workbook.toFileAsync('./date.xlsx');
});