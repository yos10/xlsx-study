'use strict';

const XlsxPopulate = require('xlsx-populate');

XlsxPopulate.fromBlankAsync().then(workbook => {  // 空のワークブックを作成する
  const sheet2 = workbook.addSheet('Sheet2');  // シート2 を作成
  workbook
    .sheet(0)
      .cell('A1')
        .value('foo')
        .style('bold', true)
      .relativeCell(1, 0)
        .formula('A1')
        .style('italic', true)
  .workbook()
    .sheet(1)
      .range('A1:B3')
        .value(5)
      .cell(0, 0)
        .style('underline', 'double');

  return workbook.toFileAsync('./chain.xlsx');
});