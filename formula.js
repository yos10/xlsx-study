'use strict';

const XlsxPopulate = require('xlsx-populate');

// 空のワークブックを作成する
XlsxPopulate.fromBlankAsync().then(workbook => {
  // 簡単な足し算の例
  workbook.sheet(0).cell('A1').value(1);  // 先頭のシートの A1 セルに 1 を設定する
  workbook.sheet(0).cell('A2').formula('=A1+2');  // 先頭のシートの A2 セルに `A1+2` の式を設定する

  // 異なるセルの値を参照する例
  const sheet2 = workbook.addSheet('Sheet2');  // シート2 を作成
  sheet2.cell('A1').value(9);  // シート2 の A1 セルに `9` を設定する
  workbook.sheet(0).cell('B1').formula('=Sheet2!A1&"*3は "&Sheet2!A1*3');  // 取得した値に 3 を掛け算する

  // 範囲にまとめて式を設定する例
  workbook.sheet(0).cell('C1').value(1);  // 先頭のシートの C1 セルに 1 を設定する
  workbook.sheet(0).range('C2:C11').formula("=INDIRECT(ADDRESS(ROW()-1,COLUMN()))*2");  // 自分自身の1つ上のセルの `2倍` を計算する式を `C2~C11` に設定する

  // Excel ファイルの書き出し
  return workbook.toFileAsync('./formula.xlsx');
});