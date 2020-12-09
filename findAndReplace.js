'use strict';

const XlsxPopulate = require('xlsx-populate');

// 既存の得点表の points.xlsx ワークブックを読み込む
XlsxPopulate.fromFileAsync('./points.xlsx').then(workbook => {
  // 一致する文字列を検索しますが置き換えはしません
  workbook.find('得点表');  // 一致したセルを配列で返します

  // 先頭のシートから検索
  workbook.sheet(0).find('得点表');  // 一致したセルを配列で返します

  // 特定のセルが文字列を持っているか調べます
  workbook.sheet('Sheet1').cell('A1').find('得点表');  // true もしくは false を返します

  // ワークブック全体から '得点表' という文字列を検索し、'点数表' に置き換えます
  workbook.find('得点表', '点数表');  // 一致したセルを配列で返します

  // 正規表現を使って小文字を大文字に置き換えます
  workbook.find(/[a-z]+/g, match => match.toUpperCase());

  // Excel ファイルの書き出し
  return workbook.toFileAsync('./points2.xlsx');
});