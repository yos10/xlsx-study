'use strict';

const XlsxPopulate = require('xlsx-populate');

// 空のワークブックを作成する
XlsxPopulate.fromBlankAsync().then(workbook => {
  // `Sheet5` という名前のシートを最後に追加する
  const sheet5 = workbook.addSheet('Sheet5'); // -> Sheet1 Sheet5

  // `Sheet2` という名前のシートを 1番目 に追加する。 (番号は 0 始まり)
  const sheet2 = workbook.addSheet('Sheet2', 1); // -> Sheet1 Sheet2 Sheet5

  // `Sheet3` という名前のシートを `Sheet5` の前に追加する
  const sheet3 = workbook.addSheet('Sheet3', 'Sheet5'); // -> Sheet1 Sheet2 Sheet3 Sheet5
  
  // シートオブジェクト `Sheet5` の変数がすでにある場合は、この変数を使って `Sheet5` の前に `Sheet4` を追加することもできます
  // `Sheet4` という名前のシートを `Sheet5` の前に追加する
  const sheet4 = workbook.addSheet('Sheet4', sheet5); // -> Sheet1 Sheet2 Sheet3 Sheet4 Sheet5

  // シートの移動
  // `Sheet1` という名前のシートを最後に移動する
  workbook.moveSheet('Sheet1');

  // `Sheet1` という名前のシートを 2番 (`0` 始まり)に移動する
  workbook.moveSheet('Sheet1', 2);

  // `Sheet1` という名前のシートを `Sheet2` の前に移動する
  workbook.moveSheet('Sheet1', 'Sheet2');

  // シートの名前変更
  // 先頭 (`0` 番目) のシートの名前を変更する
  // const sheet = workbook.sheet(0).name('new sheet name');

  // シートの削除
  // `Sheet1` という名前のシートを削除する
  workbook.deleteSheet('Sheet1');

  // 2番目 (`0` はじまり) のシートを削除する
  workbook.deleteSheet(2);

  // sheet オブジェクトがあれば `delete 関数` を実行して削除することもできます
  workbook.sheet(0).delete();

  // Excel ファイルの書き出し
  return workbook.toFileAsync('./sheetTest.xlsx');
})