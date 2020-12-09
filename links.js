'use strict';

const XlsxPopulate = require('xlsx-populate');

// 空のワークブックを作成する
XlsxPopulate.fromBlankAsync().then(workbook => {
  // リンク先となる シート2 を作成
  const sheet2 = workbook.addSheet('Sheet2');  // シート2 を作成
  sheet2.cell('A1').value('飛び先');  // シート2 の A1 セルにテキストを設定

  // A1 セルにリンクを設定します
  workbook.sheet(0).cell('A1').value('リンクテキスト')
    .style({fontColor: '0563c1', underline: true})
    .hyperlink('https://www.nicovideo.jp/');

  // A2 セルのリンク情報を取得
  const value = workbook.sheet(0).cell('A2').hyperlink();  // 'https://www.nicovideo.jp/' をかえす
  console.log(value);  // 確認

  // メール送信のリンクを設定
  workbook.sheet(0).cell('A3').value('Click to Email D_drAAgon')
    .hyperlink({email: 'davideryu_orihara@nnn.ed.jp', emailSubject: "大変お忙しい所ではござますが、折原さん..." });
  
  // Set a hyperlink to and internal cell using an address string.
  workbook.sheet(0).cell('A4').value('Click to go to an internal cell')
    .hyperlink('Sheet2!A1')

  // Excell ファイルの書き出し
  return workbook.toFileAsync('./links.xlsx');
})