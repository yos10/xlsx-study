'use strict';

const XlsxPopulate = require('xlsx-populate');

// 空のワークブックを作成する
XlsxPopulate.fromBlankAsync().then(workbook => {
  const sheet = workbook.sheet('Sheet1');

  // A1 セルに文章を入れる
  sheet.cell('A1').value('太字');
  // A1 セルに bold (太字) スタイルをあてる
  sheet.cell('A1').style('bold', true);

  // B1 セルに italic (斜体) スタイルをあてる
  sheet.cell('B1').value('イタリック').style('italic', true);

  // C1 セルに bold (太字) と italic (斜体) の2つのスタイルを同時にあてる
  sheet.cell('C1').value('両方').style({'italic': true, 'bold': true});

  // D1 セルに数値フォーマットをあてる
  sheet.cell('D1').value(1234.56);
  sheet.cell('D1').style('numberFormat', '0.00');

  // 範囲で指定することもできます
  // A2 セルに文章を入れる
  sheet.cell('A2').value('水色の背景');
  // A2~E2 セルをまとめて背景水色のスタイルをあてる
  sheet.range('A2:E2').style('fill', '00ffff');

  // A3 セルに文章を入れる
  sheet.cell('A3').value('ランダムな背景');
  // B3~F3 セルにランダムな色の背景のスタイルをあてる
  const Hex = '0123456789abcdef';
  sheet.range('B3:F3').style({
    fill: (cell, ri, ci, range) => {
      let rgb = '';
      for (let i = 0; i < 6; i++) {
        rgb += Hex[Math.floor(Math.random() * Hex.length)];
      }
      return rgb;
    }
  });

  // 行や列でも指定できます
  // 4行目に枠線のスタイルをあてる
  sheet.row(4).style('border', true);
  // C列 に中央寄せのスタイルをあてる
  sheet.column('C').style('horizontalAlignment', 'center');
  sheet.cell('C5').value('中央寄せ');

  // 複雑なパラメーターの設定
  // A6 セルに文章を入れる
  sheet.cell('A6').value('複雑な背景→');
  // B6 セルに複雑なパラメーターの背景スタイルをあててみる
  sheet.cell('B6').style('fill', {
    type: 'pattern',
    pattern: 'darkDown',
    foreground: {
      rgb: 'ff0000'
    },
    background: {
      theme: 3,
      tint: 0.4
    }
  });

  // シートの微調節
  sheet.column('A').width(15);
  // Excel ファイルの書き出し
  return workbook.toFileAsync('./styles.xlsx');
});