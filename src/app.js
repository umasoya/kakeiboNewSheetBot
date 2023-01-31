const newSheet = (auto = false) => {
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const originSheet = spreadSheet.getActiveSheet();

  const result = Browser.inputBox("シート作成", "シートの年月をyyyymm形式で入力してください\\n例: 2023/01の場合 → 202301 (空の場合はデフォルトで翌月分を作成します)", Browser.Buttons.OK_CANCEL);
  if (result === 'cancel') {
    return;
  }
  let ym;

  if (result === '') {
    // 空入力
    const now = new Date();
    const year = String(now.getFullYear()).padStart(4, '0');
    const month = String(now.getMonth() + 2).padStart(2, '0');
    ym = `${year}${month}`;
  } else {
    // 入力値あり
    if (result.match(/\d{6}/) === null) {
      Browser.msgBox('入力形式が間違っています');
      return;
    }
    ym = result;
  }

  // シート名の重複
  if (spreadSheet.getSheetByName(ym) !== null) {
    Browser.msgBox('すでに同名のシートが存在します');
    return;
  }

  // シートコピー
  const copySheet = originSheet.copyTo(spreadSheet);
  copySheet.setName(ym);

  // B2を変更(ただの年表記)
  copySheet.getRange(2,2).setValue(`${ym.substring(0,4)}`);
  // B3を変更(日付計算の起点)
  copySheet.getRange(3,2).setValue(`${ym.substring(0,4)}/${ym.substring(4,6)}/01`);
};