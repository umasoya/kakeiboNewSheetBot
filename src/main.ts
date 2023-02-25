import dayjs = require('dayjs');
const ja = require('dayjs/locale/ja')
const customParseFormat = require('dayjs/plugin/customParseFormat');
dayjs.locale(ja);
dayjs.extend(customParseFormat);

export const newSheet = (auto: boolean = false) => {
    const spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('spread_sheet_id')!);
    const template: GoogleAppsScript.Spreadsheet.Sheet = spreadSheet.getSheetByName('template')!;

    const str: string = Browser.inputBox("シート作成", "シートの年月をyyyymm形式で入力してください\\n例: 2023/01の場合 → 202301 (空の場合はデフォルトで翌月分を作成します)", Browser.Buttons.OK_CANCEL);
    if (str === 'cancel') {
        return;
    }

    // トリガー経由か空入力なら翌月分
    if (auto || str === '') {
        return newNextSheet();
    }

    // 入力値あり
    if (str.match(/\d{6}/) === null) {
        Browser.msgBox('入力形式が間違っています');
        return;
    }

    // シート名の重複
    if (spreadSheet.getSheetByName(str) !== null) {
        Browser.msgBox('すでに同名のシートが存在します');
        return;
    }

    // シートコピー
    const copySheet = template.copyTo(spreadSheet);
    copySheet.setName(str);

    const inputDay: dayjs.Dayjs = dayjs(str, 'YYYYMM');

    // B2を変更(ただの年表記)
    copySheet.getRange(2, 2).setValue(`${inputDay.year()}年`);
    // B3を変更(日付計算の起点)
    copySheet.getRange(3, 2).setValue(`${inputDay.date(1).format('YYYY/MM/DD')}`);

    checkEndOfMonth(copySheet);
};

// 翌月分のシートを生成
export const newNextSheet = () => {
    const spreadSheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('spread_sheet_id')!);
    const template: GoogleAppsScript.Spreadsheet.Sheet = spreadSheet.getSheetByName('template')!;
    const nextMonth: dayjs.Dayjs = dayjs().add(1, 'month').date(1); // 翌月1日

    const year: string = String(nextMonth.year());
    const month: string = String(nextMonth.month() + 1).padStart(2, '0');

    // シート名の重複
    if (spreadSheet.getSheetByName(`${year}${month}`) !== null) {
        Browser.msgBox('すでに同名のシートが存在します');
        return;
    }

    // シートコピー
    const copySheet = template.copyTo(spreadSheet);
    copySheet.setName(`${year}${month}`);

    // B2を変更(ただの年表記)
    copySheet.getRange(2, 2).setValue(`${year}年`);
    // B3を変更(日付計算の起点)
    copySheet.getRange(3, 2).setValue(`${nextMonth.format('YYYY/MM/DD')}`);

    checkEndOfMonth(copySheet);

    sendMessage(`シート「${year}${month}」を作成しました`);
};

// 29~31日(B31~B33)がその月に存在するか確認、なければ行データを削除
const checkEndOfMonth = (sheet: GoogleAppsScript.Spreadsheet.Sheet) => {
    const day1: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(3, 2);
    for (let row of [33, 32, 31]) {
        let targetWeek: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(row, 1);
        let targetDay: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(row, 2);
        if (dayjs(day1.getDisplayValue(), 'MM月DD日').month() === dayjs(targetDay.getDisplayValue(), 'MM月DD日').month()) {
            break;
        }
        targetWeek.clearContent();
        targetDay.clearContent();
    }
};

//  ブロードキャストメッセージで通知する
const sendMessage = (message :string) => {
    const url: string = 'https://api.line.me/v2/bot/message/broadcast';
    const channelAccessToken: string = PropertiesService.getScriptProperties().getProperty('channel_access_token')!;
    const payload = {
        'messages': [
            {
                'type': 'text',
                'text': message,
            }
        ],
        'notificationDisabled': true,
    };
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${channelAccessToken}`,
        },
        method: 'post',
        payload: JSON.stringify(payload),
    };
    UrlFetchApp.fetch(url, options);
};