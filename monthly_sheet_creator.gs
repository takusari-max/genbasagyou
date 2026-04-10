/**
 * 月次報告シート作成スクリプト
 * 
 * 【セットアップ手順】
 * 1. スプレッドシートの「拡張機能」→「Apps Script」を開く
 * 2. このコードを貼り付ける
 * 3. setupMonthlyTrigger を1回だけ実行する（毎月10日の自動実行トリガーが設定される）
 * ※ サービスの追加やGCPの設定は不要です
 */

// ============================================================
// 初期設定：毎月10日のトリガーを登録（1回だけ実行する）
// ============================================================

function setupMonthlyTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // トリガー実行時にスプレッドシートを特定するためIDを保存
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  // 既存の createMonthlySheet トリガーを削除（重複防止）
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'createMonthlySheet') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 毎月10日に実行するトリガーを作成
  ScriptApp.newTrigger('createMonthlySheet')
    .timeBased()
    .onMonthDay(10)
    .atHour(7)
    .create();

  Logger.log('毎月10日 7:00 のトリガーを設定しました（SSID: ' + ss.getId() + '）');
}


// ============================================================
// メイン関数（トリガーから自動実行される）
// ============================================================

function createMonthlySheet() {
  // トリガー実行時は getActiveSpreadsheet が使えないためIDから開く
  var ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  var ss;
  if (ssId) {
    ss = SpreadsheetApp.openById(ssId);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  }

  // ── 1. 最新シート名から次の年月を算出 ──
  var nextYM = getNextYearMonth_(ss);
  if (!nextYM) {
    Logger.log('yyyymm 形式のシートが見つかりません。処理を中止します。');
    return;
  }
  var year  = nextYM.year;
  var month = nextYM.month;
  var sheetName = String(year) + String(month).padStart(2, '0');

  Logger.log('作成対象: ' + sheetName + '（' + year + '年' + month + '月）');

  // 同名シートが既にある場合はスキップ
  if (ss.getSheetByName(sheetName)) {
    Logger.log('シート「' + sheetName + '」は既に存在します。処理をスキップします。');
    return;
  }

  // ── 2. フォーマットシートをコピー ──
  var formatSheet = ss.getSheetByName('フォーマット');
  if (!formatSheet) { Logger.log('「フォーマット」シートが見つかりません。'); return; }
  var reportSheet = ss.getSheetByName('全社報告貼付用');
  if (!reportSheet) { Logger.log('「全社報告貼付用」シートが見つかりません。'); return; }

  var newSheet = formatSheet.copyTo(ss);
  newSheet.setName(sheetName);

  // 全社報告貼付用シートの一つ右に移動
  var targetIndex = reportSheet.getIndex();
  ss.setActiveSheet(newSheet);
  ss.moveActiveSheet(targetIndex + 1);

  // ── 3. A2セルに「yyyy年mm月」を設定 ──
  var a2 = newSheet.getRange('A2');
  a2.setNumberFormat('@');
  a2.setValue(year + '年' + String(month).padStart(2, '0') + '月');
  a2.setFontSize(18);
  a2.setFontWeight('bold');

  // ── 4. カレンダーを作成 ──
  var calInfo = buildCalendar_(newSheet, year, month);

  // ── 5. メールシートを更新 ──
  updateMailSheet_(ss, newSheet, year, month, calInfo);

  // ── 6. 全社報告貼付用シートの数式を更新 ──
  updateReportFormulas_(reportSheet, sheetName);

  // ── 7. メール送信トリガーを設定 ──
  setupMailTrigger_(calInfo);

  Logger.log('シート「' + sheetName + '」の作成が完了しました');
}


// ============================================================
// シート名から次の年月を算出
// ============================================================

function getNextYearMonth_(ss) {
  var sheets = ss.getSheets();
  var latestNum = 0;

  sheets.forEach(function(s) {
    var name = s.getName();
    if (/^\d{6}$/.test(name)) {
      var num = parseInt(name, 10);
      if (num > latestNum) {
        latestNum = num;
      }
    }
  });

  if (latestNum === 0) return null;

  var latestYear  = Math.floor(latestNum / 100);
  var latestMonth = latestNum % 100;

  Logger.log('最新シート: ' + latestNum + '（' + latestYear + '年' + latestMonth + '月）');

  var nextMonth = latestMonth + 1;
  var nextYear  = latestYear;
  if (nextMonth > 12) {
    nextMonth = 1;
    nextYear  = latestYear + 1;
  }

  return { year: nextYear, month: nextMonth };
}


// ============================================================
// カレンダー生成（N列以降）
// ============================================================

function buildCalendar_(sheet, year, month) {
  var START_COL = 14; // N列
  var HEADER_START_ROW = 2;
  var FORMAT_END_ROW = 64;

  // ── 祝日一覧を取得 ──
  var holidays = getJapaneseHolidays_(year, month);

  // ── カレンダー開始日を算出 ──
  var prevMonthLastDay = new Date(year, month - 1, 0);
  var dow = prevMonthLastDay.getDay();
  var mondayOffset = (dow === 0) ? 6 : dow - 1;

  var startDate = new Date(prevMonthLastDay);
  startDate.setDate(startDate.getDate() - mondayOffset);

  // 水曜日が前月に含まれない場合、さらに1週間前から開始
  var wed = new Date(startDate);
  wed.setDate(wed.getDate() + 2);
  if (wed.getMonth() !== prevMonthLastDay.getMonth()) {
    startDate.setDate(startDate.getDate() - 7);
  }

  // ── カレンダー終了日：対象月の最終日まで ──
  var lastDay = new Date(year, month, 0).getDate();
  var endDate = new Date(year, month - 1, lastDay);
  Logger.log('カレンダー範囲: ' + startDate + ' 〜 ' + endDate + '（' + year + '年' + month + '月は' + lastDay + '日まで）');

  // ── 日付配列を作成 ──
  var dates = [];
  var cur = new Date(startDate);
  while (cur <= endDate) {
    dates.push(new Date(cur));
    cur.setDate(cur.getDate() + 1);
  }
  var numDays = dates.length;

  // 月末日の検証
  var actualLastDate = dates[numDays - 1].getDate();
  Logger.log('日付配列の最終日: ' + actualLastDate + '日 / 月の最終日: ' + lastDay + '日 / 総列数: ' + numDays);

  // ── カレンダーの最初の火曜日・最初の水曜日を特定 ──
  var firstTuesday = null;
  var firstWednesday = null;
  for (var d = 0; d < numDays; d++) {
    if (!firstTuesday && dates[d].getDay() === 2)  firstTuesday = new Date(dates[d]);
    if (!firstWednesday && dates[d].getDay() === 3) firstWednesday = new Date(dates[d]);
    if (firstTuesday && firstWednesday) break;
  }

  // ── シートの列数が足りなければ追加 ──
  var requiredCols = START_COL + numDays + 5 - 1;
  var currentCols = sheet.getMaxColumns();
  if (requiredCols > currentCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }

  // ── ヘッダー行（月・日・曜日）を書き込み ──
  var japDays = ['日', '月', '火', '水', '木', '金', '土'];
  var monthRow = [];
  var dayRow   = [];
  var wdayRow  = [];

  for (var i = 0; i < numDays; i++) {
    var m = dates[i].getMonth() + 1;
    var y = (m > month) ? year - 1 : year;
    monthRow.push(y + '年' + m + '月');
    dayRow.push(dates[i].getDate());
    wdayRow.push(japDays[dates[i].getDay()]);
  }

  // セルの書式を設定
  sheet.getRange(2, START_COL, 1, numDays).setNumberFormat('@');
  sheet.getRange(3, START_COL, 1, numDays).setNumberFormat('0');
  sheet.getRange(4, START_COL, 1, numDays).setNumberFormat('@');

  sheet.getRange(2, START_COL, 1, numDays).setValues([monthRow]);
  sheet.getRange(3, START_COL, 1, numDays).setValues([dayRow]);
  sheet.getRange(4, START_COL, 1, numDays).setValues([wdayRow]);

  // ヘッダーの中央揃え
  sheet.getRange(2, START_COL, 3, numDays).setHorizontalAlignment('center');

  // ── 列幅を調整 ──
  for (var i = 0; i < numDays; i++) {
    sheet.setColumnWidth(START_COL + i, 30);
  }

  // ══════════════════════════════════════════════════════════
  // ★ 処理順序が重要：塗りつぶし → 結合 → 罫線
  //   結合前に塗るので、個別セルに確実に色が入る
  //   罫線は最後に引くので上書きされない
  // ══════════════════════════════════════════════════════════

  // ── 1. 塗りつぶし（結合前に実行） ──
  applyColors_(sheet, START_COL, numDays, dates, holidays, HEADER_START_ROW, FORMAT_END_ROW);

  // ── 2. 2行目：同じ月のセルを結合（塗りつぶし後） ──
  mergeMonthCells_(sheet, dates, START_COL);

  // ── 3. 罫線（結合後） ──
  applyAllBorders_(sheet, START_COL, numDays, dates, HEADER_START_ROW, FORMAT_END_ROW);

  // ── 4. 余白として後ろに5列追加 ──
  var lastCalCol = START_COL + numDays - 1;
  for (var i = 1; i <= 5; i++) {
    sheet.setColumnWidth(lastCalCol + i, 30);
  }

  // カレンダー情報を返す
  return {
    firstTuesday: firstTuesday,
    firstWednesday: firstWednesday
  };
}


// ============================================================
// メールシートを更新（C1, C3, C8）
// ============================================================

function updateMailSheet_(ss, newSheet, year, month, calInfo) {
  var mailSheet = ss.getSheetByName('メール');
  if (!mailSheet) {
    Logger.log('「メール」シートが見つかりません。更新をスキップします。');
    return;
  }

  var japDays = ['日', '月', '火', '水', '木', '金', '土'];

  // C1: （yyyy年mm月分）
  mailSheet.getRange('C1').setNumberFormat('@');
  mailSheet.getRange('C1').setValue('（' + year + '年' + month + '月分）');
  Logger.log('メールC1: （' + year + '年' + month + '月分）');

  // C3: （mm月dd日(曜)13:00まで）— カレンダーの最初の火曜日
  if (calInfo.firstTuesday) {
    var t = calInfo.firstTuesday;
    var tMonth = t.getMonth() + 1;
    var tDay   = t.getDate();
    var tWday  = japDays[t.getDay()];
    var c3Value = '（' + tMonth + '月' + tDay + '日(' + tWday + ')13:00まで）';
    mailSheet.getRange('C3').setNumberFormat('@');
    mailSheet.getRange('C3').setValue(c3Value);
    Logger.log('メールC3: ' + c3Value);
  }

  // C8: 新しいシートのURL
  var sheetUrl = ss.getUrl() + '#gid=' + newSheet.getSheetId();
  mailSheet.getRange('C8').setValue(sheetUrl);
  Logger.log('メールC8: ' + sheetUrl);
}


// ============================================================
// ★ 塗りつぶし（結合前に1回だけ実行）
// ============================================================

function applyColors_(sheet, startCol, numDays, dates, holidays, startRow, endRow) {
  var LIGHT_RED  = '#f4cccc';
  var LIGHT_BLUE = '#cfe2f3';
  var WHITE      = '#ffffff';
  var totalRows = endRow - startRow + 1;
  var calRange  = sheet.getRange(startRow, startCol, totalRows, numDays);

  // ── 1. フォーマットシートから引き継がれた書式を除去 ──

  // 条件付き書式をすべて削除（カレンダー範囲と重なるルールを除去）
  var rules = sheet.getConditionalFormatRules();
  if (rules.length > 0) {
    var calStartRow = startRow;
    var calEndRow   = endRow;
    var calStartCol = startCol;
    var calEndCol   = startCol + numDays - 1;
    var newRules = [];
    for (var r = 0; r < rules.length; r++) {
      var ranges = rules[r].getRanges();
      var overlaps = false;
      for (var rr = 0; rr < ranges.length; rr++) {
        var rng = ranges[rr];
        if (rng.getRow() <= calEndRow && rng.getLastRow() >= calStartRow &&
            rng.getColumn() <= calEndCol && rng.getLastColumn() >= calStartCol) {
          overlaps = true;
          break;
        }
      }
      if (!overlaps) {
        newRules.push(rules[r]);
      }
    }
    sheet.setConditionalFormatRules(newRules);
    Logger.log('条件付き書式を ' + (rules.length - newRules.length) + ' 件削除');
  }

  // 交互の背景色（バンディング）を削除
  var bandings = sheet.getBandings();
  for (var b = 0; b < bandings.length; b++) {
    var bRange = bandings[b].getRange();
    if (bRange.getRow() <= endRow && bRange.getLastRow() >= startRow &&
        bRange.getColumn() <= startCol + numDays - 1 && bRange.getLastColumn() >= startCol) {
      bandings[b].remove();
      Logger.log('交互の背景色を削除');
    }
  }

  // ── 2. カレンダー全体を白で初期化（nullではなく白で塗って既存色を完全に上書き） ──
  calRange.setBackground(WHITE);

  // ── 3. 土日祝の列に色を設定（1回だけ、結合前なので個別セルに確実に入る） ──
  for (var i = 0; i < numDays; i++) {
    var d  = dates[i];
    var dw = d.getDay();
    var key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
    var isHoliday = holidays.indexOf(key) !== -1;

    var color = null;
    if (dw === 0 || isHoliday) {
      color = LIGHT_RED;
    } else if (dw === 6) {
      color = LIGHT_BLUE;
    }

    if (color) {
      sheet.getRange(startRow, startCol + i, totalRows, 1).setBackground(color);
    }
  }

  Logger.log('塗りつぶし完了（結合前・1回のみ）');
}


// ============================================================
// 罫線をまとめて適用
// ============================================================

function applyAllBorders_(sheet, startCol, numDays, dates, startRow, endRow) {
  var BLACK  = '#000000';
  var SOLID  = SpreadsheetApp.BorderStyle.SOLID;
  var DASHED = SpreadsheetApp.BorderStyle.DASHED;
  var totalRows = endRow - startRow + 1;

  // 1. 外枠：実線
  sheet.getRange(startRow, startCol, totalRows, numDays)
       .setBorder(true, true, true, true, false, false, BLACK, SOLID);

  // 2. 内部の縦線：破線
  for (var i = 0; i < numDays - 1; i++) {
    sheet.getRange(startRow, startCol + i, totalRows, 1)
         .setBorder(null, null, null, true, null, null, BLACK, DASHED);
  }

  // 3. 月の境目：実線の縦線
  for (var i = 1; i < numDays; i++) {
    if (dates[i].getMonth() !== dates[i - 1].getMonth()) {
      sheet.getRange(startRow, startCol + i, totalRows, 1)
           .setBorder(null, true, null, null, null, null, BLACK, SOLID);
    }
  }

  // 4. ヘッダー行間の横罫線
  sheet.getRange(2, startCol, 1, numDays)
       .setBorder(null, null, true, null, null, null, BLACK, SOLID);
  sheet.getRange(3, startCol, 1, numDays)
       .setBorder(null, null, true, null, null, null, BLACK, DASHED);
  sheet.getRange(4, startCol, 1, numDays)
       .setBorder(null, null, true, null, null, null, BLACK, SOLID);

  // 5. 5行ごとに下線（9, 14, 19, ... 64行目）
  for (var row = 9; row <= endRow; row += 5) {
    sheet.getRange(row, startCol, 1, numDays)
         .setBorder(null, null, true, null, null, null, BLACK, SOLID);
  }
}


// ============================================================
// 2行目の同じ月のセルを結合
// ============================================================

function mergeMonthCells_(sheet, dates, startCol) {
  var i = 0;
  while (i < dates.length) {
    var currentMonth = dates[i].getMonth();
    var j = i + 1;
    while (j < dates.length && dates[j].getMonth() === currentMonth) {
      j++;
    }
    var span = j - i;
    if (span > 1) {
      sheet.getRange(2, startCol + i, 1, span).merge();
    }
    i = j;
  }
}


// ============================================================
// 全社報告貼付用シートの数式内の年月を置換
// ============================================================

function updateReportFormulas_(reportSheet, newSheetName) {
  var lastRow = reportSheet.getLastRow();
  var lastCol = reportSheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return;

  var range = reportSheet.getRange(1, 1, lastRow, lastCol);
  var formulas = range.getFormulas();
  var updated = false;

  var ymPattern = /'?\d{6}'?!/g;

  for (var r = 0; r < formulas.length; r++) {
    for (var c = 0; c < formulas[r].length; c++) {
      if (formulas[r][c]) {
        var original = formulas[r][c];
        var replaced = original.replace(ymPattern, function(match) {
          var clean = match.replace(/[!']/g, '');
          if (/^\d{6}$/.test(clean)) {
            return "'" + newSheetName + "'!";
          }
          return match;
        });

        if (replaced !== original) {
          reportSheet.getRange(r + 1, c + 1).setFormula(replaced);
          updated = true;
        }
      }
    }
  }

  Logger.log(updated
    ? '全社報告貼付用シートの数式を「' + newSheetName + '」に更新しました'
    : '全社報告貼付用シートに更新対象の数式がありませんでした');
}


// ============================================================
// メール送信トリガーを設定
// ============================================================

/**
 * カレンダーの最初の水曜日の1週間前・朝8:00にメール送信トリガーを設定
 */
function setupMailTrigger_(calInfo) {
  if (!calInfo || !calInfo.firstWednesday) {
    Logger.log('最初の水曜日が見つかりません。トリガー設定をスキップします。');
    return;
  }

  // 既存の sendScheduledMail トリガーを削除（重複防止）
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendScheduledMail') {
      ScriptApp.deleteTrigger(triggers[i]);
      Logger.log('既存のメール送信トリガーを削除しました');
    }
  }

  // 最初の水曜日の1週間前
  var sendDate = new Date(calInfo.firstWednesday);
  sendDate.setDate(sendDate.getDate() - 7);
  sendDate.setHours(8, 0, 0, 0);

  // 過去の日付の場合はトリガーを設定しない
  if (sendDate <= new Date()) {
    Logger.log('送信予定日（' + sendDate + '）が過去のため、トリガーは設定しません。手動で sendScheduledMail を実行してください。');
    return;
  }

  ScriptApp.newTrigger('sendScheduledMail')
    .timeBased()
    .at(sendDate)
    .create();

  Logger.log('メール送信トリガーを設定: ' + sendDate);
}



// ============================================================
// メール送信（トリガーから実行される）
// ============================================================

function sendScheduledMail() {
  var ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!ssId) {
    Logger.log('スプレッドシートIDが保存されていません。');
    return;
  }
  var ss = SpreadsheetApp.openById(ssId);

  // ── メール本文を取得 ──
  var mailSheet = ss.getSheetByName('メール');
  if (!mailSheet) {
    Logger.log('「メール」シートが見つかりません。下書き作成をスキップします。');
    return;
  }

  // タイトル: C1〜C4をつなげる
  var subject = '';
  for (var r = 1; r <= 4; r++) {
    subject += String(mailSheet.getRange('C' + r).getDisplayValue());
  }

  // 本文: C5〜C20をつなげる
  var body = '';
  for (var r = 5; r <= 20; r++) {
    var line = String(mailSheet.getRange('C' + r).getDisplayValue());
    body += line + '\n';
  }

  Logger.log('メール件名: ' + subject);

  // ── 電話帳からメールアドレスを取得 ──
  var recipients = getMailRecipients_();
  if (!recipients) return;

  if (recipients.toList.length === 0 && recipients.ccList.length === 0) {
    Logger.log('送信先が見つかりませんでした。下書き作成をスキップします。');
    return;
  }

  var toStr = recipients.toList.join(',');
  var ccStr = recipients.ccList.join(',');

  // メール下書きを作成（送信はせず下書き保存のみ）
  GmailApp.createDraft(toStr, subject, body, {
    cc: ccStr
  });

  Logger.log('メール下書きを保存しました → TO: ' + toStr + ' / CC: ' + ccStr);
}



// ============================================================
// 電話帳からメール送信先を取得
// ============================================================

function getMailRecipients_() {
  var phoneBookId = '1GbWVn7HZ7fPWpiv2GTdpMXksU-SBzCM0UdAc7Bwb02M';
  var phoneBook;
  try {
    phoneBook = SpreadsheetApp.openById(phoneBookId)
                  .getSheetByName('社会基盤ユニット_メールアドレス一覧');
  } catch (e) {
    Logger.log('電話帳スプレッドシートを開けません: ' + e.message);
    return null;
  }
  if (!phoneBook) {
    Logger.log('「社会基盤ユニット_メールアドレス一覧」シートが見つかりません');
    return null;
  }

  var pbData = phoneBook.getDataRange().getValues();
  var toList = [];
  var ccList = [];

  // 1行目はヘッダーなのでスキップ
  for (var i = 1; i < pbData.length; i++) {
    var row = pbData[i];
    var dept     = String(row[2]).trim();  // C列: 部
    var position = String(row[4]).trim();  // E列: 職位
    var name     = String(row[5]).trim();  // F列: 氏名
    var mailAddr = String(row[11]).trim(); // L列: 送信用アドレス（優先）
    if (!mailAddr) {
      mailAddr = String(row[8]).trim();    // I列: mail（フォールバック）
    }

    if (!mailAddr) continue;

    var isBucho    = (position === '部長');
    var isGM       = (position === 'GM' || position === 'GM ●');
    var isKikaku   = (dept === '社会基盤企画総括部');

    // ── TO: 部長（ただし社会基盤企画総括部の部長はCCへ） ──
    if (isBucho && !isKikaku) {
      toList.push(mailAddr);
      Logger.log('TO対象: ' + name + '（' + position + ' / ' + dept + '）→ ' + mailAddr);
      continue;
    }

    // ── CC ──
    var ccReason = '';

    // 社会基盤企画総括部の部長
    if (isBucho && isKikaku) {
      ccReason = '部長（社会基盤企画総括部）';
    }
    // GM または GM ●
    else if (isGM) {
      ccReason = position;
    }
    // F列の氏名が「井野 元太」
    else if (name === '井野 元太') {
      ccReason = '氏名指定';
    }

    if (ccReason) {
      ccList.push(mailAddr);
      Logger.log('CC対象: ' + name + '（' + ccReason + '）→ ' + mailAddr);
    }
  }

  Logger.log('TO合計: ' + toList.length + '名 / CC合計: ' + ccList.length + '名');

  return { toList: toList, ccList: ccList };
}


// ============================================================
// 日本の祝日を取得（UrlFetchApp + Google Calendar API V3）
// ============================================================

function getJapaneseHolidays_(year, month) {
  var calendarId = 'ja.japanese%23holiday%40group.v.calendar.google.com';

  var prevMonth = month - 1;
  var prevYear  = year;
  if (prevMonth === 0) { prevMonth = 12; prevYear = year - 1; }

  var timeMin = new Date(prevYear, prevMonth - 1, 1).toISOString();
  var timeMax = new Date(year, month, 1).toISOString();

  var url = 'https://www.googleapis.com/calendar/v3/calendars/' + calendarId
          + '/events'
          + '?timeMin=' + encodeURIComponent(timeMin)
          + '&timeMax=' + encodeURIComponent(timeMax)
          + '&singleEvents=true'
          + '&orderBy=startTime'
          + '&maxResults=50';

  var holidays = [];

  try {
    var response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    if (code !== 200) {
      Logger.log('祝日API応答コード: ' + code);
      return holidays;
    }

    var data = JSON.parse(response.getContentText());
    if (data.items) {
      data.items.forEach(function(event) {
        if (event.start && event.start.date) {
          holidays.push(event.start.date);
        }
      });
    }
  } catch (e) {
    Logger.log('祝日取得エラー: ' + e.message);
  }

  return holidays;
}


// ============================================================
// OAuthスコープ宣言用（この関数は実行不要）
// ============================================================

function authScope_() {
  CalendarApp.getDefaultCalendar();
  GmailApp.getDrafts();
  GmailApp.sendEmail('', '', '');
}
