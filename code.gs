function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📌カスタムメニュー")
    .addItem("▶ 入力データをDBに転記", "appendInputToDB")
    .addItem("▶ 予測出力", "generatePointForecast")
    .addItem("▶ 送信", "sendToDiscord")
    .addToUi();
}


function appendInputToDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("入力");
  const dbSheet = ss.getSheetByName("DB");

  const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
  const inputData = inputSheet.getRange(2, 1, inputSheet.getLastRow() - 1, 4).getValues();

  inputData.forEach(row => {
    const [time, p0, p1, p2] = row;
    if (time && (p0 || p1 || p2)) {
      dbSheet.appendRow([today, time, p0, p1, p2]);
    }
  });

  inputSheet.getRange(2, 1, 1, 4).clearContent();
}

function generatePointForecast() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName("DB");
  const output = ss.getSheetByName("予測出力");
  const setting = ss.getSheetByName("設定");

  const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");

  // --- DBから今日のデータを取得 ---
  const dbData = dbSheet.getDataRange().getValues();
  const todayData = dbData.filter(row => row[0] === today);

  let input12 = null;
  let input18 = null;
  let input23 = null;

  // 12時・18時・23時のデータを取得
  todayData.forEach(row => {
    const timeLabel = normalizeTimeLabel(row[1]);
    if (timeLabel === "12:00") input12 = row;
    if (timeLabel === "18:00") input18 = row;
    if (timeLabel === "23:00") input23 = row;
  });

  // --- 増加傾向の平均を計算 ---
  const increases = {
    "12-18": { "+0": [], "+1": [], "+2": [] },
    "18-23": { "+0": [], "+1": [], "+2": [] },
  };

  for (let i = 1; i < dbData.length; i++) {
    const [date, time, plus0, plus1, plus2] = dbData[i];
    const prevRow = dbData[i - 1];
    if (date !== prevRow[0]) continue;

    const prevTime = normalizeTimeLabel(prevRow[1]);
    const currTime = normalizeTimeLabel(time);

    if (prevTime === "12:00" && currTime === "18:00") {
      const p0 = Number(plus0) - Number(prevRow[2]);
      const p1 = Number(plus1) - Number(prevRow[3]);
      const p2 = Number(plus2) - Number(prevRow[4]);

      if (!isNaN(p0)) increases["12-18"]["+0"].push(p0);
      if (!isNaN(p1)) increases["12-18"]["+1"].push(p1);
      if (!isNaN(p2)) increases["12-18"]["+2"].push(p2);
    }

    if (prevTime === "18:00" && currTime === "23:00") {
      const p0 = Number(plus0) - Number(prevRow[2]);
      const p1 = Number(plus1) - Number(prevRow[3]);
      const p2 = Number(plus2) - Number(prevRow[4]);

      if (!isNaN(p0)) increases["18-23"]["+0"].push(p0);
      if (!isNaN(p1)) increases["18-23"]["+1"].push(p1);
      if (!isNaN(p2)) increases["18-23"]["+2"].push(p2);
    }
  }

  const avgInc = (arr) => arr.length > 0 ? Math.round(arr.reduce((a, b) => a + b, 0) / arr.length) : null;

  const result = { "+0": null, "+1": null, "+2": null };
  let base = null;
  let baseTime = "";
  let segment = "";

  if (input18) {
    base = input18;
    baseTime = "18:00";
    segment = "18-23";
  } else if (input12) {
    base = input12;
    baseTime = "12:00";
    segment = "12-18";
  } else if (input23) {
    result["+0"] = Number(input23[2]);
    result["+1"] = Number(input23[3]);
    result["+2"] = Number(input23[4]);
    baseTime = "23:00 (実績)";
  }

  if (base && segment) {
    result["+0"] = avgInc(increases[segment]["+0"]) !== null ? Number(base[2]) + avgInc(increases[segment]["+0"]) : null;
    result["+1"] = avgInc(increases[segment]["+1"]) !== null ? Number(base[3]) + avgInc(increases[segment]["+1"]) : null;
    result["+2"] = avgInc(increases[segment]["+2"]) !== null ? Number(base[4]) + avgInc(increases[segment]["+2"]) : null;
  } else {
    // 12時、18時がない場合は、最新の実績（最終行）を予測に使用
    const lastRow = dbData[dbData.length - 1];
    result["+0"] = Number(lastRow[2]);
    result["+1"] = Number(lastRow[3]);
    result["+2"] = Number(lastRow[4]);
    baseTime = "最終実績時点";
  }

  // --- イベント倍率の反映 ---
  const settingData = setting.getRange(2, 1, setting.getLastRow() - 1, 5).getValues();
  let multiplier = 1;

  settingData.forEach(([eventName, status, rate, start, end]) => {
    const startDate = Utilities.formatDate(new Date(start), "Asia/Tokyo", "yyyy/MM/dd");
    const endDate = Utilities.formatDate(new Date(end), "Asia/Tokyo", "yyyy/MM/dd");
    if (today >= startDate && today <= endDate) {
      multiplier *= Number(rate);
    }
  });

  const final = {
    "+0": result["+0"] !== null ? Math.round(result["+0"] * multiplier) : null,
    "+1": result["+1"] !== null ? Math.round(result["+1"] * multiplier) : null,
    "+2": result["+2"] !== null ? Math.round(result["+2"] * multiplier) : null
  };

  // --- 出力 ---
  const outRow = [today, final["+0"], final["+1"], final["+2"]];
  const outLast = output.getRange(2, 1, 1, 8).getValues().flat();
  const matchIndex = outLast.findIndex(val => val === outRow[0]);

  if (matchIndex !== -1) {
    output.getRange(matchIndex + 2, 1, 1, outRow.length).setValues([outRow]);
  } else {
    output.getRange(2, 1, 1, outRow.length).setValues([outRow]);
  }
}

// ← 全角「１２時」とかに対応
function normalizeTimeLabel(label) {
  if (!label) return "";
  return String(label)
    .replace(/[１２３４５６７８９０]/g, s => String.fromCharCode(s.charCodeAt(0) - 65248))
    .replace("時", ":00")
    .replace(/\s/g, "")
    .trim();
}

function sendToDiscord() {
  const webhookUrl = "https://discordapp.com/api/webhooks/";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const output = ss.getSheetByName("予測出力");
  const db = ss.getSheetByName("DB");

  const today = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd");
  const todayRow = output.getRange(2, 1, 1, 7).getValues()[0];
  const dateValue = Utilities.formatDate(new Date(todayRow[0]), "Asia/Tokyo", "yyyy/MM/dd");

  if (dateValue !== today) {
    Logger.log(`今日の予測が見つかりません（見つかった日付: ${dateValue}）`);
    return;
  }

  const plus0 = todayRow[1].toLocaleString();
  const plus1 = todayRow[2].toLocaleString();
  const plus2 = todayRow[3].toLocaleString();

  const dbData = db.getRange(2, 1, db.getLastRow() - 1, 5).getValues();
  const todayTimes = dbData
    .filter(row => Utilities.formatDate(new Date(row[0]), "Asia/Tokyo", "yyyy/MM/dd") === today)
    .map(row => row[1]);

  const has12 = todayTimes.includes("１２時");
  const has18 = todayTimes.includes("１８時");
  const has23 = todayTimes.includes("２３時");

  let label = "";
  if (has23) {
    label = "【本日の最終結果】（実測）";
  } else if (has18) {
    label = "【18時時点の最終ボーダー予測】";
  } else if (has12) {
    label = "【12時時点の最終ボーダー予測】";
  } else {
    label = "【予測データ不足】";
  }

  const content = `${label}\n+0: ${plus0}\n+1: ${plus1}\n+2: ${plus2}`;

  const payload = {
    content: content
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  UrlFetchApp.fetch(webhookUrl, options);
}
