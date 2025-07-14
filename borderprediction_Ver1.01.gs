function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📌カスタムメニュー")
    .addItem("▶ すべてを実行", "all")
    .addSeparator()
    .addItem("▶ 入力データをDBに転記", "appendInputToDB")
    .addItem("▶ 予測出力", "generatePointForecast")
    .addItem("▶ 送信", "sendToDiscord")
    .addToUi();
}

function all() {
  appendInputToDB();
  generatePointForecast();
  sendToDiscord();
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

  const dbData = dbSheet.getDataRange().getValues();
  const todayData = dbData.filter(row =>
    Utilities.formatDate(new Date(row[0]), "Asia/Tokyo", "yyyy/MM/dd") === today
  );

  let input12 = null, input18 = null, input23 = null;

  Logger.log("---- 時刻の正規化チェック ----");
  todayData.forEach(row => {
    Logger.log("元: " + row[1] + " → 正規化: " + normalizeTimeLabel(row[1]));
    const time = normalizeTimeLabel(row[1]);
    if (time === "12:00") input12 = row;
    if (time === "18:00") input18 = row;
    if (time === "23:00") input23 = row;
  });

  const increases = {
    "12-18": { "+0": [], "+1": [], "+2": [] },
    "18-23": { "+0": [], "+1": [], "+2": [] }
  };

  const formatDateOnly = d => Utilities.formatDate(new Date(d), "Asia/Tokyo", "yyyy/MM/dd");

  for (let i = 1; i < dbData.length; i++) {
    const [currDate, currTime, curr0, curr1, curr2] = dbData[i];
    const [prevDate, prevTime, prev0, prev1, prev2] = dbData[i - 1];

    if (formatDateOnly(currDate) !== formatDateOnly(prevDate)) continue;

    const from = normalizeTimeLabel(prevTime);
    const to = normalizeTimeLabel(currTime);

    const p0 = Number(curr0) - Number(prev0);
    const p1 = Number(curr1) - Number(prev1);
    const p2 = Number(curr2) - Number(prev2);

    if (from === "12:00" && to === "18:00") {
      if (p0 > 0) increases["12-18"]["+0"].push(p0);
      if (p1 > 0) increases["12-18"]["+1"].push(p1);
      if (p2 > 0) increases["12-18"]["+2"].push(p2);
    }

    if (from === "18:00" && to === "23:00") {
      if (p0 > 0) increases["18-23"]["+0"].push(p0);
      if (p1 > 0) increases["18-23"]["+1"].push(p1);
      if (p2 > 0) increases["18-23"]["+2"].push(p2);
    }
  }

  const avgInc = (arr, min = 100) =>
    arr.length > 0
      ? Math.max(Math.round(arr.reduce((a, b) => a + b, 0) / arr.length), min)
      : min;

  let result = { "+0": null, "+1": null, "+2": null };
  let baseTime = "";

  if (input23) {
    result["+0"] = Number(input23[2]);
    result["+1"] = Number(input23[3]);
    result["+2"] = Number(input23[4]);
    baseTime = "23:00 (実測)";
  } else if (input18) {
    result["+0"] = Number(input18[2]) + avgInc(increases["18-23"]["+0"]);
    result["+1"] = Number(input18[3]) + avgInc(increases["18-23"]["+1"]);
    result["+2"] = Number(input18[4]) + avgInc(increases["18-23"]["+2"]);
    baseTime = "18:00 → 23:00 予測";
  } else if (input12) {
    Logger.log("input12: " + input12.slice(2).join(", "));
    Logger.log("avg 12→18: +0=" + avgInc(increases["12-18"]["+0"]) +
               " +1=" + avgInc(increases["12-18"]["+1"]) +
               " +2=" + avgInc(increases["12-18"]["+2"]));

    const pseudo18_0 = Number(input12[2]) + avgInc(increases["12-18"]["+0"]);
    const pseudo18_1 = Number(input12[3]) + avgInc(increases["12-18"]["+1"]);
    const pseudo18_2 = Number(input12[4]) + avgInc(increases["12-18"]["+2"]);

    Logger.log("仮18時: " + pseudo18_0 + ", " + pseudo18_1 + ", " + pseudo18_2);
    Logger.log("avg 18→23: +0=" + avgInc(increases["18-23"]["+0"]) +
               " +1=" + avgInc(increases["18-23"]["+1"]) +
               " +2=" + avgInc(increases["18-23"]["+2"]));

    result["+0"] = pseudo18_0 + avgInc(increases["18-23"]["+0"]);
    result["+1"] = pseudo18_1 + avgInc(increases["18-23"]["+1"]);
    result["+2"] = pseudo18_2 + avgInc(increases["18-23"]["+2"]);
    baseTime = "12:00 → 18:00 → 23:00 予測";
  } else {
    const lastRow = dbData[dbData.length - 1];
    result["+0"] = Number(lastRow[2]);
    result["+1"] = Number(lastRow[3]);
    result["+2"] = Number(lastRow[4]);
    baseTime = "最終実績時点";
  }

  // --- 倍率適用 ---
  const settingData = setting.getRange(2, 1, setting.getLastRow() - 1, 5).getValues();
  let multiplier = 1;
  settingData.forEach(([_, __, rate, start, end]) => {
    const startDate = Utilities.formatDate(new Date(start), "Asia/Tokyo", "yyyy/MM/dd");
    const endDate = Utilities.formatDate(new Date(end), "Asia/Tokyo", "yyyy/MM/dd");
    if (today >= startDate && today <= endDate) multiplier *= Number(rate);
  });

  const final = {
    "+0": Math.round(result["+0"] * multiplier),
    "+1": Math.round(result["+1"] * multiplier),
    "+2": Math.round(result["+2"] * multiplier)
  };

  const outRow = [today, final["+0"], final["+1"], final["+2"]];
  const outLast = output.getRange(2, 1, 1, 8).getValues().flat();
  const matchIndex = outLast.findIndex(val => val === today);

  if (matchIndex !== -1) {
    output.getRange(matchIndex + 2, 1, 1, outRow.length).setValues([outRow]);
  } else {
    output.getRange(2, 1, 1, outRow.length).setValues([outRow]);
  }

  Logger.log("予測完了: " + baseTime);
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
  const webhookUrl = "https://discordapp.com/api/webhooks/1393586094284476476/gXZhPXKNYMYKHAsUboLvK5JDG_ytEGBQzYgseZ2F-ROrmQAuUgshR80R2DMY89Uxrkw-";

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
