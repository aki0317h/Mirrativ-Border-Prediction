## Point Forecasting System for Google Sheets (GAS)

This repository contains a Google Apps Script-based forecasting system that predicts point totals for 23:00 based on past trends and real-time inputs at 12:00 and 18:00.

### Features

* Input and store point data at 12:00, 18:00, and 23:00.
* Predict 23:00 point totals using average growth trends from historical data.
* Apply event-based multipliers depending on active date ranges and status (Low / Medium / High).
* Output is updated dynamically based on the latest available time (either 12:00 or 18:00).
* Fallback to actual 23:00 value if no forecast is possible.
* Structured across multiple sheets: `入力 (Input)`, `DB`, `出力 (Output)`, and `設定 (Settings)`.

### Use Cases

* Forecasting competitive rankings in games or applications.
* Time-based point progression analysis.
* Event-aware projections for leaderboard systems.

### Technologies

* Google Apps Script (GAS)
* Google Sheets

### How to Use

1. Enter point data into the `入力` sheet.
2. Run the forecast function (e.g. via button or script menu).
3. View forecasted results in the `出力` sheet.
4. Manage event multipliers in the `設定` sheet.

---

Designed for automation, reliability, and clarity in forecasting point-based systems within spreadsheet environments.
