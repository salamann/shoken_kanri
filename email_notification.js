function sendReportByMail(sheetName, reportType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`${sheetName} シートが見つかりません`);
  }

  const values = sheet.getDataRange().getDisplayValues();
  const chartSheetNames = ["まとめ", "時系列円", "アセットアロケーション"];

  const inlineImages = {};
  let comment;

  try {
    comment = retryWithBackoff(
      () => generateComment(sheetName, reportType),
      4
    );
  }
  catch (e) {
    Logger.log(`AI generation failed: ${e}`);
    comment = reportType === "月報"
      ? "今月のAIコメント生成は利用できませんでした。"
      : "今週のAIコメント生成は利用できませんでした。";
  }

  let html = `
  <html>
  <head>
  <style>
  body {
    font-family: Arial, sans-serif;
    background: #f5f7fb;
    color: #333;
    padding: 20px;
  }

  .container {
    max-width: 900px;
    margin: auto;
    background: white;
    border-radius: 12px;
    padding: 24px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
  }

  h2 {
    margin-top: 0;
    color: #1a73e8;
    border-bottom: 2px solid #e0e0e0;
    padding-bottom: 8px;
  }

  .chart {
    text-align: center;
    margin: 20px 0;
  }

  .chart img {
    max-width: 100%;
    border-radius: 8px;
  }

  table {
    width: 100%;
    border-collapse: collapse;
    margin-top: 24px;
    font-size: 14px;
  }

  th {
    background: #1a73e8;
    color: white;
    padding: 10px;
    text-align: left;
  }

  td {
    padding: 10px;
    border-bottom: 1px solid #e5e7eb;
  }

  tr:nth-child(even) {
    background: #f8fafc;
  }

  tr:hover {
    background: #eef5ff;
  }
  </style>
  </head>
  <body>
  <div class="container">
  <h2>投資${reportType}</h2>

  <div style="
    background:#f8fafc;
    border-left:4px solid #1a73e8;
    padding:14px;
    margin:20px 0;
    border-radius:8px;
    line-height:1.6;
  ">
    <div style="
      font-weight:600;
      margin-bottom:6px;
      color:#1a73e8;
    ">
      ${reportType === "月報" ? "今月のまとめ" : "今週のまとめ"}
    </div>
    ${comment}
  </div>
`;

  let chartId = 0;

  chartSheetNames.forEach((name) => {
    const chartSheet = ss.getSheetByName(name);
    if (!chartSheet) {
      return;
    }

    chartSheet.getCharts().forEach((chart) => {
      const cid = `chart${chartId++}`;
      const blob = chart.getBlob().setName(cid);
      inlineImages[cid] = blob;

      html += `
      <div class="chart">
        <img src="cid:${cid}">
      </div>
    `;
    });
  });

  html += '<table>';

  const diffColNames = ["先週比", "先月比", "前月比"];
  const diffCol = diffColNames.reduce(
    (acc, header) => acc >= 0 ? acc : values[0].indexOf(header),
    -1
  );

  values.forEach((row, i) => {
    html += '<tr>';

    row.forEach((cell, col) => {
      if (i === 0) {
        html += `<th>${cell}</th>`;
        return;
      }

      let style = "";
      if (col === diffCol) {
        if (cell.includes("+")) {
          style = `
            color:#2563eb;
            font-weight:600;
            background:#eff6ff;
          `;
        }
        else if (cell.includes("-")) {
          style = `
            color:#dc2626;
            font-weight:600;
            background:#fef2f2;
          `;
        }
      }

      html += `<td style="${style}">${cell}</td>`;
    });

    html += '</tr>';
  });

  html += `
  </table>
  </div>
  </body>
  </html>
  `;

  GmailApp.sendEmail(
    EMAIL_TO,
    `資産${reportType}`,
    "HTMLメールをご覧ください",
    {
      htmlBody: html,
      inlineImages: inlineImages
    }
  );
}

function sendSheetByMail() {
  sendReportByMail("週報", "週報");
}

function sendMonthlyReportByMail() {
  sendReportByMail("月報", "月報");
}


function generateComment(sheetName = "週報", reportType = "週報") {

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("GEMINI_API_KEY");

  const url =
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`${sheetName} シートが見つかりません`);
  }

  const values = sheet.getDataRange().getDisplayValues();

  // ヘッダ
  const headers = values[0];

  const nameCol = headers.indexOf("商品名");
  const valueCol = headers.indexOf("市場価値円");
  const diffColNames = ["先週比", "先月比", "前月比"];
  const diffCol = diffColNames.reduce(
    (acc, header) => acc >= 0 ? acc : headers.indexOf(header),
    -1
  );

  // 商品行 (2～11行想定、合計除外)
  const rows = values.slice(1, -1);

  // 合計行
  const totalRow = values[values.length - 1];

  const totalValue = totalRow[valueCol];
  const totalDiff = diffCol >= 0 ? totalRow[diffCol] : "";

  // 先週比 / 先月比 抽出（数値化）
  const movers = rows.map(r => {
    const diffText = diffCol >= 0 ? r[diffCol] : "";
    const match = diffText.match(/[+-]?¥?([\d,]+)/);
    const diff =
      match ?
        Number(match[1].replace(/,/g, '')) *
        (diffText.includes('-') ? -1 : 1)
        : 0;

    return {
      name: r[nameCol],
      diff: diff
    };
  });

  movers.sort((a, b) => b.diff - a.diff);

  const topGain = movers[0] || { name: "", diff: 0 };
  const topLoss = movers[movers.length - 1] || { name: "", diff: 0 };
  const summaryLabel = reportType === "月報" ? "今月のまとめ" : "今週のまとめ";
  const compareLabel = diffCol >= 0 ? headers[diffCol] : "先週比";

  const prompt = `
あなたは金融レポート編集者です。

以下データを基に、
投資${reportType}の「${summaryLabel}」を
80〜120字で日本語で書いてください。

条件:
- 事実ベース
- 金融レポート風
- 投資助言禁止
- 誇張禁止

データ:
総資産: ${totalValue}
${compareLabel}: ${totalDiff}

上昇寄与:
${topGain.name} ${topGain.diff}

下落寄与:
${topLoss.name} ${topLoss.diff}
`;

  const payload = {
    contents: [{
      parts: [{
        text: prompt
      }]
    }]
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  });

  const json = JSON.parse(response.getContentText());

  return json.candidates[0]
    .content.parts[0].text.trim();
}



function retryWithBackoff(fn, maxRetries = 5) {

  for (let attempt = 0; attempt <= maxRetries; attempt++) {

    try {
      return fn();
    }
    catch (e) {

      Logger.log(
        `AI call failed attempt=${attempt + 1} error=${e}`
      );

      // 最後なら投げる
      if (attempt === maxRetries) {
        throw e;
      }

      // 1s → 2s → 4s → 8s → 16s
      const waitMs =
        Math.pow(2, attempt) * 1000 * 8 +
        Math.floor(Math.random() * 500); // jitter

      Logger.log(`Retry after ${waitMs} ms`);

      Utilities.sleep(waitMs);
    }
  }
}