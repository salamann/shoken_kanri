function sendSheetByMail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("メール用");

  const values = sheet.getDataRange().getDisplayValues();
  const sheet1 = ss.getSheetByName("まとめ");
  const charts = sheet1.getCharts();
  const sheet2 = ss.getSheetByName("時系列円");
  const charts2 = sheet2.getCharts();
  const sheet3 = ss.getSheetByName("アセットアロケーション");
  const charts3 = sheet3.getCharts();


  const inlineImages = {};
  let comment;

  try {
    comment = retryWithBackoff(
      () => generateComment(),
      4
    );
  }
  catch (e) {

    Logger.log(`AI generation failed: ${e}`);

    comment =
      "今週のAIコメント生成は利用できませんでした。";
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
  <h2>投資週報</h2>

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
      今週のまとめ
    </div>
    ${comment}
  </div>
`;

  let chartId = 0;

  // 時系列円シート
  charts2.forEach((chart2) => {
    const cid = `chart${chartId++}`;
    const blob = chart2.getBlob().setName(cid);

    inlineImages[cid] = blob;

    html += `
      <div class="chart">
        <img src="cid:${cid}">
      </div>
    `;
  });


  // まとめシート
  charts.forEach((chart) => {
    const cid = `chart${chartId++}`;
    const blob = chart.getBlob().setName(cid);

    inlineImages[cid] = blob;

    html += `
      <div class="chart">
        <img src="cid:${cid}">
      </div>
    `;
  });

  // アセットアロケーションシート
  charts3.forEach((chart3) => {
    const cid = `chart${chartId++}`;
    const blob = chart3.getBlob().setName(cid);

    inlineImages[cid] = blob;

    html += `
      <div class="chart">
        <img src="cid:${cid}">
      </div>
    `;
  });

  html += '<table>';

  const diffCol = values[0].indexOf("先週比");

  values.forEach((row, i) => {
    html += '<tr>';

    row.forEach((cell, col) => {

      // ヘッダ行
      if (i === 0) {
        html += `<th>${cell}</th>`;
        return;
      }

      let style = "";

      // 先週比列だけ色付け
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
    "資産週報",
    "HTMLメールをご覧ください",
    {
      htmlBody: html,
      inlineImages: inlineImages
    }
  );
}


function generateComment() {

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("GEMINI_API_KEY");

  const url =
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("メール用");

  const values = sheet.getDataRange().getDisplayValues();

  // ヘッダ
  const headers = values[0];

  const nameCol = headers.indexOf("商品名");
  const valueCol = headers.indexOf("市場価値円");
  const diffCol = headers.indexOf("先週比");

  // 商品行 (2～11行想定、合計除外)
  const rows = values.slice(1, -1);

  // 合計行
  const totalRow = values[values.length - 1];

  const totalValue = totalRow[valueCol];
  const totalDiff = totalRow[diffCol];

  // 先週比抽出（数値化）
  const movers = rows.map(r => {

    const diffText = r[diffCol];

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

  const topGain = movers[0];
  const topLoss = movers[movers.length - 1];

  const prompt = `
あなたは金融レポート編集者です。

以下データを基に、
投資週報の「今週のまとめ」を
80〜120字で日本語で書いてください。

条件:
- 事実ベース
- 金融レポート風
- 投資助言禁止
- 誇張禁止

データ:
総資産: ${totalValue}
先週比: ${totalDiff}

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