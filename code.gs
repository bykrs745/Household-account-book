function doGet() {
  return HtmlService.createTemplateFromFile('index')
  .evaluate()
  .setTitle('家計簿Surveillance')
  .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// レシート解析のダミー関数（次回以降でGemini 3.1 Proの呼び出しを実装）
function analyzeReceipt(base64Data, mimeType = "image/jpeg") {
  try {
    const props = PropertiesService.getScriptProperties();
    const geminiKey = props.getProperty('GEMINI_API_KEY');
    if (!geminiKey) throw new Error("GEMINI_API_KEYが設定されていません");

    const masters = getMasters();
    const storeNames = Object.keys(masters.store);
    const catNames = Object.keys(masters.category);
    
    if (storeNames.length === 0 || catNames.length === 0) {
        throw new Error("StoreMasterまたはCategoryMasterが正しく設定されていません");
    }

    let storeStr = storeNames.map(k => `・${k}: ${masters.store[k].criteria}`).join('\n');
    let catStr = catNames.map(k => `・${k}: ${masters.category[k].criteria}`).join('\n');
    
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.1-pro-preview:generateContent?key=${geminiKey}`;
    const prompt = `
    このレシート画像から以下のJSON形式でデータを抽出してください。JSON以外のテキストは絶対に含めないでください。
    {
    "storeName": "店舗名",
    "storeType": "${storeNames.join(' or ')}",
    "date": "YYYY/MM/DD",
    "totalAmount": 数値 (実際の支払合計額),
    "items": [{"name": "品目名[税率]", "price": 数値, "category": "${catNames.join(' or ')}"}],
    "isMealPurchase": true または false (弁当、外食、食材買い出しなど、明らかに「食事」のための買い物であるか判定),
    "nutritionalAdvice": "isMealPurchaseがtrueの場合、不足栄養素とそれを補う食品を箇条書きで出力。falseの場合は空文字。"
    }

    【抽出ルールの追加】
    1. 品名の補完: 品名が途切れている場合は、前後の文脈から推測して綺麗な名称に補完してください。
    2. 税率マーカーの付与: 全ての商品の末尾に、レシートの「軽」「外8」マークや一般的な税区分をもとに税率マーカー [10%] または [8%] を必ず付与してください（例: 「赤いきつね[8%]」「洗剤[10%]」）。
    3. 外税と内税の判別: 「小計」の有無を目印にしてください。「小計」の記載がある場合は外税計算として、別途記載の「消費税」を独立した品目として抽出し、カテゴリを税金等にして末尾に [10%] または [8%] を付与してください。「小計」がなく内税表記（コンビニなど、合計に消費税が含まれる）の場合は、消費税を独立した品目として抽出しないでください。
    4. 手数料の消費税: 「ATM手数料」等の各種手数料に掛かる消費税についても、レシートの記載から算出し、独立した品目として抽出してください。
    5. 割引・ポイント: 「ポイント利用」や「割引」がある場合、priceを【マイナス数値】にして抽出し、カテゴリを適切なものにしてください。

    各項目の "storeType" と "category" は、以下の【判定基準】から最も適切なものを1つ選んで出力してください。
    【店舗タイプの判定基準】
    ${storeStr}

    【品目カテゴリの判定基準】
    ${catStr}
    `;
    const payload = {
      contents: [{
        parts: [
          { text: prompt },
          { inlineData: { mimeType: mimeType, data: base64Data } }
        ]
      }]
    };
    
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (result.error) throw new Error(result.error.message);
    
    let textRes = result.candidates[0].content.parts[0].text;
    textRes = textRes.replace(/```json/g, "").replace(/```/g, "").trim();
    const data = JSON.parse(textRes);
    
    saveDataToSheet(data);
    return { success: true, message: "解析および保存が完了しました", data: data };
  } catch (e) {
    throw new Error("解析エラー: " + e.message);
  }
}

function saveDataToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const receiptId = Utilities.getUuid();
  
  // Receiptsシートへの保存 (F列にアドバイスを追加)
  const receiptsSheet = ss.getSheetByName('Receipts');
  if (!receiptsSheet) throw new Error("Receiptsシートが存在しません");
  const advice = data.nutritionalAdvice || "";
  receiptsSheet.appendRow([receiptId, data.date, data.storeName, data.storeType, data.totalAmount, advice]);
  
  // Itemsシートへの保存
  const itemsSheet = ss.getSheetByName('Items');
  if (!itemsSheet) throw new Error("Itemsシートが存在しません");
  
  if (data.items && data.items.length > 0) {
    const itemsData = data.items.map(item => [receiptId, item.name, item.category, item.price, ""]);
    itemsSheet.getRange(itemsSheet.getLastRow() + 1, 1, itemsData.length, itemsData[0].length).setValues(itemsData);
  }
}

function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) return {};
  
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }
  return settings;
}

function saveSettings(settingsObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    sheet.appendRow(['設定項目', '値']);
  }
  
  const data = sheet.getDataRange().getValues();
  
  for (const key in settingsObj) {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(settingsObj[key]);
        found = true;
        break;
      }
    }
    if (!found) sheet.appendRow([key, settingsObj[key]]);
  }
}

function generateSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const receiptsSheet = ss.getSheetByName('Receipts');
  const itemsSheet = ss.getSheetByName('Items');
    if (!receiptsSheet || !itemsSheet) {
    throw new Error("Receipts または Items シートが存在しません");
  }
  const oldNutrients = ss.getSheetByName('NutrientsDB');
  if (oldNutrients) ss.deleteSheet(oldNutrients);
  let lifespanSheet = ss.getSheetByName('ItemLifespan');
  if (!lifespanSheet) {
    lifespanSheet = ss.insertSheet('ItemLifespan');
    lifespanSheet.appendRow(['品名', '大人1人あたりの目安日数(日)']);
  }
  if (lifespanSheet.getLastRow() > 1) {
    lifespanSheet.getRange(2, 1, lifespanSheet.getLastRow() - 1, lifespanSheet.getLastColumn()).clearContent();
  }
    const sampleLifespan = [
    ["洗剤", 60], ["シャンプー", 90], ["トイレットペーパー", 30], ["ボディソープ", 90]
  ];
  lifespanSheet.getRange(2, 1, sampleLifespan.length, sampleLifespan[0].length).setValues(sampleLifespan);
  
  // 新マスターを意識したカテゴリ名に変更
  const sampleTemplates = [
    { name: "スーパーA", type: "スーパー", items: [{n:"豚肉",c:"食品",p:500},{n:"牛乳",c:"食品",p:200},{n:"洗剤",c:"日用品",p:400}] },
    { name: "コンビニB", type: "コンビニ", items: [{n:"おにぎり",c:"食品",p:300},{n:"お茶",c:"食品",p:150},{n:"雑誌",c:"趣味・交遊",p:400}] },
    { name: "レストランC", type: "外食", items: [{n:"ランチセット",c:"趣味・交遊",p:1500}] },
    { name: "ドラッグストアD", type: "その他", items: [{n:"ボディソープ",c:"日用品",p:800},{n:"シャンプー",c:"日用品",p:700}] },
    { name: "スーパーA", type: "スーパー", items: [{n:"鶏肉",c:"食品",p:400},{n:"トイレットペーパー",c:"日用品",p:900},{n:"卵",c:"食品",p:250}] }
  ];
  const receiptsRows = [];
  const itemsRows = [];
  const today = new Date();
  for (let i = 0; i < 30; i++) {
    const d = new Date(today);
    d.setDate(d.getDate() - Math.floor(Math.random() * 90));
    const dateStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd');
    
    const tmpl = sampleTemplates[Math.floor(Math.random() * sampleTemplates.length)];
    const receiptId = Utilities.getUuid();
    
    let total = 0;
    tmpl.items.forEach(item => {
      total += item.p;
      itemsRows.push([receiptId, item.n, item.c, item.p, ""]);
    });
    
    receiptsRows.push([receiptId, dateStr, tmpl.name, tmpl.type, total, "・ビタミンD: 椎茸\n・カルシウム: 牛乳"]);
  }
  receiptsRows.sort((a, b) => new Date(a[1]) - new Date(b[1]));
  receiptsSheet.getRange(receiptsSheet.getLastRow() + 1, 1, receiptsRows.length, receiptsRows[0].length).setValues(receiptsRows);
  itemsSheet.getRange(itemsSheet.getLastRow() + 1, 1, itemsRows.length, itemsRows[0].length).setValues(itemsRows);
  return "3ヶ月分（約30件）のサンプルデータを追加しました。";
  }

function clearAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const receiptsSheet = ss.getSheetByName('Receipts');
  const itemsSheet = ss.getSheetByName('Items');
  
  if (receiptsSheet && receiptsSheet.getLastRow() > 1) {
    receiptsSheet.getRange(2, 1, receiptsSheet.getLastRow() - 1, receiptsSheet.getLastColumn()).clearContent();
  }
  
  if (itemsSheet && itemsSheet.getLastRow() > 1) {
    itemsSheet.getRange(2, 1, itemsSheet.getLastRow() - 1, itemsSheet.getLastColumn()).clearContent();
  }

  return "レシートと品目のデータをクリアしました。";
}

function getAnalysisData(offsetMonth) {
  if (offsetMonth === undefined) offsetMonth = 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const receiptsSheet = ss.getSheetByName('Receipts');
  const itemsSheet = ss.getSheetByName('Items');
  const lifespanSheet = ss.getSheetByName('ItemLifespan');
  if (!receiptsSheet || !itemsSheet) return { error: "データがありません。" };
  const masters = getMasters();
  const storeTypesList = Object.keys(masters.store);
  const defaultStore = storeTypesList.includes('その他') ? 'その他' : storeTypesList[storeTypesList.length - 1];
  const receiptsData = receiptsSheet.getDataRange().getValues().slice(1);
  const itemsData = itemsSheet.getDataRange().getValues().slice(1);
    const now = new Date();
  const fixedMonths = [];
  for (let i = 2; i >= 0; i--) {
    let d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    fixedMonths.push(`${d.getFullYear()}/${(d.getMonth() + 1).toString().padStart(2, '0')}`);
  }
  
  const monthlyDataStore = {}; const monthlyDataCat = {};
  const monthlyItemsStore = {}; const monthlyItemsCat = {};
  const monthlyTotals = {};
  fixedMonths.forEach(m => {
    monthlyDataStore[m] = {}; monthlyDataCat[m] = {};
    monthlyItemsStore[m] = {}; monthlyItemsCat[m] = {};
    storeTypesList.forEach(st => { monthlyDataStore[m][st] = 0; });
    monthlyTotals[m] = 0;
  });
  
  const targetDate = new Date(now.getFullYear(), now.getMonth() + offsetMonth, 1);
  const targetYear = targetDate.getFullYear(); const targetMonth = targetDate.getMonth();
  const targetYm = `${targetYear}/${(targetMonth + 1).toString().padStart(2, '0')}`;
    const calendarData = {}; const itemsMap = {};
  itemsData.forEach(item => {
    if (!item[0]) return;
    if (!itemsMap[item[0]]) itemsMap[item[0]] = [];
    itemsMap[item[0]].push({ name: item[1], category: item[2], price: Number(item[3]) });
  });
    const lifespanData = lifespanSheet ? lifespanSheet.getDataRange().getValues().slice(1) : [];
  const lifeSpanMap = {};
  lifespanData.forEach(row => { if (row[0]) lifeSpanMap[row[0]] = Number(row[1]) || 0; });
  let familyFactor = 1;
  const settings = getSettings();
  if (settings.FamilySettings) {
    try {
      const fs = JSON.parse(settings.FamilySettings);
      let adults = 0, kids = 0;
      fs.forEach(f => { if (f.role === '大人') adults++; else if (f.role === '子供') kids++; });
      familyFactor = (adults || 1) + (kids * 0.5);
    } catch(e){}
  }
  const itemPurchaseHistory = {};
  receiptsData.forEach(r => {
    if (!r[0]) return;
    const time = new Date(r[1]).getTime();
    const rItems = itemsMap[r[0]] || [];
    rItems.forEach(i => {
      if (!itemPurchaseHistory[i.name]) itemPurchaseHistory[i.name] = [];
      itemPurchaseHistory[i.name].push({ time: time, rId: r[0], price: i.price });
    });
  });
  const wasteItemSet = new Set();
  const ONE_DAY = 24 * 60 * 60 * 1000;
  for (const name in itemPurchaseHistory) {
    const history = itemPurchaseHistory[name].sort((a, b) => a.time - b.time);
    const adjustedLifespan = (lifeSpanMap[name] || 0) / familyFactor;
    for (let i = 1; i < history.length; i++) {
      const diffDays = (history[i].time - history[i-1].time) / ONE_DAY;
      if (adjustedLifespan > 0 && diffDays > 0 && diffDays < adjustedLifespan * 0.4) {
        wasteItemSet.add(`${history[i].rId}_${name}`);
      }
    }
  }
  let categoryTotal = {};
  const categoryItems = {};
  let storeTotal = {}; 
  const storeItemsDoughnut = {}; 
  
  receiptsData.forEach(r => {
    if (!r[0]) return;
    const date = new Date(r[1]);
    const ym = `${date.getFullYear()}/${(date.getMonth() + 1).toString().padStart(2, '0')}`;
    const storeName = r[2];
    const storeType = storeTypesList.includes(r[3]) ? r[3] : defaultStore;
    const amount = Number(r[4]) || 0;
    const rItems = itemsMap[r[0]] || [];
    
    if (monthlyDataStore[ym]) {
      monthlyDataStore[ym][storeType] = (monthlyDataStore[ym][storeType] || 0) + amount;
      monthlyTotals[ym] += amount;
      if (!monthlyItemsStore[ym][storeType]) monthlyItemsStore[ym][storeType] = {};
      
      rItems.forEach(i => {
        let cat = i.category;
        if (wasteItemSet.has(`${r[0]}_${i.name}`)) cat = '要検証';
        else if (!masters.category[cat] && cat !== '要検証') cat = 'その他';
        
        monthlyDataCat[ym][cat] = (monthlyDataCat[ym][cat] || 0) + i.price;
        if (!monthlyItemsStore[ym][storeType][i.name]) monthlyItemsStore[ym][storeType][i.name] = { count: 0, price: i.price };
        monthlyItemsStore[ym][storeType][i.name].count++;
        
        if (!monthlyItemsCat[ym][cat]) monthlyItemsCat[ym][cat] = {};
        if (!monthlyItemsCat[ym][cat][i.name]) monthlyItemsCat[ym][cat][i.name] = { count: 0, price: i.price };
        monthlyItemsCat[ym][cat][i.name].count++;
      });
    }
    
    if (ym === targetYm) {
      const ymd = `${date.getFullYear()}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getDate().toString().padStart(2, '0')}`;
      if (!calendarData[ymd]) calendarData[ymd] = [];
      storeTotal[storeType] = (storeTotal[storeType] || 0) + amount;
      if (!storeItemsDoughnut[storeType]) storeItemsDoughnut[storeType] = {};
      
      const displayItems = [];
      rItems.forEach(i => {
        let cat = i.category;
        if (wasteItemSet.has(`${r[0]}_${i.name}`)) cat = '要検証';
        else if (!masters.category[cat] && cat !== '要検証') cat = 'その他';
        
        categoryTotal[cat] = (categoryTotal[cat] || 0) + i.price;
        displayItems.push({ name: i.name, category: cat, price: i.price });
        
        if (!categoryItems[cat]) categoryItems[cat] = {};
        if (!categoryItems[cat][i.name]) categoryItems[cat][i.name] = { count: 0, price: i.price };
        categoryItems[cat][i.name].count++;
        
        if (!storeItemsDoughnut[storeType][i.name]) storeItemsDoughnut[storeType][i.name] = { count: 0, price: i.price };
        storeItemsDoughnut[storeType][i.name].count++;
      });
      calendarData[ymd].push({ storeName: storeName, type: storeType, amount: amount, items: displayItems });
    }
  });
  const monthlyStats = fixedMonths.map((m, idx) => {
    let diffText = '';
    let diffColor = '#e0e0e0';
    if (idx === 2) {
      const current = monthlyTotals[m]; const prev = monthlyTotals[fixedMonths[1]];
      if (current < prev) { diffText = `⬇︎ ¥${prev - current} 減少`; diffColor = '#64b5f6'; } 
      else if (current > prev) { diffText = `⬆︎ ¥${current - prev} 増加`; diffColor = '#e57373'; }
    }
    const monthNum = parseInt(m.split('/')[1], 10);
    return { label: `${monthNum}月`, total: monthlyTotals[m], diffText: diffText, diffColor: diffColor, rawLabel: m };
  });
  const holidayDates = {};
  try {
    const holidayCalendar = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
    if (holidayCalendar) {
      const holidays = holidayCalendar.getEvents(new Date(targetYear, targetMonth, 1), new Date(targetYear, targetMonth + 1, 1));
      holidays.forEach(h => {
        const dStr = Utilities.formatDate(h.getStartTime(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
        holidayDates[dStr] = h.getTitle();
      });
    }
  } catch (e) {}
    return {
    months: fixedMonths, monthlyDataStore: monthlyDataStore, monthlyDataCat: monthlyDataCat,
    monthlyStats: monthlyStats, monthlyItemsStore: monthlyItemsStore, monthlyItemsCat: monthlyItemsCat,
    storeTotal: storeTotal, storeItemsDoughnut: storeItemsDoughnut,
    calendarData: calendarData, categoryTotal: categoryTotal, categoryItems: categoryItems, holidayDates: holidayDates,
    targetYear: targetYear, targetMonth: targetMonth, targetMonthLabel: `${targetYear}年 ${targetMonth + 1}月`,
    masters: masters
  };
}

function getSummaryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const receiptsSheet = ss.getSheetByName('Receipts');
  const itemsSheet = ss.getSheetByName('Items');
  const lifespanSheet = ss.getSheetByName('ItemLifespan');
  
  if (!receiptsSheet || !itemsSheet) return { totalExpense: 0, missingNutrients: '-', wasteWarning: 'なし' };
  
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();
  let currentTotal = 0;
  let latestAdvice = 'まだ食事の記録がありません';
  let latestDate = 0;
  
  const receiptsData = receiptsSheet.getDataRange().getValues().slice(1);
  const itemsData = itemsSheet.getDataRange().getValues().slice(1);
  
  const receiptDates = {};
  const sortedReceipts = [];
  
  receiptsData.forEach(r => {
    if (!r[0]) return;
    const d = new Date(r[1]);
    const amount = Number(r[4]) || 0;
    const advice = r[5] || '';
    
    receiptDates[r[0]] = d.getTime();
    sortedReceipts.push({id: r[0], time: d.getTime(), amount: amount});
    
    if (d.getMonth() === currentMonth && d.getFullYear() === currentYear) {
      currentTotal += amount;
      if (advice && d.getTime() > latestDate) {
        latestAdvice = advice;
        latestDate = d.getTime();
      }
    }
  });

  sortedReceipts.sort((a, b) => a.time - b.time);
  let wasteWarningItems = [];

  // --- 1. 早期再購入のチェック ---
  if (lifespanSheet && receiptsData.length > 0) {
    const lifespanData = lifespanSheet.getDataRange().getValues().slice(1);
    const lifeSpanMap = {};
    lifespanData.forEach(row => { if (row[0]) lifeSpanMap[row[0]] = Number(row[1]) || 0; });

    const settings = getSettings();
    let familyFactor = 1;
    if (settings.FamilyStructure) {
      const adults = (settings.FamilyStructure.match(/大人(\d+)人/) || [0,1])[1];
      const kids = (settings.FamilyStructure.match(/子供(\d+)人/) || [0,0])[1];
      familyFactor = Number(adults) + (Number(kids) * 0.5);
    }

    const itemPurchaseHistory = {};
    itemsData.forEach(item => {
      const rId = item[0], name = item[1];
      if (receiptDates[rId] && lifeSpanMap[name]) {
        if (!itemPurchaseHistory[name]) itemPurchaseHistory[name] = [];
        itemPurchaseHistory[name].push(receiptDates[rId]);
      }
    });

    const ONE_DAY = 24 * 60 * 60 * 1000;
    for (const name in itemPurchaseHistory) {
      const dates = itemPurchaseHistory[name].sort((a, b) => a - b);
      if (dates.length >= 2) {
        const diffDays = (dates[dates.length - 1] - dates[dates.length - 2]) / ONE_DAY;
        const adjustedLifespan = lifeSpanMap[name] / familyFactor;
        if (diffDays > 0 && diffDays < adjustedLifespan * 0.4) {
          wasteWarningItems.push(name);
        }
      }
    }
  }

  // --- 2. 前回レシートとの比較（増額時の要因特定） ---
  if (sortedReceipts.length >= 2) {
    const latest = sortedReceipts[sortedReceipts.length - 1];
    const prev = sortedReceipts[sortedReceipts.length - 2];
    
    if (latest.amount > prev.amount) {
      const latestItems = itemsData.filter(i => i[0] === latest.id);
      latestItems.forEach(item => {
        const name = item[1], category = item[2], price = Number(item[3]) || 0;
        // 臨時カテゴリ、または今回の合計額の30%以上を占める品を抽出
        if (category === '臨時' || price > latest.amount * 0.3) {
          wasteWarningItems.push(name);
        }
      });
    }
  }

  const uniqueWarnings = [...new Set(wasteWarningItems)];
  const warningText = uniqueWarnings.length > 0 ? `<span style="color:#cf6679;">${uniqueWarnings.join(', ')}</span>` : 'なし';

  return {
    totalExpense: currentTotal,
    missingNutrients: latestAdvice,
    wasteWarning: warningText
  };
}

function generateMonthlyReport(offsetMonth) {
  // 数値以外（トリガー実行時のイベントオブジェクトや未指定時）は「前月（-1）」とする
  if (typeof offsetMonth !== 'number') offsetMonth = -1;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const receiptsSheet = ss.getSheetByName('Receipts');
  const itemsSheet = ss.getSheetByName('Items');
  if (!receiptsSheet || !itemsSheet) throw new Error("データがありません");

  const now = new Date();
  const targetDate = new Date(now.getFullYear(), now.getMonth() + offsetMonth, 1);
  const currentMonth = targetDate.getMonth();
  const currentYear = targetDate.getFullYear();
  const targetYearMonth = `${currentYear}/${(currentMonth + 1).toString().padStart(2, '0')}`;
  
  const receiptsData = receiptsSheet.getDataRange().getValues().slice(1);
  const currentReceiptIds = new Set();
  receiptsData.forEach(r => {
    if (!r[0]) return;
    const d = new Date(r[1]);
    if (d.getMonth() === currentMonth && d.getFullYear() === currentYear) currentReceiptIds.add(r[0]);
  });

  if (currentReceiptIds.size === 0) throw new Error(`${currentMonth + 1}月の購入履歴がありません`);
  const itemsData = itemsSheet.getDataRange().getValues().slice(1);
  const foodItems = new Set();
  itemsData.forEach(item => {
    // 「食品」に加えて「食事」カテゴリも月次レポートの評価対象に含める
    if (currentReceiptIds.has(item[0]) && ['食品', '食事'].includes(item[2])) foodItems.add(item[1]);
  });

  const settings = getSettings();
  const cookingPref = settings.CookingPreference === 'no_cook' ? '調理しないで済むもの' : '調理するパターン';
  
  // 年齢計算
  const calculateAge = (birthDateStr) => {
    if (!birthDateStr) return '年齢不明';
    const birth = new Date(birthDateStr); const today = new Date();
    let age = today.getFullYear() - birth.getFullYear();
    if (today.getMonth() < birth.getMonth() || (today.getMonth() === birth.getMonth() && today.getDate() < birth.getDate())) age--;
    return age + '歳';
  };

  let familyStr = "一人暮らし";
  if (settings.FamilySettings) {
    try {
      const fs = JSON.parse(settings.FamilySettings);
      if (fs.length > 0) familyStr = fs.map(f => `${f.role}(${f.gender}, ${calculateAge(f.birth)}, 事情: ${f.note || '特になし'})`).join(', ');
    } catch(e){}
  }

  const props = PropertiesService.getScriptProperties();
  const geminiKey = props.getProperty('GEMINI_API_KEY');
  const prompt = `
  あなたは少し辛口な栄養・健康・家計管理アドバイザーです。
  ユーザーの家族構成: ${familyStr}

  以下の当月購入した食品リストから、ユーザーの買い物傾向を分析し、以下の内容でレポートを作成してください。
  1. 栄養バランスに対する辛口な「ダメ出し」
  2. 無駄遣いの傾向や行動パターンの評価と改善案
  3. 添加物や有害成分（リン酸塩、過剰な塩分など）への警告
  4. ユーザーの傾向と家族構成に沿った、次に取り入れるべきプラスアルファの食品アドバイス（条件: ${cookingPref}）

  購入食品リスト:
  ${Array.from(foodItems).join(', ')}
  `;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key=${geminiKey}`;
  const payload = { contents: [{ parts: [{ text: prompt }] }] };
  const options = { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true };
  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());
  if (result.error) throw new Error(result.error.message);
  const reportText = result.candidates[0].content.parts[0].text;

  let reportsSheet = ss.getSheetByName('Reports');
  if (!reportsSheet) { reportsSheet = ss.insertSheet('Reports'); reportsSheet.appendRow(['年月', 'レポート内容']); }
  const reportsData = reportsSheet.getDataRange().getValues();
  let foundRow = -1;
  for (let i = 1; i < reportsData.length; i++) { 
    let cellVal = reportsData[i][0];
    let cellStr = (cellVal instanceof Date) ? Utilities.formatDate(cellVal, Session.getScriptTimeZone(), 'yyyy/MM') : String(cellVal);
    if (cellStr === targetYearMonth) { foundRow = i + 1; break; } 
  }
  if (foundRow > 0) reportsSheet.getRange(foundRow, 2).setValue(reportText);
  else reportsSheet.appendRow([targetYearMonth, reportText]);

  const email = Session.getActiveUser().getEmail();
  if (email) {
    MailApp.sendEmail({
      to: email, subject: `【家計簿Surveillance】${currentMonth + 1}月分の辛口レポート`,
      body: "家計簿の傾向に基づいた辛口レポートが完成しました。\nアプリの「分析」タブから確認してください。\n\n※このメールは自動送信です。"
    });
  }
  return reportText;
}

function getLatestReport(offsetMonth) {
  if (offsetMonth === undefined) offsetMonth = 0;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reportsSheet = ss.getSheetByName('Reports');
  if (!reportsSheet) return null;
    const now = new Date();
  const targetDate = new Date(now.getFullYear(), now.getMonth() + offsetMonth, 1);
  const targetYearMonth = `${targetDate.getFullYear()}/${(targetDate.getMonth() + 1).toString().padStart(2, '0')}`;
    const reportsData = reportsSheet.getDataRange().getValues();
  for (let i = reportsData.length - 1; i > 0; i--) {
    let cellVal = reportsData[i][0];
    let cellStr = (cellVal instanceof Date) ? Utilities.formatDate(cellVal, Session.getScriptTimeZone(), 'yyyy/MM') : String(cellVal);
    if (cellStr === targetYearMonth) {
      return reportsData[i][1];
    }
  }
  return null;
}

// --- 設定の一括取得・保存 ---
function getAllSettings() {
  const settings = getSettings();
  const masters = getMasters();
  return { settings: settings, masters: masters };
}

function saveAllSettings(data) {
  saveSettings({
    FamilySettings: JSON.stringify(data.fam),
    CookingPreference: data.cook
  });
}

// --- マスタデータの取得 ---
function getMasters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const storeSheet = ss.getSheetByName('StoreMaster');
  const storeData = storeSheet ? storeSheet.getDataRange().getValues().slice(1) : [];
  const storeMaster = {};
  storeData.forEach(r => {
    if (r[0]) storeMaster[r[0]] = { svg: r[1], hex: r[2], criteria: r[3] };
  });

  const categorySheet = ss.getSheetByName('CategoryMaster');
  const categoryData = categorySheet ? categorySheet.getDataRange().getValues().slice(1) : [];
  const categoryMaster = {};
  categoryData.forEach(r => {
    if (r[0]) categoryMaster[r[0]] = { hex: r[1], criteria: r[2] };
  });

  return { store: storeMaster, category: categoryMaster };
}

// --- ドライブTempフォルダからの一括処理 ---
function analyzeBatchFromDrive() {
  try {
    const tempFolder = getOrCreateFolder("家計簿Surveillance_Temp");
    const doneFolder = getOrCreateFolder("家計簿Surveillance_Done");
    const files = tempFolder.getFiles();
    let count = 0;
    let errCount = 0;
    
    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();
      if (mimeType.includes("image") || mimeType === "application/pdf") {
        try {
          const base64Data = Utilities.base64Encode(file.getBlob().getBytes());
          analyzeReceipt(base64Data, mimeType);
          file.moveTo(doneFolder); // 処理成功したものはDoneへ移動
          count++;
        } catch(e) {
          errCount++; // エラーのものはTempに残す
        }
      }
    }
    
    if (count === 0 && errCount === 0) return { message: "Tempフォルダに処理できる画像・PDFがありませんでした。" };
    let msg = `${count}件のレシートを一括登録しました。`;
    if (errCount > 0) msg += `\n※${errCount}件の解析に失敗しTempフォルダに残りました。`;
    return { success: true, message: msg };
  } catch (e) {
    throw new Error("一括処理エラー: " + e.message);
  }
}

function getOrCreateFolder(folderName) {
  // ルート限定ではなく、ドライブ全体からフォルダ名で検索する
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  // 見つからなければルート階層に新規作成
  return DriveApp.getRootFolder().createFolder(folderName);
}

// --- アプリからTempフォルダへ画像を保存する ---
function saveImageToTemp(base64Data) {
  try {
    const folder = getOrCreateFolder("家計簿Surveillance_Temp");
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', 'receipt_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.jpg');
    folder.createFile(blob);
    return { success: true, message: "画像をTempフォルダに保存しました。後で一括登録できます。" };
  } catch (e) {
    throw new Error("保存エラー: " + e.message);
  }
}