


const LINE_CHANNEL_ACCESS_TOKEN = '';


const LINE_CHANNEL_SECRET = '';


const GOOGLE_SHEET_ID = '';


/****************************************************************
 * 主要進入點函式
 * doGet(e) 用於驗證 Webhook URL
 * doPost(e) 用於接收與處理 LINE 的事件請求
 ****************************************************************/

/**
 * 當 LINE Developers Console 設定 Webhook URL 時，用於驗證。
 */
function doGet(e) {
  return ContentService.createTextOutput("Google Apps Script is running.");
}

/**
 * 接收來自 LINE 的 Webhook 請求
 * @param {Object} e - LINE Webhook 事件物件
 */
function doPost(e) {
  try {
    const event = JSON.parse(e.postData.contents).events[0];
    const replyToken = event.replyToken;
    const userId = event.source.userId;

    // 根據事件類型進行處理
    if (event.type === 'message' && event.message.type === 'text') {
      handleTextMessage(replyToken, event.message.text, userId);
    } else if (event.type === 'postback') {
      // 處理 postback 事件，例如 Flex Message 的按鈕點擊
      const postbackData = parsePostbackData(event.postback.data);
      handlePostback(replyToken, postbackData, userId);
    }
  } catch (error) {
    Logger.log('發生錯誤: ' + error.message);
    Logger.log('錯誤堆疊: ' + error.stack);
  }
}


/****************************************************************
 * 訊息與事件處理核心函式
 ****************************************************************/

/**
 * 處理文字訊息
 * @param {string} replyToken - 回覆用的 Token
 * @param {string} userMessage - 使用者傳送的訊息
 * @param {string} userId - 使用者 ID
 */
function handleTextMessage(replyToken, userMessage, userId) {
  const userState = getUserState(userId);

  // 檢查使用者是否處於某個對話流程中
  if (userState.action) {
    switch (userState.action) {
      case 'awaiting_item':
        userState.item = userMessage;
        userState.action = 'awaiting_amount';
        setUserState(userId, userState);
        replyMessage(replyToken, createTextMessage('請輸入金額：'));
        break;
      case 'awaiting_amount':
        const amount = parseFloat(userMessage);
        if (!isNaN(amount) && amount > 0) {
          recordTransaction(replyToken, userId, userState.type, userState.category, userState.item, amount);
          clearUserState(userId);
        } else {
          replyMessage(replyToken, createTextMessage('金額格式錯誤，請輸入一個有效的數字。'));
        }
        break;
      case 'awaiting_new_category':
        addNewCategory(replyToken, userId, userMessage.trim());
        clearUserState(userId);
        break;
      case 'awaiting_stock_code':
        getStockPrice(replyToken, userMessage.trim());
        clearUserState(userId);
        break;
      default:
        clearUserState(userId);
        routeMainCommands(replyToken, userMessage, userId);
        break;
    }
  } else {
    // 根據關鍵字路由到不同功能
    routeMainCommands(replyToken, userMessage, userId);
  }
}

/**
 * 根據使用者輸入的指令，分流至對應的功能
 */
function routeMainCommands(replyToken, userMessage, userId) {
  switch (userMessage.trim()) {
    case '記帳':
      askIncomeOrExpense(replyToken);
      break;
    case '收支類別':
      manageCategories(replyToken, userId);
      break;
    case '收支報表':
      askReportPeriod(replyToken);
      break;
    case '匯率':
      askCurrency(replyToken);
      break;
    case '股票':
      setUserState(userId, { action: 'awaiting_stock_code' });
      replyMessage(replyToken, createTextMessage('請輸入股票代碼：'));
      break;
    default:
      // 可在此加入預設回覆或幫助訊息
      break;
  }
}

/**
 * 處理 Postback 事件
 * @param {string} replyToken - 回覆用的 Token
 * @param {Object} data - Postback 資料物件
 * @param {string} userId - 使用者 ID
 */
function handlePostback(replyToken, data, userId) {
  switch (data.action) {
    case 'select_type':
      askCategory(replyToken, data.type, userId);
      break;
    case 'select_category':
      const state = {
        action: 'awaiting_item',
        type: data.type,
        category: data.category
      };
      setUserState(userId, state);
      replyMessage(replyToken, createTextMessage('請輸入花費品項：'));
      break;
    case 'manage_category':
      if (data.do === 'add') {
        setUserState(userId, { action: 'awaiting_new_category' });
        replyMessage(replyToken, createTextMessage('請輸入要新增的類別名稱：'));
      } else if (data.do === 'delete') {
        askDeleteCategory(replyToken, userId);
      }
      break;
    case 'delete_category':
      deleteCategory(replyToken, userId, data.category);
      break;
    case 'generate_report':
      generateReport(replyToken, data.period, userId);
      break;
    case 'query_exchange_rate':
      getExchangeRate(replyToken, data.currency);
      break;
  }
}


/****************************************************************
 * 1. 收支紀錄功能
 ****************************************************************/

/**
 * 發送 Flex Message 詢問使用者要記錄「支出」還是「收入」
 */
function askIncomeOrExpense(replyToken) {
  const flexMessage = {
    "type": "flex",
    "altText": "請選擇要記錄支出或收入",
    "contents": {
      "type": "bubble",
      "body": {
        "type": "box",
        "layout": "vertical",
        "contents": [
          { "type": "text", "text": "請選擇記帳類型", "weight": "bold", "size": "xl" }
        ]
      },
      "footer": {
        "type": "box",
        "layout": "vertical",
        "spacing": "sm",
        "contents": [
          {
            "type": "button",
            "style": "primary",
            "height": "sm",
            "action": { "type": "postback", "label": "支出", "data": "action=select_type&type=支出" },
            "color": "#DF6C4F"
          },
          {
            "type": "button",
            "style": "primary",
            "height": "sm",
            "action": { "type": "postback", "label": "收入", "data": "action=select_type&type=收入" },
             "color": "#4CAF50"
          }
        ]
      }
    }
  };
  replyMessage(replyToken, flexMessage);
}

/**
 * 根據使用者選擇的類型（支出/收入），從 Google Sheet 取得類別並以 Flex Message 呈現
 */
function askCategory(replyToken, type, userId) {
  const categories = getCategoriesFromSheet(userId);
  if (categories.length === 0) {
    replyMessage(replyToken, createTextMessage('目前沒有任何收支類別，請先使用「收支類別」指令新增。'));
    return;
  }

  const buttons = categories.map(category => ({
    "type": "button",
    "style": "link",
    "height": "sm",
    "action": {
      "type": "postback",
      "label": category,
      "data": `action=select_category&type=${encodeURIComponent(type)}&category=${encodeURIComponent(category)}`
    }
  }));

  const flexMessage = {
    "type": "flex",
    "altText": "請選擇類別",
    "contents": {
      "type": "bubble",
      "body": {
        "type": "box",
        "layout": "vertical",
        "contents": [
          { "type": "text", "text": `請選擇${type}類別`, "weight": "bold", "size": "xl" }
        ]
      },
      "footer": {
        "type": "box",
        "layout": "vertical",
        "spacing": "sm",
        "contents": buttons
      }
    }
  };
  replyMessage(replyToken, flexMessage);
}

/**
 * 將最終的收支資訊寫入 Google Sheet
 */
function recordTransaction(replyToken, userId, type, category, item, amount) {
  try {
    const sheet = getSheetByName('收支紀錄');
    const headers = ['時間', '使用者ID', '類型', '類別', '品項', '金額'];
    
    // 檢查並寫入標頭
    if (sheet.getLastRow() < 1) {
      sheet.appendRow(headers);
    }
    
    const timestamp = new Date();
    sheet.appendRow([timestamp, userId, type, category, item, amount]);

    const confirmationMessage = `✅ 已成功記錄：\n類型：${type}\n類別：${category}\n品項：${item}\n金額：${amount} 元`;
    replyMessage(replyToken, createTextMessage(confirmationMessage));
  } catch (error) {
    Logger.log('寫入收支紀錄時發生錯誤: ' + error);
    replyMessage(replyToken, createTextMessage('記錄失敗，請稍後再試。'));
  }
}


/****************************************************************
 * 2. 收支類別管理功能
 ****************************************************************/

/**
 * 顯示目前的類別，並提供新增/刪除選項
 */
function manageCategories(replyToken, userId) {
  let categories = getCategoriesFromSheet(userId);
  
  let messageText = "目前的收支類別有：\n";
  if (categories.length > 0) {
      messageText += categories.join('\n');
  } else {
      messageText = "您目前沒有任何自訂類別。";
  }

  const flexMessage = {
    "type": "flex",
    "altText": "管理收支類別",
    "contents": {
      "type": "bubble",
      "body": {
        "type": "box",
        "layout": "vertical",
        "contents": [
          { "type": "text", "text": "管理收支類別", "weight": "bold", "size": "xl" },
          { "type": "separator", "margin": "md" },
          { "type": "text", "text": messageText, "wrap": true, "margin": "md" }
        ]
      },
      "footer": {
        "type": "box",
        "layout": "horizontal",
        "spacing": "sm",
        "contents": [
          {
            "type": "button", "style": "primary", "color": "#5C6BC0",
            "action": { "type": "postback", "label": "新增", "data": "action=manage_category&do=add" }
          },
          {
            "type": "button", "style": "primary", "color": "#E57373",
            "action": { "type": "postback", "label": "刪除", "data": "action=manage_category&do=delete" }
          }
        ]
      }
    }
  };
  replyMessage(replyToken, flexMessage);
}

/**
 * 新增一個類別
 */
function addNewCategory(replyToken, userId, newCategory) {
  try {
    const sheet = getSheetByName('收支類別');
    const allData = sheet.getDataRange().getValues();
    const userCategories = allData.filter(row => row[0] === userId).map(row => row[1]);

    if (userCategories.includes(newCategory)) {
        replyMessage(replyToken, createTextMessage(`類別「${newCategory}」已經存在。`));
        return;
    }

    sheet.appendRow([userId, newCategory]);
    replyMessage(replyToken, createTextMessage(`已成功新增類別：「${newCategory}」`));
  } catch (error) {
    Logger.log('新增類別時發生錯誤: ' + error);
    replyMessage(replyToken, createTextMessage('新增失敗，請稍後再試。'));
  }
}

/**
 * 詢問要刪除哪個類別
 */
function askDeleteCategory(replyToken, userId) {
  const categories = getCategoriesFromSheet(userId);
  if (categories.length === 0) {
    replyMessage(replyToken, createTextMessage('目前沒有可刪除的類別。'));
    return;
  }
  
  const buttons = categories.map(category => ({
    "type": "button",
    "style": "link",
    "height": "sm",
    "color": "#C90000",
    "action": {
      "type": "postback",
      "label": `刪除「${category}」`,
      "data": `action=delete_category&category=${encodeURIComponent(category)}`
    }
  }));

  const flexMessage = {
    "type": "flex",
    "altText": "請選擇要刪除的類別",
    "contents": {
      "type": "bubble",
      "body": {
        "type": "box",
        "layout": "vertical",
        "contents": [
          { "type": "text", "text": "請選擇要刪除的類別", "weight": "bold", "size": "xl" }
        ]
      },
      "footer": {
        "type": "box",
        "layout": "vertical",
        "spacing": "sm",
        "contents": buttons
      }
    }
  };
  replyMessage(replyToken, flexMessage);
}


/**
 * 刪除一個指定的類別
 */
function deleteCategory(replyToken, userId, categoryToDelete) {
  try {
    const sheet = getSheetByName('收支類別');
    const data = sheet.getDataRange().getValues();
    let rowDeleted = false;
    
    // 從後往前刪除，避免索引錯亂
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][0] === userId && data[i][1] === categoryToDelete) {
        sheet.deleteRow(i + 1);
        rowDeleted = true;
        break; 
      }
    }
    
    if (rowDeleted) {
      replyMessage(replyToken, createTextMessage(`已成功刪除類別：「${categoryToDelete}」`));
    } else {
      replyMessage(replyToken, createTextMessage(`找不到類別：「${categoryToDelete}」`));
    }
  } catch (error) {
    Logger.log('刪除類別時發生錯誤: ' + error);
    replyMessage(replyToken, createTextMessage('刪除失敗，請稍後再試。'));
  }
}


/****************************************************************
 * 3. 收支報表與圖表功能
 ****************************************************************/

/**
 * 詢問使用者要產生的報表區間
 */
function askReportPeriod(replyToken) {
  const flexMessage = {
    "type": "flex",
    "altText": "請選擇報表區間",
    "contents": {
      "type": "bubble",
      "body": {
        "type": "box",
        "layout": "vertical",
        "contents": [
          { "type": "text", "text": "請選擇報表區間", "weight": "bold", "size": "xl" }
        ]
      },
      "footer": {
        "type": "box",
        "layout": "vertical",
        "spacing": "sm",
        "contents": [
          {
            "type": "button", "style": "primary",
            "action": { "type": "postback", "label": "今日報表", "data": "action=generate_report&period=today" }
          },
          {
            "type": "button", "style": "primary",
            "action": { "type": "postback", "label": "本月報表", "data": "action=generate_report&period=month" }
          },
          {
            "type": "button", "style": "primary",
            "action": { "type": "postback", "label": "今年報表", "data": "action=generate_report&period=year" }
          }
        ]
      }
    }
  };
  replyMessage(replyToken, flexMessage);
}

/**
 * 產生報表、圖表並回傳
 */
function generateReport(replyToken, period, userId) {
  const sheet = getSheetByName('收支紀錄');
  if (sheet.getLastRow() <= 1) {
    replyMessage(replyToken, createTextMessage('目前沒有任何收支紀錄可供分析。'));
    return;
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const records = data.slice(1);

  const now = new Date();
  let startDate, endDate;
  let reportTitle = '';

  switch (period) {
    case 'today':
      startDate = new Date(now.setHours(0, 0, 0, 0));
      endDate = new Date(now.setHours(23, 59, 59, 999));
      reportTitle = '今日收支報表';
      break;
    case 'month':
      startDate = new Date(now.getFullYear(), now.getMonth(), 1);
      endDate = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59, 999);
      reportTitle = '本月收支報表';
      break;
    case 'year':
      startDate = new Date(now.getFullYear(), 0, 1);
      endDate = new Date(now.getFullYear(), 11, 31, 23, 59, 59, 999);
      reportTitle = '今年收支報表';
      break;
  }
  
  const timeCol = headers.indexOf('時間');
  const userCol = headers.indexOf('使用者ID');
  const typeCol = headers.indexOf('類型');
  const categoryCol = headers.indexOf('類別');
  const amountCol = headers.indexOf('金額');

  const filteredRecords = records.filter(row => {
    const recordDate = new Date(row[timeCol]);
    return row[userCol] === userId && recordDate >= startDate && recordDate <= endDate;
  });

  if (filteredRecords.length === 0) {
    replyMessage(replyToken, createTextMessage(`您在指定區間內沒有任何收支紀錄。`));
    return;
  }

  let totalIncome = 0;
  let totalExpense = 0;
  const incomeByCategory = {};
  const expenseByCategory = {};

  filteredRecords.forEach(row => {
    const type = row[typeCol];
    const category = row[categoryCol];
    const amount = parseFloat(row[amountCol]);

    if (type === '收入') {
      totalIncome += amount;
      incomeByCategory[category] = (incomeByCategory[category] || 0) + amount;
    } else if (type === '支出') {
      totalExpense += amount;
      expenseByCategory[category] = (expenseByCategory[category] || 0) + amount;
    }
  });

  // 建立報表文字
  let reportText = `${reportTitle}\n`;
  reportText += `--------------------\n`;
  reportText += `【收入】\n`;
  if (Object.keys(incomeByCategory).length > 0) {
    for (const category in incomeByCategory) {
      reportText += `${category}: ${incomeByCategory[category]} 元\n`;
    }
  } else {
    reportText += `無\n`;
  }
  reportText += `總收入: ${totalIncome} 元\n\n`;

  reportText += `【支出】\n`;
   if (Object.keys(expenseByCategory).length > 0) {
    for (const category in expenseByCategory) {
      reportText += `${category}: ${expenseByCategory[category]} 元\n`;
    }
  } else {
    reportText += `無\n`;
  }
  reportText += `總支出: ${totalExpense} 元\n`;
  reportText += `--------------------\n`;
  reportText += `結餘: ${totalIncome - totalExpense} 元`;
  
  const messages = [createTextMessage(reportText)];

  // 建立並上傳圖表
  try {
    const incomeChartUrl = createPieChart(incomeByCategory, '收入圓餅圖');
    if(incomeChartUrl) messages.push(createImageMessage(incomeChartUrl));
    
    const expenseChartUrl = createPieChart(expenseByCategory, '支出圓餅圖');
    if(expenseChartUrl) messages.push(createImageMessage(expenseChartUrl));
  } catch (err) {
    Logger.log("圖表生成失敗: " + err);
    // 即使圖表失敗，仍然發送文字報表
  }

  replyMessage(replyToken, messages);
}

/**
 * 使用 Google Charts Service 建立圓餅圖並上傳至 Google Drive
 * @param {Object} data - { 類別1: 金額1, 類別2: 金額2, ... }
 * @param {string} title - 圖表標題
 * @returns {string|null} - 公開的圖片網址，或 null
 */
function createPieChart(data, title) {
  if (Object.keys(data).length === 0) {
    return null; // 沒有資料就不建立圖表
  }

  const dataTable = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Category')
    .addColumn(Charts.ColumnType.NUMBER, 'Amount');
  
  for (const category in data) {
    dataTable.addRow([category, data[category]]);
  }

  const chart = Charts.newPieChart()
    .setDataTable(dataTable)
    .setTitle(title)
    .setOption('titleTextStyle', { color: '#333', fontSize: 20 })
    .setOption('legend', { position: 'right', textStyle: { color: 'black', fontSize: 16 } })
    .setOption('pieSliceText', 'value')
    .setOption('width', 800)
    .setOption('height', 500)
    .build();

  const chartBlob = chart.getAs('image/png');
  
  // 將圖表儲存至 Google Drive
  const folderName = "LINE_Bot_Charts";
  let folders = DriveApp.getFoldersByName(folderName);
  let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  
  const fileName = `${title}_${new Date().getTime()}.png`;
  const file = folder.createFile(chartBlob).setName(fileName);
  
  // 設定檔案為公開讀取
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  // 取得公開網址
  return `https://drive.google.com/uc?id=${file.getId()}`;
}


/****************************************************************
 * 4. 匯率查詢功能
 ****************************************************************/
/**
 * 詢問使用者要查詢的外幣
 */
function askCurrency(replyToken) {
  const currencies = [
      {label: '美金 (USD)', data: 'USD'},
      {label: '日圓 (JPY)', data: 'JPY'},
      {label: '歐元 (EUR)', data: 'EUR'},
      {label: '英鎊 (GBP)', data: 'GBP'},
      {label: '人民幣 (CNY)', data: 'CNY'}
  ];

  const buttons = currencies.map(c => ({
      "type": "button",
      "style": "link",
      "height": "sm",
      "action": {
          "type": "postback",
          "label": c.label,
          "data": `action=query_exchange_rate&currency=${c.data}`
      }
  }));

  const flexMessage = {
    "type": "flex",
    "altText": "請選擇要查詢的匯率",
    "contents": {
      "type": "bubble",
      "body": {
        "type": "box",
        "layout": "vertical",
        "contents": [
          { "type": "text", "text": "請選擇要查詢的匯率", "weight": "bold", "size": "xl" }
        ]
      },
      "footer": {
        "type": "box",
        "layout": "vertical",
        "spacing": "sm",
        "contents": buttons
      }
    }
  };
  replyMessage(replyToken, flexMessage);
}

/**
 * 從台灣銀行網站抓取匯率資料
 */
function getExchangeRate(replyToken, currency) {
  try {
    const url = 'https://rate.bot.com.tw/xrt?Lang=zh-TW';
    const response = UrlFetchApp.fetch(url).getContentText();
    const $ = Cheerio.load(response);

    const rateRow = $(`.currency:contains("${currency}")`).closest('tr');
    
    if (rateRow.length === 0) {
      replyMessage(replyToken, createTextMessage(`找不到 ${currency} 的匯率資訊。`));
      return;
    }

    const cashBuy = rateRow.find('td[data-table="本行現金買入"]').text().trim();
    const cashSell = rateRow.find('td[data-table="本行現金賣出"]').text().trim();
    const spotBuy = rateRow.find('td[data-table="本行即期買入"]').text().trim();
    const spotSell = rateRow.find('td[data-table="本行即期賣出"]').text().trim();
    const currencyName = rateRow.find('.currency .visible-phone').text().trim();
    const updateTime = $('span.time').text().trim();

    const message = `查詢幣別：${currencyName} (${currency})\n` +
                  `更新時間：${updateTime}\n` +
                  `--------------------\n` +
                  `現金買入: ${cashBuy}\n` +
                  `現金賣出: ${cashSell}\n` +
                  `即期買入: ${spotBuy}\n` +
                  `即期賣出: ${spotSell}`;

    replyMessage(replyToken, createTextMessage(message));
  } catch (error) {
    Logger.log('抓取匯率失敗: ' + error);
    replyMessage(replyToken, createTextMessage('查詢匯率時發生錯誤，請稍後再試。'));
  }
}

/****************************************************************
 * 5. 台股查詢功能
 ****************************************************************/
/**
 * 從台灣證券交易所網站抓取股票資訊
 */
function getStockPrice(replyToken, stockCode) {
    try {
        const url = `https://mis.twse.com.tw/stock/api/getStockInfo.jsp?ex_ch=tse_${stockCode}.tw&json=1&delay=0`;
        const response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
        const data = JSON.parse(response.getContentText());

        if (!data.msgArray || data.msgArray.length === 0) {
            replyMessage(replyToken, createTextMessage(`查無股票代碼 ${stockCode} 的資訊，請確認代碼是否正確。`));
            return;
        }

        const stockInfo = data.msgArray[0];
        const stockName = stockInfo.n; // 公司簡稱
        const currentPrice = stockInfo.z; // 成交價
        const change = parseFloat(stockInfo.z) - parseFloat(stockInfo.y); // 漲跌 = 成交價 - 昨收價
        const changePercent = ((change / parseFloat(stockInfo.y)) * 100).toFixed(2); // 漲跌幅
        const openPrice = stockInfo.o; // 開盤價
        const highPrice = stockInfo.h; // 最高價
        const lowPrice = "l" in stockInfo ? stockInfo.l : 'N/A';   // 最低價
        const volume = stockInfo.v; // 成交量

        let changeSymbol = change > 0 ? '🔼' : (change < 0 ? '🔽' : '⏹️');
        
        const message = `📈 ${stockName} (${stockCode})\n` +
                      `--------------------\n` +
                      `目前股價: ${currentPrice}\n` +
                      `漲跌: ${changeSymbol} ${change.toFixed(2)} (${changePercent}%)\n` +
                      `開盤價: ${openPrice}\n` +
                      `最高價: ${highPrice}\n` +
                      `最低價: ${lowPrice}\n` +
                      `成交量: ${volume} 張`;

        replyMessage(replyToken, createTextMessage(message));
    } catch (error) {
        Logger.log('查詢股價失敗: ' + error);
        replyMessage(replyToken, createTextMessage('查詢股價時發生錯誤，或股票代碼不存在。'));
    }
}


/****************************************************************
 * Google Sheet & 使用者狀態 & 工具函式
 ****************************************************************/

/**
 * 根據名稱取得 Google Sheet 中的工作表，若不存在則建立
 * @param {string} name - 工作表名稱
 * @returns {Sheet} - Google Apps Script 的 Sheet 物件
 */
function getSheetByName(name) {
  const ss = SpreadsheetApp.openById(GOOGLE_SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // 初始化 Sheet
    if (name === '收支紀錄') {
      sheet.appendRow(['時間', '使用者ID', '類型', '類別', '品項', '金額']);
    } else if (name === '收支類別') {
      sheet.appendRow(['使用者ID', '類別']);
    }
  }
  return sheet;
}

/**
 * 從 Sheet 中取得指定使用者的類別列表，如果該使用者沒有任何類別，則建立預設類別
 * @param {string} userId - 使用者 ID
 * @returns {Array<string>} - 類別字串陣列
 */
function getCategoriesFromSheet(userId) {
  const sheet = getSheetByName('收支類別');
  if (sheet.getLastRow() <= 1 && sheet.getLastColumn() <=1) { // 處理完全空白的狀況
      sheet.getRange(1, 1, 1, 2).setValues([['使用者ID', '類別']]);
  }
  const allData = sheet.getDataRange().getValues();
  const userCategories = allData.filter(row => row[0] === userId).map(row => row[1]);

  // 如果使用者沒有任何類別，則新增預設值
  if (userCategories.length === 0) {
    const defaultCategories = ['飲食', '交通', '購物', '娛樂', '工作', '其他'];
    const rowsToAdd = defaultCategories.map(cat => [userId, cat]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 2).setValues(rowsToAdd);
    return defaultCategories;
  }
  
  return userCategories;
}


/**
 * 使用 PropertiesService 儲存使用者的對話狀態
 */
function setUserState(userId, state) {
  PropertiesService.getUserProperties().setProperty(userId, JSON.stringify(state));
}

/**
 * 取得使用者的對話狀態
 */
function getUserState(userId) {
  const state = PropertiesService.getUserProperties().getProperty(userId);
  return state ? JSON.parse(state) : {};
}

/**
 * 清除使用者的對話狀態
 */
function clearUserState(userId) {
  PropertiesService.getUserProperties().deleteProperty(userId);
}


/**
 * 解析 Postback data (e.g., "action=add&item=milk" -> {action: "add", item: "milk"})
 * 內建 Polyfill 以解決 GAS 環境中沒有 URLSearchParams 的問題
 */
function parsePostbackData(dataString) {
  const params = {};
  const pairs = dataString.split('&');
  for (let i = 0; i < pairs.length; i++) {
    const pair = pairs[i].split('=');
    if (pair[0]) {
      params[decodeURIComponent(pair[0])] = decodeURIComponent(pair[1] || '');
    }
  }
  return params;
}

/****************************************************************
 * LINE Message API 傳送與格式化函式
 ****************************************************************/

/**
 * 回覆訊息給 LINE 使用者
 * @param {string} replyToken - 回覆用的 Token
 * @param {Object|Array<Object>} messages - 單一或多個 LINE Message 物件
 */
function replyMessage(replyToken, messages) {
  if (!Array.isArray(messages)) {
    messages = [messages];
  }
  const url = 'https://api.line.me/v2/bot/message/reply';
  const payload = {
    'replyToken': replyToken,
    'messages': messages
  };
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': {
      'Authorization': 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN
    },
    'payload': JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options);
}

/**
 * 建立一個標準的文字訊息物件
 */
function createTextMessage(text) {
  return { 'type': 'text', 'text': text };
}

/**
 * 建立一個標準的圖片訊息物件
 */
function createImageMessage(imageUrl) {
  return {
    'type': 'image',
    'originalContentUrl': imageUrl,
    'previewImageUrl': imageUrl
  };
}

/****************************************************************
 * Cheerio Library (用於網頁爬蟲)
 * 這是為了解析台銀匯率網頁而引入的外部函式庫
 * 如果您的專案沒有安裝，請手動加入。
 * 部署方式：在 Apps Script 編輯器中，點擊「程式庫」旁邊的 + 號，
 * 輸入腳本 ID: 1ReeQ6WO8kKNxagqrJnOf29QhBOcrg5NQD_bA7XwnpiKqA9jL00g9vA2I
 * 選擇最新版本後加入。
 ****************************************************************/
// Cheerio library will be added via Apps Script libraries.
