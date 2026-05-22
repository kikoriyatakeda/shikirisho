/**
 * ★【ステップ1】権限の承認ダイアログを強制表示する関数★
 */
function triggerAuthorization() {
  DriveApp.getRootFolder();
  DocumentApp.create("permission_trigger_dummy");
}

/**
 * ★【ステップ2】権限が取れたか確認するテスト関数★
 */
function testDrivePermission() {
  try {
    const blob = Utilities.newBlob("test", "text/plain", "test_permission.txt");
    const file = Drive.Files.insert({title: "test_permission.txt"}, blob);
    Drive.Files.remove(file.id);
    console.log("【テスト成功】Drive APIの完全な権限が取得できました！");
  } catch (e) {
    console.error("【テスト失敗】" + e.toString());
  }
}

// ========== 修正ポイント: スプレッドシートIDを定数化 ==========
const MASTER_SS_ID = '1-HV-cb7tPiOvTD4nzKnlhfmPnfr59Kq2qxNNsSlblWU';      // マスタ用（所有者マスタの移動先）
const INVOICE_SS_ID = '1hlzou-dX-hBmUwryu1YJgme4_IOiP4V8CswMQc3YXac';   // 仕切り書用（既存）

/**
 * 初期化データ（現場リストと所有者マスタ）をまとめて取得する
 */
function getInitialData() {
  return {
    genbaList: getGenbaList(),
    ownerMaster: getOwnerMasterList()
  };
}

/**
 * 現場マスタから「施業地区 + 現場名」のリストを取得する
 */
function getGenbaList() {
  try {
    const ss = SpreadsheetApp.openById(MASTER_SS_ID);
    const sheet = ss.getSheetByName("現場マスタ");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const chikuIdx = headers.indexOf("施業地区");
    const nameIdx = headers.indexOf("現場名");
    
    if (chikuIdx === -1 || nameIdx === -1) return [];
    
    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][nameIdx]) {
        const chiku = data[i][chikuIdx] ? data[i][chikuIdx] + " " : "";
        list.push(chiku + data[i][nameIdx]);
      }
    }
    return list;
  } catch (e) {
    console.error("マスタ取得エラー", e);
    return [];
  }
}

/**
 * 所有者マスタから「現場名 - 所有者名」のペアを取得する
 */
function getOwnerMasterList() {
  try {
    const ss = SpreadsheetApp.openById(MASTER_SS_ID);
    let sheet = ss.getSheetByName("所有者マスタ");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    return data.slice(1).map(row => ({
      genba: row[0],
      owner: row[1]
    }));
  } catch (e) {
    return [];
  }
}

/**
 * Google Drive の OCR機能 を使用して画像を解析する
 * （自動リトライ機能追加版）
 */
function analyzeDocumentImage(base64Image, dummyPrompt) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Image), MimeType.JPEG, "temp_ocr.jpg");
    
    let fileId;
    let retryCount = 0;
    const maxRetries = 3; // 最大3回再試行する
    let success = false;

    // OCR処理の実行とエラー時のリトライループ
    while (!success && retryCount <= maxRetries) {
      try {
        const resource = { title: "temp_ocr_doc" };
        const file = Drive.Files.insert(resource, blob, {ocr: true, ocrLanguage: 'ja'});
        fileId = file.id;
        success = true; // 成功したらループを抜ける
      } catch (apiError) {
        const errorStr = apiError.toString();
        // 利用制限エラーの場合のみリトライする
        if (errorStr.includes("User rate limit exceeded for OCR") || errorStr.includes("Rate Limit Exceeded")) {
          retryCount++;
          if (retryCount > maxRetries) {
            return JSON.stringify({ 
              error: "【OCR利用制限エラー】\n裏側で複数回自動再試行を行いましたが、制限が解除されませんでした。\n恐れ入りますが、数分〜数十分ほど時間を置いてから再度お試しください。"
            });
          }
          // 待機する（1回目:3秒, 2回目:6秒, 3回目:9秒）
          Utilities.sleep(retryCount * 3000);
        } else {
          // 別のエラーの場合は即座にエラーを返す
          return JSON.stringify({ 
            error: "【Drive API エラー】\n" + errorStr + "\n\n(GASエディタで triggerAuthorization を実行して権限を許可してください)"
          });
        }
      }
    }

    const doc = DocumentApp.openById(fileId);
    const text = doc.getBody().getText();
    
    Drive.Files.remove(fileId);

    return parseOcrText(text);

  } catch (e) {
    return JSON.stringify({ error: "OCR解析エラー: " + e.toString() });
  }
}

/**
 * OCRテキストの解析ロジック
 */
function parseOcrText(text) {
  const lines = text.split('\n');
  let summaryData = [];
  let dateStr = "";
  let periodStr = "";
  let extractedYear = 0;

  // 1. 日付の抽出（和暦から西暦への変換）
  const dateMatch = text.match(/(令和|平成)?\s*([元0-9０-９]+)\s*年\s*([0-9０-９]+)\s*月\s*([0-9０-９]+)\s*日/);
  if (dateMatch) {
    let nengo = dateMatch[1];
    let yearText = dateMatch[2].replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
    let year = yearText === "元" ? 1 : parseInt(yearText) || 1;
    let month = parseInt(dateMatch[3].replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)));
    let day = parseInt(dateMatch[4].replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)));
    
    if (nengo === "令和" || !nengo) year += 2018; 
    else if (nengo === "平成") year += 1988;
    
    extractedYear = year;
    dateStr = `${year}/${month.toString().padStart(2, '0')}/${day.toString().padStart(2, '0')}`;
  }

  // 2. 期間の抽出（パターンを大幅に拡張・強化）
  // 4/9 ～ 4/22, 2026/04/09 - 2026/04/22, R6.4.9 ~ R6.4.22, 1月1日～1月31日 などに対応
  const periodMatch = text.match(/(?:([0-9０-９]{1,4})\s*[\/\.年／]\s*)?([0-9０-９]{1,2})\s*[\/\.月／]\s*([0-9０-９]{1,2})\s*日?\s*[～~－-]\s*(?:([0-9０-９]{1,4})\s*[\/\.年／]\s*)?([0-9０-９]{1,2})\s*[\/\.月／]\s*([0-9０-９]{1,2})\s*日?/);
  
  if (periodMatch) {
    // 全角数字を半角に変換
    let y1Str = periodMatch[1] ? periodMatch[1].replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)) : null;
    let m1Str = periodMatch[2].replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
    let d1Str = periodMatch[3].replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
    
    let y2Str = periodMatch[4] ? periodMatch[4].replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)) : null;
    let m2Str = periodMatch[5].replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
    let d2Str = periodMatch[6].replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));

    let y1 = y1Str ? parseInt(y1Str) : extractedYear;
    let m1 = parseInt(m1Str);
    let d1 = parseInt(d1Str);
    
    let y2 = y2Str ? parseInt(y2Str) : (y1 || extractedYear);
    let m2 = parseInt(m2Str);
    let d2 = parseInt(d2Str);

    // 2桁年の補正（例: 26 -> 2026）
    if (y1 > 0 && y1 < 100) y1 += 2000;
    if (y2 > 0 && y2 < 100) y2 += 2000;

    // ▼ 年またぎ自動補正ロジックの追加
    // 年の指定がなく、基準年の設定が有効な場合のみ補正を適用
    if (!y1Str && !y2Str && extractedYear > 0) {
      if (m1 > m2) {
        // 例: 開始月が12月、終了月が1月などの場合
        // 開始日(y1)は基準年前年、終了日(y2)は基準年となるようセット
        y1 = extractedYear - 1;
        y2 = extractedYear;
      }
    }

    let start = y1 > 0 ? `${y1}/${m1.toString().padStart(2, '0')}/${d1.toString().padStart(2, '0')}` : `${m1}/${d1}`;
    let end = y2 > 0 ? `${y2}/${m2.toString().padStart(2, '0')}/${d2.toString().padStart(2, '0')}` : `${m2}/${d2}`;
    
    // 必ず「～」で繋いだフォーマットにして返す
    periodStr = `${start} ～ ${end}`;
  }

  // 3. 所有者候補の抽出（補助的に3文字以上の文字列を出す）
  let ownerCandidates = [];
  const noiseKeywords = ["仕切書", "計算書", "明細書", "精算書", "売上", "消費税", "手数料", "合計", "材積", "単価", "金額", "樹種", "現場", "電話", "TEL", "口座"];
  for (let i = 0; i < lines.length; i++) {
    let cleanLine = lines[i].replace(/\s+/g, ' ').trim();
    if (cleanLine.length < 3) continue;
    if (cleanLine.match(/[0-9０-９]{3,}/) || cleanLine.match(/[0-9０-９]+\.[0-9０-９]+/)) continue;
    if (noiseKeywords.some(kw => cleanLine.includes(kw))) continue;
    let name = cleanLine.replace(/[:：]/g, '').trim();
    if (name.length >= 3 && !ownerCandidates.includes(name)) ownerCandidates.push(name);
  }

  // ★ ドロップダウン候補用の全数字収集
  let globalVolumeCandidates = new Set([0]);
  let globalPriceCandidates = new Set([0]);
  for (let line of lines) {
    let nums = line.replace(/,/g, '').match(/\d+(\.\d+)?/g);
    if (nums) {
      nums.forEach(numStr => {
        let num = parseFloat(numStr);
        if (numStr.includes('.')) globalVolumeCandidates.add(num);
        else if (num >= 10) globalPriceCandidates.add(num);
      });
    }
  }

  // 4. 新・樹種マッピング & 厳格グループ抽出
  const speciesWords = ["スギ", "ヒノキ", "マツ", "杉", "桧", "その他"];
  const endWords = ["売上", "合計", "精算", "消費税", "総計", "差引"];

  // ★ 抽出結果を格納するマップ（初期状態で「スギ」「ヒノキ」「その他」の3行を必ず確保）
  let speciesResult = {
    "スギ": { volume: 0, avgPrice: 0 },
    "ヒノキ": { volume: 0, avgPrice: 0 },
    "その他": { volume: 0, avgPrice: 0 }
  };

  // 誤反応を避けるため、まず「最初に樹種（スギやヒノキ等）が出現する行」を探す
  let firstSpeciesIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    for (let sp of speciesWords) {
      if (lines[i].includes(sp)) {
        firstSpeciesIdx = i;
        break;
      }
    }
    if (firstSpeciesIdx !== -1) break;
  }

  if (firstSpeciesIdx !== -1) {
    let currentTargetSp = null;
    let targetGroupNums = [];

    for (let i = firstSpeciesIdx; i < lines.length; i++) {
      let line = lines[i];

      // 集計の終わりキーワードを見つけたら、そこでデータ取得をクローズ
      if (endWords.some(w => line.includes(w))) {
        if (currentTargetSp && targetGroupNums.length > 0) {
          processGroupNumbers(currentTargetSp, targetGroupNums, speciesResult);
        }
        currentTargetSp = null;
        targetGroupNums = [];
        break;
      }

      // 樹種キーワードが含まれているかチェック
      let foundSp = null;
      for (let sp of speciesWords) {
        if (line.includes(sp)) {
          foundSp = (sp === "杉") ? "スギ" : (sp === "桧") ? "ヒノキ" : sp;
          break;
        }
      }

      // 新しい樹種が見つかった場合、前の樹種データを集計・確定する
      if (foundSp) {
        if (currentTargetSp && targetGroupNums.length > 0) {
          processGroupNumbers(currentTargetSp, targetGroupNums, speciesResult);
        }
        currentTargetSp = foundSp;
        targetGroupNums = [];
      }

      // 現在ターゲット中の樹種があれば、同じ行の数値をすべて集める
      if (currentTargetSp) {
        let nums = line.replace(/,/g, '').match(/\d+(\.\d+)?|[-ー－—_]{2,}/g);
        if (nums) {
          nums.forEach(n => {
            if (n.match(/[-ー－—_]/)) targetGroupNums.push(0);
            else targetGroupNums.push(parseFloat(n));
          });
        }
      }
    }

    // ループ終了後の残処理
    if (currentTargetSp && targetGroupNums.length > 0) {
      processGroupNumbers(currentTargetSp, targetGroupNums, speciesResult);
    }
  }

  // 最終的に、確保されたデフォルト3樹種の枠をそのまま `summaryData` にマッピング
  summaryData = Object.keys(speciesResult).map(spName => {
    return {
      species: spName,
      volume: speciesResult[spName].volume,
      avgPrice: speciesResult[spName].avgPrice
    };
  });

  return JSON.stringify({
    date: dateStr,
    period: periodStr, 
    ownerCandidates: ownerCandidates, 
    summaryData: summaryData,
    volumeCandidates: Array.from(globalVolumeCandidates).sort((a,b)=>a-b),
    priceCandidates: Array.from(globalPriceCandidates).sort((a,b)=>a-b),
    rawOcrText: text
  });
}

/**
 * 樹種グループから高精度に数値ペア（材積、単価）を解くサブ関数
 */
function processGroupNumbers(speciesName, nums, speciesResult) {
  let volumes = [];
  let prices = [];

  // 数値の特徴で仕分け
  nums.forEach(n => {
    if (n === 0) {
      volumes.push(0);
      prices.push(0);
    } else if (!Number.isInteger(n) || n < 100) {
      // 材積候補（小数、または100未満の極めて小さい整数）
      volumes.push(n);
    } else if (Number.isInteger(n) && n >= 100) {
      // 単価・金額候補（100以上の大きな整数）
      prices.push(n);
    }
  });

  let assignedV = 0;
  let assignedP = 0;
  let success = false;

  // 同一の樹種グループ内で検算（材積 × 単価 ＝ 小計）パズル
  for (let i = 0; i < volumes.length; i++) {
    let v = volumes[i];
    if (v === 0) continue;

    for (let j = 0; j < prices.length; j++) {
      let p = prices[j];
      let expectedSub = Math.round(v * p);

      // 小計が価格リストの別インデックスにあるか探す（誤差5円以内）
      let sIndex = prices.findIndex((val, idx) => idx !== j && Math.abs(val - expectedSub) <= 5);

      if (sIndex !== -1) {
        assignedV = v;
        assignedP = p;
        success = true;
        break;
      }
    }
    if (success) break;
  }

  // 完全にパズルが解けなかった場合のフォールバック（最も手前にある数字をセット）
  if (!success) {
    assignedV = volumes[0] || 0;
    assignedP = prices[0] || 0;
  }

  // マップにデータを格納・上書き
  speciesResult[speciesName] = {
    volume: assignedV,
    avgPrice: Math.round(assignedP)
  };
}

/**
 * 抽出データをスプレッドシートに保存 & マスタ自動更新
 */
function saveToSpreadsheet(jsonString) {
  try {
    if (!jsonString) throw new Error("データが空です。");
    const data = JSON.parse(jsonString);
    
    const invoiceSs = SpreadsheetApp.openById(INVOICE_SS_ID);
    
    // 期間（period）を開始日と終了日に分割する処理
    let startDate = "-";
    let endDate = "-";
    if (data.period && data.period !== "-") {
      const parts = data.period.split(/～|~|－|-/);
      if (parts.length >= 2) {
        startDate = parts[0].trim();
        endDate = parts.slice(1).join('-').trim(); // 複数ハイフン対策
      } else {
        startDate = data.period.trim();
      }
    }

    // 1. 仕切り書への保存（13列構成）
    let sheet = invoiceSs.getSheetByName("仕切り書");
    if (!sheet) {
      sheet = invoiceSs.insertSheet("仕切り書");
      sheet.appendRow(["保存日時", "現場名", "事業区分", "所有者", "日付", "開始日", "終了日", "樹種", "材積", "単価", "小計", "合計材積", "合計金額"]);
      sheet.getRange(1, 1, 1, 13).setBackground("#059669").setFontColor("white").setFontWeight("bold");
    } else {
      const existingData = sheet.getDataRange().getDisplayValues();
      for (let i = 1; i < existingData.length; i++) {
        // 重複チェック
        if (existingData[i][1] === data.genbaName && 
            existingData[i][3] === data.ownerName && 
            existingData[i][4] === data.date) {
           throw new Error("DuplicateData: 同じデータが既に登録されています。");
        }
      }
    }
    
    const now = new Date();
    // スプレッドシートの13列にピタリと合うように配列を組む
    const rows = (data.summaryData || []).map(item => [
      now, 
      data.genbaName || "-", 
      data.businessType || "-", 
      data.ownerName || "-", 
      data.date || "-",     // 日付
      startDate,            // 開始日
      endDate,              // 終了日
      item.species || "", 
      item.volume || 0, 
      item.avgPrice || 0, 
      item.subtotal || 0, 
      data.totalVolume || 0, 
      data.totalAmount || 0  
    ]);
    if (rows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    // 2. 所有者マスタの自動更新
    if (data.genbaName && data.ownerName) {
      const masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
      updateOwnerMaster(masterSs, data.genbaName, data.ownerName);
    }

    return "Success";
  } catch (e) {
    return "Error: " + e.toString();
  }
}

/**
 * 所有者マスタに現場名と所有者のペアを登録
 */
function updateOwnerMaster(ss, genba, owner) {
  let masterSheet = ss.getSheetByName("所有者マスタ");
  if (!masterSheet) {
    masterSheet = ss.insertSheet("所有者マスタ");
    masterSheet.appendRow(["現場名", "所有者名"]);
    masterSheet.getRange(1, 1, 1, 2).setBackground("#059669").setFontColor("white").setFontWeight("bold");
  }
  const data = masterSheet.getDataRange().getValues();
  const exists = data.some(row => row[0] === genba && row[1] === owner);
  if (!exists) masterSheet.appendRow([genba, owner]);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('林建DX PDF解析')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}