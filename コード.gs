/**
 * Webアプリとしてアクセスしたときに表示する画面（タブ切り替え対応）
 */
function doGet(e) {
  var page = e.parameter.page;
  if (page === 'dashboard') {
    return HtmlService.createTemplateFromFile('dashboard').evaluate()
        .setTitle('高度分析ダッシュボード')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else if (page === 'dialog') {
    return HtmlService.createHtmlOutputFromFile('dialog')
        .setTitle('データ整理')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else {
    return HtmlService.createTemplateFromFile('index').evaluate()
        .setTitle('パタカラ データ分析システム')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function parseUploadedData(fileData) {
  if (typeof fileData === 'string' && fileData.indexOf('[') === 0) {
    return JSON.parse(fileData);
  }
  var base64 = fileData.split(',')[1];
  var blob = Utilities.newBlob(Utilities.base64Decode(base64), MimeType.CSV);
  var text = blob.getDataAsString("UTF-8");
  if (text.indexOf("\uFFFD") !== -1 || text.indexOf("ï¿½") !== -1) {
      text = blob.getDataAsString("Shift_JIS");
  }
  return Utilities.parseCsv(text);
}

function cleanText(text) {
  if (!text) return "";
  return String(text).replace(/[\s　\uFEFF\u200B]/g, '');
}

function normalizeDataArray(dataArray) {
  if (!dataArray || dataArray.length === 0) return dataArray;
  var targetLength = dataArray[0].length;
  return dataArray.map(function(row) {
    if (row.length === targetLength) return row;
    var newRow = [];
    for (var i = 0; i < targetLength; i++) {
      newRow.push(i < row.length ? row[i] : "");
    }
    return newRow;
  });
}

function getFilterDataFromUploadedFile(dataUrl) {
  var data = parseUploadedData(dataUrl);
  if (!data || data.length === 0) return null;
  var headers = data[0];
  var cleanHeaders = headers.map(cleanText);
  var idxGroup = cleanHeaders.findIndex(function(h) { return h.indexOf('連携グループ') !== -1 || h.indexOf('連携済みグループ') !== -1; });
  if (idxGroup === -1) return { error: "連携グループ of 列が見つかりません" };
  var groups = [];
  for (var i = 1; i < data.length; i++) {
    var g = data[i][idxGroup];
    var gClean = cleanText(g);
    // 空欄、"-"、"未設定"、"未登録" はデータ整理の除外対象リストに表示させない
    if (g && gClean !== "" && gClean !== "-" && gClean !== "未設定" && gClean !== "未登録" && groups.indexOf(g) === -1) {
      groups.push(g);
    }
  }
  return { groups: groups };
}

function savePinnedCharts(pinned) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('pinnedCharts', JSON.stringify(pinned));
}

function loadPinnedCharts() {
  var userProperties = PropertiesService.getUserProperties();
  var pinned = userProperties.getProperty('pinnedCharts');
  return pinned ? JSON.parse(pinned) : [];
}

function saveExcludedGroups(groups) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('excludedGroups', JSON.stringify(groups));
}

function loadExcludedGroups() {
  var userProperties = PropertiesService.getUserProperties();
  var groups = userProperties.getProperty('excludedGroups');
  return groups ? JSON.parse(groups) : [];
}

function getSheetList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return [];
  return ss.getSheets().map(function(s) { return s.getName(); }).reverse();
}

/**
 * 列ヘッダーから特定の指標（パタカラ）のインデックスを特定する（秒間回数列 N-Q を優先）
 */
function findPatakaIndex(headers, char) {
  var idx = headers.findIndex(function(h) { 
    return h.indexOf(char) !== -1 && (
      h.indexOf('秒') !== -1 || 
      h.indexOf('/') !== -1 || 
      h.indexOf('hz') !== -1 || 
      h.indexOf('Hz') !== -1 || 
      h.indexOf('HZ') !== -1
    ); 
  });
  
  if (idx === -1) {
    for (var i = headers.length - 1; i >= 0; i--) {
      if (headers[i].indexOf(char) !== -1) return i;
    }
  }
  return idx;
}

/**
 * データ整理：インポートデータを各除外条件に基づいて加工し、新規シートとして出力する
 */
function processUploadedData(formObject) {
  try {
    var rawData = parseUploadedData(formObject.fileData);
    if (!rawData || rawData.length === 0) return { error: "アップロードデータのパースに失敗しました。" };
    
    var rawHeaders = rawData[0];
    
    // ゴースト空白列（右側に無限に広がる空セル）の徹底切り落とし
    var lastValidColIdx = 0;
    for (var c = rawHeaders.length - 1; c >= 0; c--) {
      if (cleanText(rawHeaders[c]) !== "") {
        lastValidColIdx = c;
        break;
      }
    }
    
    var headers = rawHeaders.slice(0, lastValidColIdx + 1);
    var cleanHeaders = headers.map(cleanText);
    
    var idxName = cleanHeaders.findIndex(function(h) { return h.indexOf('ニックネーム') !== -1; });
    var idxDate = cleanHeaders.findIndex(function(h) { return h.indexOf('日付') !== -1 || h.indexOf('日時') !== -1; });
    var idxAge = cleanHeaders.findIndex(function(h) { return h.indexOf('年齢') !== -1; });
    var idxGroup = cleanHeaders.findIndex(function(h) { return h.indexOf('連携グループ') !== -1 || h.indexOf('連携済みグループ') !== -1; });
    var idxA4 = cleanHeaders.findIndex(function(h) { return h.indexOf('4文字平均') !== -1; });
    var idxA3 = cleanHeaders.findIndex(function(h) { return h.indexOf('3文字平均') !== -1; });
    
    var idxP = findPatakaIndex(cleanHeaders, 'パ');
    var idxT = findPatakaIndex(cleanHeaders, 'タ');
    var idxK = findPatakaIndex(cleanHeaders, 'カ');
    var idxR = findPatakaIndex(cleanHeaders, 'ラ');

    var excludeGroups = formObject.excludeGroups || [];
    
    var shouldExcludeMetric = function(valueStr, limitStr, type) {
      if (limitStr === "" || limitStr === undefined || limitStr === null) return false;
      var val = parseFloat(valueStr);
      var limit = parseFloat(limitStr);
      if (isNaN(val) || isNaN(limit)) return false;
      if (type === "lessThan") return val < limit;
      if (type === "lessThanOrEqual") return val <= limit;
      if (type === "greaterThanOrEqual") return val >= limit;
      if (type === "greaterThan") return val > limit;
      return false;
    };

    var filteredRows = [headers];
    var excludedCount = 0;
    
    for (var i = 1; i < rawData.length; i++) {
      var fullRow = rawData[i];
      if (!fullRow || fullRow.length === 0) continue;
      
      var row = fullRow.slice(0, lastValidColIdx + 1);

      // 【徹底除外ガードレール】ゴースト行は書き込まない
      if (idxName !== -1) {
        var nameVal = cleanText(row[idxName]);
        if (!nameVal) continue;
      }
      if (idxDate !== -1) {
        var dateVal = cleanText(row[idxDate]);
        if (!dateVal) continue;
      }
      
      var isExcluded = false;
      
      if (idxGroup !== -1 && excludeGroups.length > 0) {
        var groupVal = row[idxGroup];
        if (groupVal && excludeGroups.indexOf(groupVal) !== -1) isExcluded = true;
      }
      if (!isExcluded && idxAge !== -1) {
        var ageVal = row[idxAge];
        if (shouldExcludeMetric(ageVal, formObject.ageMinVal, formObject.ageMinType)) isExcluded = true;
        if (shouldExcludeMetric(ageVal, formObject.ageMaxVal, formObject.ageMaxType)) isExcluded = true;
        // 年齢が未入力（0または空）のユーザーを除外する（デフォルトON。年齢別分布タブの散布図等は元々この条件でしか描画できないため）
        if (!isExcluded && formObject.excludeEmptyAge) {
          var ageNum = parseFloat(ageVal);
          if (ageVal === "" || ageVal === null || ageVal === undefined || isNaN(ageNum) || ageNum === 0) isExcluded = true;
        }
      }
      
      // 口腔機能指標除外
      if (!isExcluded && idxA4 !== -1 && formObject.a4MinVal !== "") {
        if (shouldExcludeMetric(row[idxA4], formObject.a4MinVal, formObject.a4MinType)) isExcluded = true;
        if (shouldExcludeMetric(row[idxA4], formObject.a4MaxVal, formObject.a4MaxType)) isExcluded = true;
      }
      if (!isExcluded && idxA3 !== -1 && formObject.a3MinVal !== "") {
        if (shouldExcludeMetric(row[idxA3], formObject.a3MinVal, formObject.a3MinType)) isExcluded = true;
        if (shouldExcludeMetric(row[idxA3], formObject.a3MaxVal, formObject.a3MaxType)) isExcluded = true;
      }
      if (!isExcluded && idxP !== -1 && formObject.pMinVal !== "") {
        if (shouldExcludeMetric(row[idxP], formObject.pMinVal, formObject.pMinType)) isExcluded = true;
        if (shouldExcludeMetric(row[idxP], formObject.pMaxVal, formObject.pMaxType)) isExcluded = true;
      }
      if (!isExcluded && idxT !== -1 && formObject.tMinVal !== "") {
        if (shouldExcludeMetric(row[idxT], formObject.tMinVal, formObject.tMinType)) isExcluded = true;
        if (shouldExcludeMetric(row[idxT], formObject.tMaxVal, formObject.tMaxType)) isExcluded = true;
      }
      if (!isExcluded && idxK !== -1 && formObject.kMinVal !== "") {
        if (shouldExcludeMetric(row[idxK], formObject.kMinVal, formObject.kMinType)) isExcluded = true;
        if (shouldExcludeMetric(row[idxK], formObject.kMaxVal, formObject.kMaxType)) isExcluded = true;
      }
      if (!isExcluded && idxR !== -1 && formObject.rMinVal !== "") {
        if (shouldExcludeMetric(row[idxR], formObject.rMinVal, formObject.rMinType)) isExcluded = true;
        if (shouldExcludeMetric(row[idxR], formObject.rMaxVal, formObject.rMaxType)) isExcluded = true;
      }
      
      if (isExcluded) excludedCount++;
      else filteredRows.push(row);
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 蓄積した古い「抽出DB_xxxx」シートを自動整理 (ローテーション消去)
    // シート名の "抽出DB_MMdd_HHmm" から日時を読み取って新旧を判定する
    // (タブの並び順に依存すると、手動で並べ替えられていた場合などに誤って
    //  新しいシートを消してしまうため)
    var allSheets = ss.getSheets();
    var dbSheets = [];
    allSheets.forEach(function(s) {
      var sName = s.getName();
      var m = sName.match(/^抽出DB_(\d{4})_(\d{4})/);
      if (m) {
        dbSheets.push({ sheet: s, key: m[1] + m[2] });
      }
    });
    dbSheets.sort(function(a, b) { return a.key < b.key ? -1 : (a.key > b.key ? 1 : 0); }); // 古い順

    var maxKeepSheets = 3;
    if (dbSheets.length >= maxKeepSheets) {
      for (var k = 0; k <= dbSheets.length - maxKeepSheets; k++) {
        try {
          ss.deleteSheet(dbSheets[k].sheet);
        } catch(e) {}
      }
    }
    
    var formattedDate = Utilities.formatDate(new Date(), "JST", "MMdd_HHmm");
    var newSheetName = "抽出DB_" + formattedDate;
    var sheet = ss.getSheetByName(newSheetName);
    var count = 1;
    while (sheet) {
      newSheetName = "抽出DB_" + formattedDate + "_" + count;
      sheet = ss.getSheetByName(newSheetName);
      count++;
    }
    
    var newSheet = ss.insertSheet(newSheetName);
    var normalizedRows = normalizeDataArray(filteredRows);
    
    newSheet.getRange(1, 1, normalizedRows.length, normalizedRows[0].length).setValues(normalizedRows);
    newSheet.getRange(1, 1, 1, normalizedRows[0].length).setBackground("#e8f0fe").setFontWeight("bold").setHorizontalAlignment("center");
    
    var maxRows = newSheet.getMaxRows();
    var maxCols = newSheet.getMaxColumns();
    var dataRows = normalizedRows.length;
    var dataCols = normalizedRows[0].length;
    
    if (maxRows > dataRows) {
      newSheet.deleteRows(dataRows + 1, maxRows - dataRows);
    }
    if (maxCols > dataCols) {
      newSheet.deleteColumns(dataCols + 1, maxCols - dataCols);
    }
    
    return {
      message: "データのクレンジングが成功しました。<br>抽出DBシート「" + newSheetName + "」に保存されました。<br>(除外レコード数: " + excludedCount.toLocaleString() + "件)",
      url: ss.getUrl() + "#gid=" + newSheet.getSheetId()
    };
  } catch (err) { return { error: "データ整理実行エラー: " + err.toString() }; }
}

/**
 * ダッシュボード集計：新定義のユニークユーザー（ニックネーム＋性別＋生年月日＋都道府県）で集計
 * @param {boolean} [enableDedup=true] ニックネーム＋性別＋生年月日＋都道府県による同一人物の名寄せを行うか（falseの場合は行単位＝そのままの数で集計）
 * @param {boolean} [excludeEmptyAge=true] 年齢が未入力（0または空）のユーザーを対象者数等の集計から除外するか
 */
function getAdvancedAnalyticsData(sheetName, enableDedup, excludeEmptyAge) {
  try {
    enableDedup = (enableDedup !== false);
    excludeEmptyAge = (excludeEmptyAge !== false);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { error: "シートが見つかりません" };
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { error: "データがありません" };
    
    var headers = data[0];
    var cleanHeaders = headers.map(cleanText);
    
    var idxName = cleanHeaders.findIndex(function(h) { return h.indexOf('ニックネーム') !== -1; });
    var idxDate = cleanHeaders.findIndex(function(h) { return h.indexOf('日付') !== -1 || h.indexOf('日時') !== -1; });
    var idxAge = cleanHeaders.findIndex(function(h) { return h.indexOf('年齢') !== -1; });
    var idxGender = cleanHeaders.findIndex(function(h) { return h.indexOf('性別') !== -1; });
    var idxGroup = cleanHeaders.findIndex(function(h) { return h.indexOf('連携グループ') !== -1 || h.indexOf('連携済みグループ') !== -1; });
    var idxA4 = cleanHeaders.findIndex(function(h) { return h.indexOf('4文字平均') !== -1; });
    var idxA3 = cleanHeaders.findIndex(function(h) { return h.indexOf('3文字平均') !== -1; });
    
    // 【新定義のための属性カラム】生年月日・都道府県の特定
    var idxBirth = cleanHeaders.findIndex(function(h) { return h.indexOf('生年月日') !== -1; });
    var idxPref = cleanHeaders.findIndex(function(h) { return h.indexOf('都道府県') !== -1; });
    
    var idxP = findPatakaIndex(cleanHeaders, 'パ');
    var idxT = findPatakaIndex(cleanHeaders, 'タ');
    var idxK = findPatakaIndex(cleanHeaders, 'カ');
    var idxR = findPatakaIndex(cleanHeaders, 'ラ');

    var userMap = {};
    // 「集計条件」の名寄せOFF時に使う、行単位（そのままの数）の対象者数・データ件数カウンター
    var rawRowCounts = {
      total: { male: 0, female: 0, other: 0 },
      managed: { male: 0, female: 0, other: 0 },
      unmanaged: { male: 0, female: 0, other: 0 }
    };
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var name = idxName !== -1 ? row[idxName] : '';
      if (!name) continue;
      var dateVal = idxDate !== -1 ? row[idxDate] : '';
      if (!dateVal) continue;
      
      var age = idxAge !== -1 ? parseFloat(row[idxAge]) : null;
      var gender = idxGender !== -1 ? row[idxGender] : '';
      var group = idxGroup !== -1 ? row[idxGroup] : '';
      
      // 生年月日、都道府県の取得 (空白許容ガード)
      var birth = idxBirth !== -1 ? row[idxBirth] : '';
      var pref = idxPref !== -1 ? row[idxPref] : '';
      
      var pVal = idxP !== -1 ? parseFloat(row[idxP]) || 0 : 0;
      var tVal = idxT !== -1 ? parseFloat(row[idxT]) || 0 : 0;
      var kVal = idxK !== -1 ? parseFloat(row[idxK]) || 0 : 0;
      var rVal = idxR !== -1 ? parseFloat(row[idxR]) || 0 : 0;
      var a4Val = idxA4 !== -1 ? parseFloat(row[idxA4]) || 0 : 0;
      var calculatedA3 = (pVal + tVal + kVal) / 3;
      var a3Val = idxA3 !== -1 ? parseFloat(row[idxA3]) || calculatedA3 : calculatedA3;
      
      // ニックネーム、性別、生年月日、都道府県の値を前後のスペースや改行を除去した上で結合し、userIdKeyを生成
      var nameClean = cleanText(name);
      var genderClean = cleanText(gender);
      var birthClean = cleanText(birth);
      var prefClean = cleanText(pref);
      var groupClean = cleanText(group);
      
      // 散布図・離脱率等のグラフ用データ(userMap)は、名寄せ設定に関わらず常にフル名寄せキーで構築する（グラフの見た目を変えないため）
      var userIdKey = nameClean + '|' + genderClean + '|' + birthClean + '|' + prefClean;

      // =========================================================
      // 【管理・未管理の判定ロジックバグの修正】
      // グループ名が 空欄 / "-" / "未設定" / "未登録" のいずれかである場合は、
      // 完璧に「管理なし (unmanaged)」に自動分類します。
      // =========================================================
      var isUnmanaged = (groupClean === "" || groupClean === "-" || groupClean === "未設定" || groupClean === "未登録");
      var groupType = isUnmanaged ? 'unmanaged' : 'managed';
      var groupName = isUnmanaged ? '' : group; // 管理なしの場合はグループ名（系列名）として扱わない

      // 「集計条件」の名寄せOFF時に使う、行単位（そのままの数）の対象者数・データ件数カウンター
      var rowGenderKey = "other";
      if (genderClean === "男性" || genderClean === "男" || genderClean === "M") rowGenderKey = "male";
      else if (genderClean === "女性" || genderClean === "女" || genderClean === "F") rowGenderKey = "female";
      if (!(excludeEmptyAge && !(age > 0))) {
        rawRowCounts.total[rowGenderKey]++;
        rawRowCounts[groupType][rowGenderKey]++;
      }

      if (!userMap[userIdKey]) {
        userMap[userIdKey] = { 
          uid: name, // 表示用ニックネーム
          age: age, 
          gender: gender, 
          groupName: groupName, 
          groupType: groupType, 
          records: [] 
        };
      }
      userMap[userIdKey].records.push({ date: new Date(dateVal), p: pVal, t: tVal, k: kVal, r: rVal, a4: a4Val, a3: a3Val });
    }

    var userNames = Object.keys(userMap);
    var createEmptyResult = function() { return { userCount: 0, userCounts: { male: 0, female: 0, other: 0 }, recordCounts: { male: 0, female: 0, other: 0 }, scatter: { male: [], female: [], other: [] } }; };
    var result = { total: createEmptyResult(), managed: createEmptyResult(), unmanaged: createEmptyResult(), groups: {}, usersChurn: [] };

    userNames.forEach(function(userIdKey) {
      var user = userMap[userIdKey];
      var genderKey = "other";
      if (user.gender === "男性" || user.gender === "男" || user.gender === "M") genderKey = "male";
      else if (user.gender === "女性" || user.gender === "女" || user.gender === "F") genderKey = "female";

      // 「対象者数」「データ件数」（=集計条件表示専用の数値）だけを、年齢未入力除外設定に応じて加算する。
      // 散布図・離脱率等のグラフ用データ生成（この下の処理）は設定に関わらず常に実行し、グラフの見た目は変えない。
      if (!(excludeEmptyAge && !(user.age > 0))) {
        result.total.userCount++; result.total.userCounts[genderKey]++; result.total.recordCounts[genderKey] += user.records.length;
        result[user.groupType].userCount++; result[user.groupType].userCounts[genderKey]++; result[user.groupType].recordCounts[genderKey] += user.records.length;
        if (user.groupType === 'managed' && user.groupName) {
          if (!result.groups[user.groupName]) result.groups[user.groupName] = createEmptyResult();
          result.groups[user.groupName].userCount++; result.groups[user.groupName].userCounts[genderKey]++; result.groups[user.groupName].recordCounts[genderKey] += user.records.length;
        }
      }

      user.records.sort(function(a, b) { return a.date - b.date; });
      user.records.forEach(function(rec) {
        if (user.age > 0) {
          var pt = { 
            x: user.age, 
            y: rec.a4, 
            uid: user.uid, // 複合キー（userIdKey）ではなくニックネームで描画
            a4: rec.a4, 
            a3: rec.a3, 
            p: rec.p, 
            t: rec.t, 
            k: rec.k, 
            r: rec.r 
          };
          result.total.scatter[genderKey].push(pt);
          result[user.groupType].scatter[genderKey].push(pt);
          if (user.groupType === 'managed' && user.groupName) result.groups[user.groupName].scatter[genderKey].push(pt);
        }
      });

      if (user.records.length > 0) {
        var firstDate = user.records[0].date;
        var firstMidnight = new Date(firstDate.getFullYear(), firstDate.getMonth(), firstDate.getDate()).getTime();
        var lastRecord = user.records[user.records.length - 1];
        var lastMidnight = new Date(lastRecord.date.getFullYear(), lastRecord.date.getMonth(), lastRecord.date.getDate()).getTime();
        var accountAgeDays = Math.round((lastMidnight - firstMidnight) / (1000 * 60 * 60 * 24));
        var weeklyDaySets = {}; var weeklyScoreMap = {};
        user.records.forEach(function(rec) {
          var recMidnight = new Date(rec.date.getFullYear(), rec.date.getMonth(), rec.date.getDate()).getTime();
          var relDay = Math.round((recMidnight - firstMidnight) / (1000 * 60 * 60 * 24));
          var w = Math.floor(relDay / 7);
          if (w < 24) {
            if (!weeklyDaySets[w]) weeklyDaySets[w] = new Set();
            weeklyDaySets[w].add(relDay);
            if (!weeklyScoreMap[w]) weeklyScoreMap[w] = [];
            weeklyScoreMap[w].push(rec.a4);
          }
        });
        var dataArr = []; var scoresArr = [];
        for (var w = 0; w < 24; w++) {
          if (accountAgeDays >= w * 7) {
            dataArr.push(weeklyDaySets[w] ? weeklyDaySets[w].size : 0);
            var wScores = weeklyScoreMap[w];
            if (wScores && wScores.length > 0) {
              var sum = wScores.reduce(function(a, b){ return a + b; }, 0);
              scoresArr.push(sum / wScores.length);
            } else scoresArr.push(null);
          } else { dataArr.push(null); scoresArr.push(null); }
        }
        result.usersChurn.push({ uid: user.uid, groupType: user.groupType, groupName: user.groupName, age: user.age, gender: user.gender, data: dataArr, scores: scoresArr, accountAgeDays: accountAgeDays, activeDaysSet: Array.from(new Set(user.records.map(function(r) { return Math.round((new Date(r.date.getFullYear(), r.date.getMonth(), r.date.getDate()).getTime() - firstMidnight) / (1000 * 60 * 60 * 24)); }))) });
      }
    });

    var generateChurnStats = function(targetObj, filterFn) {
      targetObj.churn = { weekly: [], dailyFirstWeek: [] };
      for (var w = 0; w < 24; w++) {
        var sumDays = 0, activeUsers = 0, values = [], html = {0:0, 1:0, 2:0, 3:0, 4:0, 5:0, 6:0, 7:0};
        result.usersChurn.forEach(function(u) {
          if (!filterFn(u)) return;
          if (u.accountAgeDays >= w * 7) { activeUsers++; var days = u.data[w] || 0; sumDays += days; values.push(days); html[days]++; }
        });
        var avg = activeUsers > 0 ? parseFloat((sumDays / activeUsers).toFixed(2)) : null;
        var min = activeUsers > 0 ? Math.min.apply(null, values) : null;
        var max = activeUsers > 0 ? Math.max.apply(null, values) : null;
        var sd = null;
        if (activeUsers > 0) {
          var variance = values.reduce(function(a, b) { return a + Math.pow(b - avg, 2); }, 0) / activeUsers;
          sd = parseFloat(Math.sqrt(variance).toFixed(2));
        }
        targetObj.churn.weekly.push({ x: w + 1, y: avg, users: activeUsers, min: min, max: max, sd: sd, hist: html });
      }
      for (var d = 0; d < 168; d++) {
        var activeUsers = 0, doneUsers = 0, html = { "0": 0, "1": 0 };
        result.usersChurn.forEach(function(u) {
          if (!filterFn(u)) return;
          if (u.accountAgeDays >= d) { activeUsers++; if (u.activeDaysSet.indexOf(d) !== -1) { doneUsers++; html["1"]++; } else html["0"]++; }
        });
        targetObj.churn.dailyFirstWeek.push({ x: d + 1, y: activeUsers > 0 ? parseFloat((doneUsers / activeUsers * 100).toFixed(1)) : null, users: activeUsers, done: doneUsers, html: html });
      }
    };
    generateChurnStats(result.total, function(u) { return true; });
    generateChurnStats(result.managed, function(u) { return u.groupType === 'managed'; });
    generateChurnStats(result.unmanaged, function(u) { return u.groupType === 'unmanaged'; });
    
    Object.keys(result.groups).forEach(function(gName) { generateChurnStats(result.groups[gName], function(u) { return u.groupType === 'managed' && u.groupName === gName; }); });
    result.groupRanking = Object.keys(result.groups).map(function(gName) { return { name: gName, userCount: result.groups[gName].userCount }; });
    result.groupRanking.sort(function(a, b) { return b.userCount - a.userCount; });

    // 名寄せOFF時は「対象者数」「データ件数」を行単位（そのままの数）に置き換える（グラフ用データには影響させない）
    if (!enableDedup) {
      ['total', 'managed', 'unmanaged'].forEach(function(key) {
        var counts = rawRowCounts[key];
        result[key].userCounts = { male: counts.male, female: counts.female, other: counts.other };
        result[key].userCount = counts.male + counts.female + counts.other;
        result[key].recordCounts = { male: counts.male, female: counts.female, other: counts.other };
      });
    }

    return JSON.stringify(result);
  } catch (err) { return { error: "GAS実行エラー: " + err.toString() }; }
}

/**
 * デバッグ用: 指定したユーザーの週次データ・24週ブロック判定を直接確認する
 */
function debugUserChurn(sheetName, uidQuery) {
  try {
    var raw = getAdvancedAnalyticsData(sheetName);
    var result = typeof raw === 'string' ? JSON.parse(raw) : raw;
    if (result.error) return result;
    var matches = result.usersChurn.filter(function(u) { return String(u.uid).indexOf(uidQuery) !== -1; });
    return matches.map(function(u) {
      var blocks = [];
      for (var b = 0; b < 12; b++) {
        var w1 = b * 2, w2 = b * 2 + 1;
        var d1 = u.data[w1] || 0, d2 = u.data[w2] || 0;
        blocks.push(d1 + d2);
      }
      return {
        uid: u.uid, groupType: u.groupType, groupName: u.groupName,
        accountAgeDays: u.accountAgeDays,
        data: u.data, scores: u.scores, blocks: blocks
      };
    });
  } catch (err) { return { error: "debugUserChurn エラー: " + err.toString() }; }
}

/**
 * デバッグ用: uid(ニックネーム)が文字列型でないユーザーが何件あるか確認する
 */
function debugNonStringUids(sheetName) {
  try {
    var raw = getAdvancedAnalyticsData(sheetName);
    var result = typeof raw === 'string' ? JSON.parse(raw) : raw;
    if (result.error) return result;
    var total = result.usersChurn.length;
    var offenders = result.usersChurn
      .map(function(u, i) { return { index: i, uid: u.uid, type: typeof u.uid }; })
      .filter(function(x) { return x.type !== 'string'; });
    return { totalUsers: total, nonStringCount: offenders.length, offenders: offenders };
  } catch (err) { return { error: "debugNonStringUids エラー: " + err.toString() }; }
}