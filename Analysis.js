// 分析モジュール
const AnalysisModule = {
    /**
     * すべての分析を実行する
     */
    analyzeAllData: function() {
      try {
        LogModule.logOperation("全分析処理", "開始");
        
        // コア分析の実行
        this.analyzeTagDifficulty();
        this.analyzeTagDistribution();
        this.analyzeIndividualSkills();
        this.analyzeTeamPerformance();
        this.analyzeAreaCoverage();
        // 新規追加：相談フロー分析
        this.analyzeConsultationFlow();
        
        // ダッシュボードの初期化
        DashboardModule.initTeamPerformanceDashboard();
        DashboardModule.initAreaCoverageDashboard();
        DashboardModule.initPersonalDashboard();
        // 新規追加：相談フローダッシュボード
        DashboardModule.initConsultationFlowDashboard();
        
        LogModule.logOperation("全分析処理", "完了");
        UtilsModule.showToast('全ての分析が完了しました！', 'スキルマップ分析');
        
        return true;
      } catch (e) {
        LogModule.handleError(e, "全分析処理");
        return false;
      }
    },
  
    /**
     * 機能タグ難易度分析を実行する
     */
    analyzeTagDifficulty: function() {
      LogModule.logOperation("タグ難易度分析", "開始");
      
      try {
        // データ準備
        const data = DataModule.prepareAnalysisData();
        if (!data) {
          throw new Error('分析データの準備に失敗しました');
        }
        
        // 設定値取得
        const settings = SettingsModule.getSettings();
        
        // 必要な列のインデックスを取得
        const colIndexes = data.txColIndexes;
        
        // 必須列のチェック
        const requiredColumns = ['tagName', 'opName', 'responseTime', 'rallyCount', 'userName'];
        for (const colName of requiredColumns) {
          if (colIndexes[colName] === -1) {
            throw new Error(`必要な列 ${colName} が見つかりません`);
          }
        }
        
        // タグごとにデータを集計
        const tagStats = this.collectTagStats(data);
        
        // 分析結果の作成
        const results = this.calculateTagDifficulty(tagStats, settings, data);
        
        // 難易度調整マトリックスの準備
        this.prepareAdjustmentMatrix(results);
        
        // 結果を書き込み
        this.writeTagDifficultyResults(results);
        
        LogModule.logOperation("タグ難易度分析", `完了: ${results.length}件のタグを分析`);
        UtilsModule.showToast(`${results.length}件のタグ難易度分析が完了しました！`, 'タグ難易度分析');
        
        return results;
      } catch (e) {
        LogModule.handleError(e, 'タグ難易度分析');
        return null;
      }
    },
    
    /**
     * タグごとの統計データを収集する
     */
    collectTagStats: function(data) {
      const tagStats = {};
      const allResponders = new Set();
      const txData = data.transactionData;
      const colIndexes = data.txColIndexes;
      
      for (let i = 1; i < txData.length; i++) {
        const row = txData[i];
        const tagName = row[colIndexes.tagName];
        const opName = row[colIndexes.opName];
        const responseTime = row[colIndexes.responseTime];
        // Rally数を考慮
        const rallyCount = colIndexes.rallyCount !== -1 ? (row[colIndexes.rallyCount] || 1) : 1;
        const userName = row[colIndexes.userName];
        
        // OJTタグの除外（オプション）
        const isOjtCase = colIndexes.hasOjtTag !== -1 && row[colIndexes.hasOjtTag] === "TRUE";
        if (isOjtCase && settings && settings.excludeOjtData === true) {
          continue;
        }
        
        // マクロ使用フラグ
        const useMacro = colIndexes.useMacro !== -1 ? (row[colIndexes.useMacro] === "マクロ使用") : false;
        
        // 相談フラグ（関連カラムに値があるかどうか）
        const hasConsultation = 
          (colIndexes.postTimestamp !== -1 && row[colIndexes.postTimestamp]) || 
          (colIndexes.soudanUserName !== -1 && row[colIndexes.soudanUserName]) || 
          (colIndexes.adviserUserName !== -1 && row[colIndexes.adviserUserName]);
        
        // 機能タグのみを処理
        if (tagName && UtilsModule.isFunctionTag(tagName)) {
          if (userName) allResponders.add(userName);
          
          if (!tagStats[tagName]) {
            tagStats[tagName] = {
              tagName: tagName,
              opName: opName,
              responseTimes: [],
              rallyCounts: [],
              responders: new Set(),
              totalCount: 0,
              totalRallyCount: 0,
              macroUsageCount: 0,
              consultationCount: 0
            };
          }
          
          // 異常値を除外（24時間以上など）
          if (responseTime > 0 && responseTime < 1440) {
            tagStats[tagName].responseTimes.push(responseTime);
            tagStats[tagName].rallyCounts.push(rallyCount);
            tagStats[tagName].responders.add(userName);
            tagStats[tagName].totalCount++;
            tagStats[tagName].totalRallyCount += rallyCount;
            
            // マクロ使用のカウント
            if (useMacro) {
              tagStats[tagName].macroUsageCount++;
            }
            
            // 相談発生のカウント
            if (hasConsultation) {
              tagStats[tagName].consultationCount++;
            }
          }
        }
      }
  
      tagStats.allResponders = allResponders;
      tagStats.totalRecords = txData.length - 1;
      tagStats.totalRallyRecords = 0;
      
      // 総Rally数の計算
      Object.keys(tagStats).forEach(key => {
        if (key !== 'allResponders' && key !== 'totalRecords' && key !== 'totalRallyRecords') {
          tagStats.totalRallyRecords += tagStats[key].totalRallyCount;
        }
      });
      
      return tagStats;
    },
    
    /**
     * タグの難易度を計算する
     */
    calculateTagDifficulty: function(tagStats, settings, data) {
      const results = [];
      const totalResponders = tagStats.allResponders.size;
      const totalTags = Object.keys(tagStats).length - 3; // allResponders, totalRecords, totalRallyRecordsを除く
      const totalRecords = tagStats.totalRecords;
      const totalRallyRecords = tagStats.totalRallyRecords || totalRecords; // ゼロの場合はレコード数で代用
      
      Object.entries(tagStats).forEach(([tagName, tag]) => {
        // allResponders, totalRecords, totalRallyRecordsはスキップ
        if (tagName === 'allResponders' || tagName === 'totalRecords' || tagName === 'totalRallyRecords') return;
        
        const responseTimes = tag.responseTimes.sort((a, b) => a - b);
        const rallyCounts = tag.rallyCounts.sort((a, b) => a - b);
        const count = responseTimes.length;
        
        if (count > 0) {
          // 統計値計算
          const medianTime = UtilsModule.calculatePercentile(responseTimes, 50);
          const p25 = UtilsModule.calculatePercentile(responseTimes, 25);
          const p75 = UtilsModule.calculatePercentile(responseTimes, 75);
          const iqr = p75 - p25;
          
          // Rally平均計算
          const avgRallyCount = tag.rallyCounts.reduce((sum, count) => sum + count, 0) / tag.rallyCounts.length;
          
          // スキルカバー率
          const responderCoverage = tag.responders.size / totalResponders * 100;
          
          // マクロ利用率計算
          const macroUsageRate = tag.totalCount > 0 ? (tag.macroUsageCount / tag.totalCount) * 100 : 0;
          
          // 相談発生率計算
          const consultationRate = tag.totalCount > 0 ? (tag.consultationCount / tag.totalCount) * 100 : 0;
          
          // 多次元評価指標の計算
          
          // 1. 技術的複雑性（Technical Complexity）
          const timeScore = UtilsModule.calculateTimeScore(medianTime);
          const coverageScore = UtilsModule.calculateCoverageScore(responderCoverage);
          const macroAdjustment = Math.max(0, 5 - (macroUsageRate / 20)); // マクロ利用率が高いほど下がる（0-5）
          
          const technicalComplexity = 
            timeScore * 0.5 + 
            coverageScore * 0.3 + 
            macroAdjustment * 0.2;
          
          // 2. 解決難易度（Resolution Difficulty）
          const rallyScore = UtilsModule.calculateRallyScore(avgRallyCount);
          const consultationScore = Math.min(5, consultationRate / 10); // 相談率が高いほど高い（0-5）
          const variabilityScore = Math.min(5, iqr / 10); // ばらつきが大きいほど高い（0-5）
          
          const resolutionDifficulty = 
            rallyScore * 0.5 + 
            consultationScore * 0.3 + 
            variabilityScore * 0.2;
          
          // 3. 知識要求度（Knowledge Requirement）
          const knowledgeRequirement = 
            coverageScore * 0.7 + 
            variabilityScore * 0.3;
          
          // 総合難易度スコア
          const difficultyScore = 
            technicalComplexity * 0.4 + 
            resolutionDifficulty * 0.4 + 
            knowledgeRequirement * 0.2;
          
          // 各次元のレベル判定
          const tcLevel = UtilsModule.getDimensionLevel(technicalComplexity, "tc", settings);
          const rdLevel = UtilsModule.getDimensionLevel(resolutionDifficulty, "rd", settings);
          const krLevel = UtilsModule.getDimensionLevel(knowledgeRequirement, "kr", settings);
          
          // 総合難易度レベル判定
          const difficultyLevel = UtilsModule.getDimensionLevel(difficultyScore, "tc", settings);
          
          // タグタイプ判定（Rally出現頻度によるコア/標準/レア分類）
          const rallyFrequencyPercent = tag.totalRallyCount / totalRallyRecords * 100;
          const frequencyPercent = tag.totalCount / totalRecords * 100;
          
          // 両方の頻度の高い方を採用
          const effectiveFrequency = Math.max(rallyFrequencyPercent, frequencyPercent);
          
          let tagType;
          if (effectiveFrequency >= settings.coreSkillThreshold) {
            tagType = "コア";
          } else if (effectiveFrequency <= settings.rareCaseThreshold) {
            tagType = "レア";
          } else {
            tagType = "標準";
          }
          
          // 初級者向けフラグ
          const beginnerFriendly = macroUsageRate >= 50 ? "○" : "";
          
          // 高相談タグフラグ
          const highConsultation = consultationRate >= 30 ? "○" : "";
          
          results.push([
            tag.tagName,                    // タグ名
            tag.opName,                     // OPチーム名
            count,                          // 対応件数
            tag.totalRallyCount,           // 総Rally数
            avgRallyCount.toFixed(1),       // Rally平均
            tag.responders.size,            // 対応者数
            responderCoverage.toFixed(1),   // 対応者カバー率(%)
            medianTime.toFixed(1),          // 対応時間中央値
            p25.toFixed(1),                 // 25パーセンタイル
            p75.toFixed(1),                 // 75パーセンタイル
            iqr.toFixed(1),                 // IQR
            macroUsageRate.toFixed(1),      // マクロ利用率(%)
            consultationRate.toFixed(1),    // 相談発生率(%)
            technicalComplexity.toFixed(1), // 技術的複雑性スコア(1-5)
            tcLevel,                        // 技術的複雑性レベル
            resolutionDifficulty.toFixed(1), // 解決難易度スコア(1-5)
            rdLevel,                        // 解決難易度レベル
            knowledgeRequirement.toFixed(1), // 知識要求度スコア(1-5)
            krLevel,                        // 知識要求度レベル
            difficultyScore.toFixed(1),     // 総合難易度スコア(1-5)
            difficultyLevel,                // 難易度レベル
            tagType,                        // タグタイプ
            beginnerFriendly,               // 初級者向け
            highConsultation,               // 高相談タグ
            new Date()                      // 更新日時
          ]);
        }
      });
      
      return results;
    },
    
    /**
     * 難易度調整マトリックスを準備する
     */
    prepareAdjustmentMatrix: function(difficultyResults) {
      try {
        const adjustmentSheet = UtilsModule.getOrCreateSheet(SHEETS.ADJUSTMENT);
        
        // 既存の調整値を保持
        const existingData = {};
        if (adjustmentSheet.getLastRow() > 1) {
          const adjustmentData = adjustmentSheet.getDataRange().getValues();
          const headers = adjustmentData[0];
          
          const tagIdx = headers.indexOf('タグ名');
          const adjustIdx = headers.indexOf('定性調整値');
          
          if (tagIdx !== -1 && adjustIdx !== -1) {
            for (let i = 1; i < adjustmentData.length; i++) {
              const tag = adjustmentData[i][tagIdx];
              const adjustment = adjustmentData[i][adjustIdx];
              if (tag) {
                existingData[tag] = adjustment;
              }
            }
          }
        }
        
        // シートをクリア
        adjustmentSheet.clear();
        
        // ヘッダー設定
        const outputheaders = [
          "タグ名", "OPチーム名", "自動計算難易度", "定性調整値(-1～+1)", "最終難易度"
        ];
        
        adjustmentSheet.getRange(1, 1, 1, outputheaders.length).setValues([outputheaders])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // データの準備
        const adjustmentData = difficultyResults.map(row => {
          const tagName = row[0];
          const opTeam = row[1];
          const autoLevel = row[20];
          
          // 既存の調整値があれば使用、なければデフォルト0
          const adjustment = existingData[tagName] !== undefined ? existingData[tagName] : 0;
          
          // 調整後の難易度レベル計算
          const autoScore = parseFloat(row[19]);
          const settings = SettingsModule.getSettings();
          const adjustedScore = Math.max(1, Math.min(5, autoScore + adjustment));
          
          let finalLevel = UtilsModule.getDimensionLevel(adjustedScore, "tc", settings);
          
          return [
            tagName,
            opTeam,
            autoLevel,
            adjustment,
            finalLevel
          ];
        });
        
        // データの書き込み
        if (adjustmentData.length > 0) {
          adjustmentSheet.getRange(2, 1, adjustmentData.length, outputheaders.length).setValues(adjustmentData);
        }
        
        // 条件付き書式の設定
        this.setAdjustmentSheetFormatting(adjustmentSheet);
        
        // 列幅の調整
        adjustmentSheet.setColumnWidth(1, 200); // タグ名
        adjustmentSheet.setColumnWidth(2, 150); // OPチーム名
        
        return true;
        
      } catch (e) {
        LogModule.handleError(e, "難易度調整マトリックス準備");
        return false;
      }
    },
    
    /**
     * 難易度調整マトリックスの条件付き書式を設定する
     */
    setAdjustmentSheetFormatting: function(sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return;
      
      // 調整値列の書式設定
      const adjustmentRange = sheet.getRange(2, 4, lastRow - 1, 1);
      
      // -1,0,+1で色分け
      const negativeRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(0)
        .setBackground('#FFCDD2') // 薄い赤
        .setRanges([adjustmentRange])
        .build();
      
      const positiveRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setBackground('#C8E6C9') // 薄い緑
        .setRanges([adjustmentRange])
        .build();
      
      let rules = sheet.getConditionalFormatRules();
      rules.push(negativeRule, positiveRule);
      sheet.setConditionalFormatRules(rules);
      
      // 最終難易度列の書式設定
      const finalLevelRange = sheet.getRange(2, 5, lastRow - 1, 1);
      
      // 難易度レベル別の色分け
      const blueColors = ['#E3F2FD', '#90CAF9', '#42A5F5', '#1976D2', '#0D47A1'];
      
      for (let i = 1; i <= 5; i++) {
        const level = 'L' + i;
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(level)
          .setBackground(blueColors[i - 1])
          .setFontColor(i >= 4 ? 'white' : 'black') // 濃い色には白文字
          .setRanges([finalLevelRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(rule);
        sheet.setConditionalFormatRules(rules);
      }
    },
    
    /**
     * タグ難易度分析結果をシートに書き込む
     */
    writeTagDifficultyResults: function(results) {
      const difficultySheet = UtilsModule.getOrCreateSheet(SHEETS.TAG_DIFFICULTY);
      
      // シートをクリア
      difficultySheet.clear();
      
      // ヘッダー設定
      const outputheaders = [
        "タグ名", "OPチーム名", "対応件数", "総Rally数", "Rally平均", "対応者数", "対応者カバー率(%)",
        "対応時間中央値", "25パーセンタイル", "75パーセンタイル", "IQR",
        "マクロ利用率(%)", "相談発生率(%)",
        "技術的複雑性(1-5)", "技術的複雑性レベル",
        "解決難易度(1-5)", "解決難易度レベル",
        "知識要求度(1-5)", "知識要求度レベル",
        "総合難易度スコア(1-5)", "難易度レベル", "タグタイプ", 
        "初級者向け", "高相談タグ", "更新日時"
      ];
      
      difficultySheet.getRange(1, 1, 1, outputheaders.length).setValues([outputheaders])
        .setFontWeight('bold').setBackground('#f3f3f3');
      
      if (results.length > 0) {
        difficultySheet.getRange(2, 1, results.length, outputheaders.length).setValues(results);
        // 条件付き書式の設定
        this.setDifficultySheetFormatting(difficultySheet);
      }
    },
    
    /**
     * タグ難易度シートの条件付き書式を設定する
     */
    setDifficultySheetFormatting: function(sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return;
      
      // 各スコア列のインデックス
      const scoreColumns = [
        {column: 14, dimension: "技術的複雑性"}, 
        {column: 16, dimension: "解決難易度"}, 
        {column: 18, dimension: "知識要求度"}, 
        {column: 20, dimension: "総合難易度"}
      ];
      
      // スコア列の書式設定
      scoreColumns.forEach(col => {
        const range = sheet.getRange(2, col.column, lastRow - 1, 1);
        
        // 1-5の色分け
        const blueColors = ['#E3F2FD', '#90CAF9', '#42A5F5', '#1976D2', '#0D47A1'];
        
        for (let i = 1; i <= 5; i++) {
          const lowerBound = i - 0.5;
          const upperBound = i + 0.49;
          
          const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberBetween(lowerBound, upperBound)
            .setBackground(blueColors[i - 1])
            .setFontColor(i >= 4 ? 'white' : 'black')
            .setRanges([range])
            .build();
          
          const rules = sheet.getConditionalFormatRules();
          rules.push(rule);
          sheet.setConditionalFormatRules(rules);
        }
      });
      
      // 各次元のレベル列
      const levelColumns = [
        {column: 15, dimension: "技術的複雑性レベル"},
        {column: 17, dimension: "解決難易度レベル"},
        {column: 19, dimension: "知識要求度レベル"},
        {column: 21, dimension: "難易度レベル"}
      ];
      
      // レベル列の書式設定
      levelColumns.forEach(col => {
        const range = sheet.getRange(2, col.column, lastRow - 1, 1);
        
        // L1-L5の色分け
        const blueColors = ['#E3F2FD', '#90CAF9', '#42A5F5', '#1976D2', '#0D47A1'];
        
        for (let i = 1; i <= 5; i++) {
          const level = 'L' + i;
          const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(level)
            .setBackground(blueColors[i - 1])
            .setFontColor(i >= 4 ? 'white' : 'black')
            .setRanges([range])
            .build();
          
          const rules = sheet.getConditionalFormatRules();
          rules.push(rule);
          sheet.setConditionalFormatRules(rules);
        }
      });
      
      // タグタイプ列の書式設定
      const typeRange = sheet.getRange(2, 22, lastRow - 1, 1);
      
      // コア、標準、レアで色分け
      const coreRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("コア")
        .setBackground('#FFF9C4') // 黄色
        .setRanges([typeRange])
        .build();
      
      const rareRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("レア")
        .setBackground('#FFEBEE') // 薄い赤
        .setRanges([typeRange])
        .build();
      
      let rules = sheet.getConditionalFormatRules();
      rules.push(coreRule, rareRule);
      sheet.setConditionalFormatRules(rules);
      
      // 初級者向け列と高相談タグ列の書式設定
      const beginnerRange = sheet.getRange(2, 23, lastRow - 1, 1);
      const consultationRange = sheet.getRange(2, 24, lastRow - 1, 1);
      
      // ○マークがある場合に色付け
      const beginnerRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("○")
        .setBackground('#E8F5E9') // 薄い緑
        .setRanges([beginnerRange])
        .build();
      
      const consultationRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo("○")
        .setBackground('#FFF3E0') // 薄いオレンジ
        .setRanges([consultationRange])
        .build();
      
      rules = sheet.getConditionalFormatRules();
      rules.push(beginnerRule, consultationRule);
      sheet.setConditionalFormatRules(rules);
      
      // マクロ利用率と相談発生率の書式設定
      const macroRange = sheet.getRange(2, 12, lastRow - 1, 1);
      
      const highMacroRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(50)
        .setBackground('#E8F5E9') // 薄い緑
        .setRanges([macroRange])
        .build();
      
      rules = sheet.getConditionalFormatRules();
      rules.push(highMacroRule);
      sheet.setConditionalFormatRules(rules);
      
      const consultRateRange = sheet.getRange(2, 13, lastRow - 1, 1);
      
      const highConsultRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(30)
        .setBackground('#FFF3E0') // 薄いオレンジ
        .setRanges([consultRateRange])
        .build();
      
      rules = sheet.getConditionalFormatRules();
      rules.push(highConsultRule);
      sheet.setConditionalFormatRules(rules);
      
      // 列幅の調整
      sheet.setColumnWidth(1, 200); // タグ名
      sheet.setColumnWidth(2, 120); // OPチーム名
      sheet.autoResizeColumns(3, 22); // 他の列を自動調整
    },
  
    /**
     * タグ内難易度分布分析を実行する
     */
    analyzeTagDistribution: function() {
      LogModule.logOperation("タグ内難易度分布分析", "開始");
      
      try {
        // データ準備
        const data = DataModule.prepareAnalysisData();
        if (!data) {
          throw new Error('分析データの準備に失敗しました');
        }
        
        // 設定値取得
        const settings = SettingsModule.getSettings();
        
        // シート取得
        const difficultySheet = data.sheets.difficultySheet;
        const distributionSheet = UtilsModule.getOrCreateSheet(SHEETS.TAG_DISTRIBUTION);
        
        // 必要な列インデックスの取得
        const colIndexes = data.txColIndexes;
        const diffColIndexes = data.diffColIndexes;
        
        // タグごとにデータを集計
        const tagStats = {};
        const txData = data.transactionData;
        
        for (let i = 1; i < txData.length; i++) {
          const row = txData[i];
          const tagName = row[colIndexes.tagName];
          const userName = row[colIndexes.userName];
          const responseTime = row[colIndexes.responseTime];
          const rallyCount = colIndexes.rallyCount !== -1 ? (row[colIndexes.rallyCount] || 1) : 1;
          
          // OJTタグの除外（オプション）
          const isOjtCase = colIndexes.hasOjtTag !== -1 && row[colIndexes.hasOjtTag] === "TRUE";
          if (isOjtCase && settings && settings.excludeOjtData === true) {
            continue;
          }
          
          // マクロ使用フラグ
          const useMacro = colIndexes.useMacro !== -1 ? (row[colIndexes.useMacro] === "マクロ使用") : false;
          
          // 相談フラグ（関連カラムに値があるかどうか）
          const hasConsultation = 
            (colIndexes.postTimestamp !== -1 && row[colIndexes.postTimestamp]) || 
            (colIndexes.soudanUserName !== -1 && row[colIndexes.soudanUserName]) || 
            (colIndexes.adviserUserName !== -1 && row[colIndexes.adviserUserName]);
          
          // 機能タグのみを処理
          if (tagName && UtilsModule.isFunctionTag(tagName)) {
            if (!tagStats[tagName]) {
              tagStats[tagName] = {
                responseTimes: [],
                rallyCounts: [],
                complexityScores: [],  // 複雑性スコア（時間、Rally、マクロ、相談から算出）
                useMacroFlags: [],     // マクロ使用フラグの配列
                consultationFlags: [],  // 相談フラグの配列
                totalRallyCount: 0
              };
            }
            
            // 異常値を除外
            if (responseTime > 0 && responseTime < 1440) {
              tagStats[tagName].responseTimes.push(responseTime);
              tagStats[tagName].rallyCounts.push(rallyCount);
              tagStats[tagName].totalRallyCount += rallyCount;
              
              // マクロ使用フラグを記録
              tagStats[tagName].useMacroFlags.push(useMacro);
              
              // 相談フラグを記録
              tagStats[tagName].consultationFlags.push(hasConsultation);
              
              // 複雑性スコアの計算
              // 対応時間（40%）、Rally数（30%）、マクロ非使用（20%）、相談あり（10%）
              const timeScore = UtilsModule.calculateTimeScore(responseTime);
              const rallyScore = UtilsModule.calculateRallyScore(rallyCount);
              const macroScore = useMacro ? 1 : 3; // マクロ使用なら簡易(1)、使用なしなら複雑(3)
              const consultScore = hasConsultation ? 5 : 1; // 相談ありなら複雑(5)、なしなら簡易(1)
              
              const complexityScore = 
                timeScore * 0.4 + 
                rallyScore * 0.3 + 
                macroScore * 0.2 + 
                consultScore * 0.1;
              
              tagStats[tagName].complexityScores.push(complexityScore);
            }
          }
        }
        
        // 難易度データのマッピング
        const tagDifficultyMap = data.tagDifficultyMap;
        
        // 分析結果の作成
        const results = [];
        
        Object.entries(tagStats).forEach(([tagName, stats]) => {
          const count = stats.responseTimes.length;
          
          if (count >= 5) { // 最低5件以上あるタグのみ分析
            const responseTimes = stats.responseTimes.sort((a, b) => a - b);
            const rallyCounts = stats.rallyCounts.sort((a, b) => a - b);
            const complexityScores = stats.complexityScores.sort((a, b) => a - b);
            
            // パーセンタイル計算
            const p25Time = UtilsModule.calculatePercentile(responseTimes, 25);
            const p50Time = UtilsModule.calculatePercentile(responseTimes, 50);
            const p75Time = UtilsModule.calculatePercentile(responseTimes, 75);
            
            const p25Rally = UtilsModule.calculatePercentile(rallyCounts, 25);
            const p50Rally = UtilsModule.calculatePercentile(rallyCounts, 50);
            const p75Rally = UtilsModule.calculatePercentile(rallyCounts, 75);
            
            const p25Complexity = UtilsModule.calculatePercentile(complexityScores, 25);
            const p50Complexity = UtilsModule.calculatePercentile(complexityScores, 50);
            const p75Complexity = UtilsModule.calculatePercentile(complexityScores, 75);
            
            // マクロ使用率と相談率の計算
            const macroUsageRate = 
              (stats.useMacroFlags.filter(flag => flag).length / stats.useMacroFlags.length) * 100;
            
            const consultationRate = 
              (stats.consultationFlags.filter(flag => flag).length / stats.consultationFlags.length) * 100;
            
            // 基本難易度レベル（難易度分析から）
            const tagLevels = tagDifficultyMap[tagName] || { level: "L3", tcLevel: "L3", rdLevel: "L3", krLevel: "L3" };
            const baseLevel = tagLevels.level;
            
            // ケースタイプごとのレベル判定
            // ケースタイプはベースレベルから相対的に設定
            let simpleLevel, standardLevel, complexLevel;
            
            // ベースレベルから相対的に難易度を設定
            switch (baseLevel) {
              case "L1":
                simpleLevel = "L1";
                standardLevel = "L1";
                complexLevel = "L2";
                break;
              case "L2":
                simpleLevel = "L1";
                standardLevel = "L2";
                complexLevel = "L3";
                break;
              case "L3":
                simpleLevel = "L2";
                standardLevel = "L3";
                complexLevel = "L4";
                break;
              case "L4":
                simpleLevel = "L3";
                standardLevel = "L4";
                complexLevel = "L5";
                break;
              case "L5":
                simpleLevel = "L4";
                standardLevel = "L5";
                complexLevel = "L5";
                break;
              default:
                simpleLevel = "L2";
                standardLevel = "L3";
                complexLevel = "L4";
            }
            
            // 複雑性スコアに基づく分類の閾値
            const simpleThreshold = p25Complexity;
            const complexThreshold = p75Complexity;
            
            // タグ内案件数の分布
            const simpleCasesCount = stats.complexityScores.filter(score => score <= simpleThreshold).length;
            const complexCasesCount = stats.complexityScores.filter(score => score >= complexThreshold).length;
            const standardCasesCount = count - simpleCasesCount - complexCasesCount;
            
            const simpleCasesRate = (simpleCasesCount / count) * 100;
            const standardCasesRate = (standardCasesCount / count) * 100;
            const complexCasesRate = (complexCasesCount / count) * 100;
            
            // 月間件数と月間Rally数の推定（年間件数 ÷ 12で概算）
            const monthlyCount = Math.round(count / 3); // 3ヶ月分のデータという前提
            const monthlyRallyCount = Math.round(stats.totalRallyCount / 3);
            
            results.push([
              tagName,                               // タグ名
              count,                                 // 件数
              stats.totalRallyCount,                 // 総Rally数
              baseLevel,                             // 基本難易度
              monthlyCount,                          // 月間件数（概算）
              monthlyRallyCount,                     // 月間Rally数（概算）
              p25Time.toFixed(1),                    // 時間P25
              p50Time.toFixed(1),                    // 時間P50
              p75Time.toFixed(1),                    // 時間P75
              p25Rally.toFixed(1),                   // RallyP25
              p50Rally.toFixed(1),                   // RallyP50
              p75Rally.toFixed(1),                   // RallyP75
              macroUsageRate.toFixed(1) + "%",       // マクロ使用率
              consultationRate.toFixed(1) + "%",     // 相談率
              simpleCasesCount,                      // 簡易ケース件数
              standardCasesCount,                    // 標準ケース件数
              complexCasesCount,                     // 複雑ケース件数
              simpleCasesRate.toFixed(1) + "%",      // 簡易ケース割合
              standardCasesRate.toFixed(1) + "%",    // 標準ケース割合
              complexCasesRate.toFixed(1) + "%",     // 複雑ケース割合
              simpleCasesCount > 0 ? `${p25Time.toFixed(1)}分以下、複雑性スコア${simpleThreshold.toFixed(1)}以下` : "-",  // 簡易ケース基準
              standardCasesCount > 0 ? `複雑性スコア${simpleThreshold.toFixed(1)}～${complexThreshold.toFixed(1)}` : "-", // 標準ケース基準
              complexCasesCount > 0 ? `${p75Time.toFixed(1)}分以上、複雑性スコア${complexThreshold.toFixed(1)}以上` : "-",  // 複雑ケース基準
              simpleLevel,                           // 簡易ケースレベル
              standardLevel,                         // 標準ケースレベル
              complexLevel,                          // 複雑ケースレベル
            ]);
          }
        });
        
        // 結果を書き込み
        distributionSheet.clear();
        
        const outputheaders = [
          "タグ名", "件数", "総Rally数", "基本難易度", "月間件数", "月間Rally数",
          "時間P25", "時間中央値", "時間P75",
          "RallyP25", "Rally中央値", "RallyP75",
          "マクロ使用率", "相談率",
          "簡易ケース件数", "標準ケース件数", "複雑ケース件数",
          "簡易ケース割合", "標準ケース割合", "複雑ケース割合",
          "簡易ケース基準", "標準ケース基準", "複雑ケース基準",
          "簡易ケースレベル", "標準ケースレベル", "複雑ケースレベル"
        ];
        
        distributionSheet.getRange(1, 1, 1, outputheaders.length).setValues([outputheaders])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        if (results.length > 0) {
          // 件数順にソート
          results.sort((a, b) => b[1] - a[1]);
          
          distributionSheet.getRange(2, 1, results.length, outputheaders.length).setValues(results);
          // 条件付き書式の設定
          this.setDistributionSheetFormatting(distributionSheet);
        }
        
        LogModule.logOperation("タグ内難易度分布分析", `完了: ${results.length}件のタグを分析`);
        UtilsModule.showToast(`${results.length}件のタグ内難易度分布分析が完了しました！`, 'タグ内難易度分布分析');
        
        return results;
      } catch (e) {
        LogModule.handleError(e, 'タグ内難易度分布分析');
        return null;
      }
    },
    
    /**
     * タグ内難易度分布シートの条件付き書式を設定する
     */
    setDistributionSheetFormatting: function(sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return;
      
      // 基本難易度列の書式設定
      const baseLevelRange = sheet.getRange(2, 4, lastRow - 1, 1);
      
      // L1-L5の色分け
      const blueColors = ['#E3F2FD', '#90CAF9', '#42A5F5', '#1976D2', '#0D47A1'];
      
      for (let i = 1; i <= 5; i++) {
        const level = 'L' + i;
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(level)
          .setBackground(blueColors[i - 1])
          .setFontColor(i >= 4 ? 'white' : 'black')
          .setRanges([baseLevelRange])
          .build();
        
        let rules = sheet.getConditionalFormatRules();
        rules.push(rule);
        sheet.setConditionalFormatRules(rules);
      }
      
      // ケース割合の書式設定
      const caseRateRanges = [
        {range: sheet.getRange(2, 18, lastRow - 1, 1), color: '#E3F2FD'}, // 簡易ケース割合
        {range: sheet.getRange(2, 19, lastRow - 1, 1), color: '#90CAF9'}, // 標準ケース割合
        {range: sheet.getRange(2, 20, lastRow - 1, 1), color: '#1976D2'}  // 複雑ケース割合
      ];
      
      caseRateRanges.forEach(item => {
        const highRateRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("50")
          .setBackground(item.color)
          .setRanges([item.range])
          .build();
        
        let rules = sheet.getConditionalFormatRules();
        rules.push(highRateRule);
        sheet.setConditionalFormatRules(rules);
      });
      
      // 難易度レベル列の書式設定（簡易、標準、複雑）
      const levelColumns = [23, 24, 25]; // 列インデックス
      
      levelColumns.forEach(col => {
        const range = sheet.getRange(2, col, lastRow - 1, 1);
        
        // L1-L5の色分け
        const blueColors = ['#E3F2FD', '#90CAF9', '#42A5F5', '#1976D2', '#0D47A1'];
        
        for (let i = 1; i <= 5; i++) {
          const level = 'L' + i;
          const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(level)
            .setBackground(blueColors[i - 1])
            .setFontColor(i >= 4 ? 'white' : 'black')
            .setRanges([range])
            .build();
          
          const rules = sheet.getConditionalFormatRules();
          rules.push(rule);
          sheet.setConditionalFormatRules(rules);
        }
      });
      
      // 月間件数の書式設定
      const monthlyCountRange = sheet.getRange(2, 5, lastRow - 1, 1);
      
      const highCountRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(50)
        .setBackground('#C8E6C9') // 緑
        .setRanges([monthlyCountRange])
        .build();
      
      const lowCountRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(10)
        .setBackground('#FFEBEE') // 薄い赤
        .setRanges([monthlyCountRange])
        .build();
      
      let rules = sheet.getConditionalFormatRules();
      rules.push(highCountRule, lowCountRule);
      sheet.setConditionalFormatRules(rules);
      
      // マクロ使用率と相談率の書式設定
      const macroRateRange = sheet.getRange(2, 13, lastRow - 1, 1);
      
      const highMacroRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("50")
        .setBackground('#E8F5E9') // 薄い緑
        .setRanges([macroRateRange])
        .build();
      
      rules = sheet.getConditionalFormatRules();
      rules.push(highMacroRule);
      sheet.setConditionalFormatRules(rules);
      
      const consultRateRange = sheet.getRange(2, 14, lastRow - 1, 1);
      
      const highConsultRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("30")
        .setBackground('#FFF3E0') // 薄いオレンジ
        .setRanges([consultRateRange])
        .build();
      
      rules = sheet.getConditionalFormatRules();
      rules.push(highConsultRule);
      sheet.setConditionalFormatRules(rules);
      
      // 列幅の調整
      sheet.setColumnWidth(1, 200); // タグ名
      sheet.autoResizeColumns(2, 24); // 他の列を自動調整
    },
  
    /**
     * 個人スキル分析を実行する
     */
    analyzeIndividualSkills: function() {
      LogModule.logOperation("個人スキル分析", "開始");
      
      try {
        // データ準備
        const data = DataModule.prepareAnalysisData();
        if (!data) {
          throw new Error('分析データの準備に失敗しました');
        }
        
        // 設定値取得
        const settings = SettingsModule.getSettings();
        
        // シート取得
        const skillSheet = UtilsModule.getOrCreateSheet(SHEETS.INDIVIDUAL_SKILL);
        
        // データ取得
        const transactionData = data.transactionData;
        const txColIndexes = data.txColIndexes;
        const tagDifficultyMap = data.tagDifficultyMap;
        const tagDistributionMap = data.tagDistributionMap;
        const responderMasterMap = data.responderMasterMap;
        
        // 全体の相談ケース数をカウントするための変数
        let totalConsultationCases = 0;
        
        // 1回目のループ：全体の相談ケース数を集計
        for (let i = 1; i < transactionData.length; i++) {
          const row = transactionData[i];
          
          // 相談フラグ（関連カラムに値があるかどうか）
          const hasConsultation = 
            (txColIndexes.soudanUserName !== -1 && row[txColIndexes.soudanUserName]);
          
          // 相談ケースの集計
          if (hasConsultation) {
            totalConsultationCases++;
          }
        }
        
        // 個人スキルデータ集計のための準備
        const responderStats = {};
        
        // 個人スキルデータ集計
        for (let i = 1; i < transactionData.length; i++) {
          const row = transactionData[i];
          const tagName = row[txColIndexes.tagName];
          const userName = row[txColIndexes.userName];
          const ihPt = row[txColIndexes.ihPt];
          const div = row[txColIndexes.div];
          const opName = row[txColIndexes.opName];
          const responseTime = row[txColIndexes.responseTime];
          const rallyCount = txColIndexes.rallyCount !== -1 ? (row[txColIndexes.rallyCount] || 1) : 1;
          
          // OJTタグの除外（オプション）
          const isOjtCase = txColIndexes.hasOjtTag !== -1 && row[txColIndexes.hasOjtTag] === "TRUE";
          if (isOjtCase && settings && settings.excludeOjtData === true) {
            continue;
          }
          
          // マクロ使用フラグ
          const useMacro = txColIndexes.useMacro !== -1 ? (row[txColIndexes.useMacro] === "マクロ使用") : false;
          
          // 機能タグのみを処理
          if (tagName && UtilsModule.isFunctionTag(tagName) && userName) {
            if (!responderStats[userName]) {
              responderStats[userName] = {
                name: userName,
                ihPt: ihPt,
                div: div,
                primaryTeam: responderMasterMap[userName] ? responderMasterMap[userName].primaryTeam : "",
                secondaryTeam: responderMasterMap[userName] ? responderMasterMap[userName].secondaryTeam : "",
                workTimeRatio: responderMasterMap[userName] && responderMasterMap[userName].ihPtRatio ? 
                              responderMasterMap[userName].ihPtRatio : 
                              (ihPt === 'IH' ? settings.ihDefaultHours : settings.ptDefaultHours),
                tags: new Set(),
                tagResponses: {},
                levelCounts: { L1: 0, L2: 0, L3: 0, L4: 0, L5: 0 },
                totalResponseTime: 0,
                totalRallyCount: 0,
                totalResponses: 0,
                macroUsageCount: 0,
                primaryTeamResponses: 0,
                secondaryTeamResponses: 0,
                externalTeamResponses: 0,
                adviceProvided: 0,
                questionsAsked: 0
              };
            }
            
            responderStats[userName].tags.add(tagName);
            responderStats[userName].totalResponses++;
            responderStats[userName].totalRallyCount += rallyCount;
            
            // マクロ使用のカウント
            if (useMacro) {
              responderStats[userName].macroUsageCount++;
            }
                    
            // タグの担当OPチームと所属OPチームの一致判定
            const tagOPTeam = opName || (tagDifficultyMap[tagName] ? tagDifficultyMap[tagName].opTeam : "");
            
            if (tagOPTeam === responderStats[userName].primaryTeam) {
              responderStats[userName].primaryTeamResponses++;
            } else if (responderStats[userName].secondaryTeam && 
                      responderStats[userName].secondaryTeam.split(',').some(team => team.trim() === tagOPTeam)) {
              responderStats[userName].secondaryTeamResponses++;
            } else {
              responderStats[userName].externalTeamResponses++;
            }
            
            // 難易度ベースの集計
            const tagDifficulty = tagDifficultyMap[tagName] || { 
              level: 'L3', 
              score: 3,
              tcLevel: 'L3',
              rdLevel: 'L3',
              krLevel: 'L3'
            };
            
            // タグごとの対応詳細
            if (!responderStats[userName].tagResponses[tagName]) {
              responderStats[userName].tagResponses[tagName] = {
                count: 0,
                rallyCount: 0,
                responseTimes: [],
                difficultyLevel: tagDifficulty.level,
                tcLevel: tagDifficulty.tcLevel,
                rdLevel: tagDifficulty.rdLevel,
                krLevel: tagDifficulty.krLevel,
                macroCount: 0,
                consultCount: 0,
                simpleCases: 0,
                standardCases: 0,
                complexCases: 0
              };
            }
            
            responderStats[userName].tagResponses[tagName].count++;
            responderStats[userName].tagResponses[tagName].rallyCount += rallyCount;
            
            if (useMacro) {
              responderStats[userName].tagResponses[tagName].macroCount++;
            }
            
            // 相談フラグ（関連カラムに値があるかどうか）
            const hasConsultation = 
              (txColIndexes.postTimestamp !== -1 && row[txColIndexes.postTimestamp]) || 
              (txColIndexes.soudanUserName !== -1 && row[txColIndexes.soudanUserName]) || 
              (txColIndexes.adviserUserName !== -1 && row[txColIndexes.adviserUserName]);
            
            if (hasConsultation) {
              responderStats[userName].tagResponses[tagName].consultCount++;
            }
            
            // 応答時間の集計（異常値を除外）
            if (responseTime > 0 && responseTime < 1440) {
              responderStats[userName].totalResponseTime += responseTime;
              responderStats[userName].tagResponses[tagName].responseTimes.push(responseTime);
              
              // タグ内難易度分布に基づくケースタイプ判定
              // 複雑性スコアに基づく判定
              // 1. 対応時間（40%）、Rally数（30%）、マクロ非使用（20%）、相談あり（10%）
              const timeScore = UtilsModule.calculateTimeScore(responseTime);
              const rallyScore = UtilsModule.calculateRallyScore(rallyCount);
              const macroScore = useMacro ? 1 : 3; // マクロ使用なら簡易(1)、使用なしなら複雑(3)
              const consultScore = hasConsultation ? 5 : 1; // 相談ありなら複雑(5)、なしなら簡易(1)
              
              const complexityScore = 
                timeScore * 0.4 + 
                rallyScore * 0.3 + 
                macroScore * 0.2 + 
                consultScore * 0.1;
                
              if (tagDistributionMap[tagName]) {
                // 複雑性スコアの閾値を抽出（文字列から数値部分を抽出）
                let simpleThreshold = 0;
                let complexThreshold = 5;
                
                const simpleThresholdStr = tagDistributionMap[tagName].simpleThreshold;
                const complexThresholdStr = tagDistributionMap[tagName].complexThreshold;
                
                if (simpleThresholdStr && simpleThresholdStr.includes("複雑性スコア")) {
                  const match = simpleThresholdStr.match(/複雑性スコア(\d+\.?\d*)/);
                  if (match && match[1]) {
                    simpleThreshold = parseFloat(match[1]);
                  }
                }
                
                if (complexThresholdStr && complexThresholdStr.includes("複雑性スコア")) {
                  const match = complexThresholdStr.match(/複雑性スコア(\d+\.?\d*)/);
                  if (match && match[1]) {
                    complexThreshold = parseFloat(match[1]);
                  }
                }
                
                // ケースタイプの判定
                if (complexityScore <= simpleThreshold) {
                  responderStats[userName].tagResponses[tagName].simpleCases++;
                } else if (complexityScore >= complexThreshold) {
                  responderStats[userName].tagResponses[tagName].complexCases++;
                } else {
                  responderStats[userName].tagResponses[tagName].standardCases++;
                }
              }
            }
          }
        }
        
        // 2回目のループ：相談対応数と質問数を集計
        for (let i = 1; i < transactionData.length; i++) {
          const row = transactionData[i];
          
          // すべての担当者に対して処理
          for (const responderName in responderStats) {
            // 相談対応のカウント - adviser_user_nameが担当者と一致し、soudan_user_nameが担当者と一致しない場合
            if (txColIndexes.adviserUserName !== -1 && 
                row[txColIndexes.adviserUserName] === responderName &&
                txColIndexes.soudanUserName !== -1 && 
                row[txColIndexes.soudanUserName] !== responderName) {
              responderStats[responderName].adviceProvided++;
            }
            
            // 質問のカウント - soudan_user_nameが担当者と一致する場合
            if (txColIndexes.soudanUserName !== -1 && 
                row[txColIndexes.soudanUserName] === responderName) {
              responderStats[responderName].questionsAsked++;
            }
          }
        }
        
        // 最高対応レベルの計算と結果データの作成
        const results = [];
        
        Object.values(responderStats).forEach(stats => {
          // タグ対応レベルの集計
          let l1Tags = 0, l2Tags = 0, l3Tags = 0, l4Tags = 0, l5Tags = 0;
          let tcL4L5Tags = 0; // 技術的複雑性がL4/L5のタグ数
          let complexCasesTags = 0; // 複雑ケース対応があるタグ数
          
          Object.entries(stats.tagResponses).forEach(([tagName, tagStats]) => {
            let highestLevel = "L1";
            
            // タグ内難易度分布を考慮した最高対応レベル判定
            if (tagDistributionMap[tagName]) {
              if (tagStats.complexCases > 0) {
                highestLevel = tagDistributionMap[tagName].complexLevel;
                complexCasesTags++;
              } else if (tagStats.standardCases > 0) {
                highestLevel = tagDistributionMap[tagName].standardLevel;
              } else if (tagStats.simpleCases > 0) {
                highestLevel = tagDistributionMap[tagName].simpleLevel;
              }
            } else {
              // 分布データがない場合はタグの基本難易度を使用
              highestLevel = tagDifficultyMap[tagName]?.level || "L3";
            }
            
            // 技術的複雑性がL4/L5のタグをカウント
            const tcLevel = tagDifficultyMap[tagName]?.tcLevel || "L3";
            if (tcLevel === "L4" || tcLevel === "L5") {
              tcL4L5Tags++;
            }
            
            // レベル別タグ数をカウント
            switch (highestLevel) {
              case "L1": l1Tags++; break;
              case "L2": l2Tags++; break;
              case "L3": l3Tags++; break;
              case "L4": l4Tags++; break;
              case "L5": l5Tags++; break;
            }
          });
          
          // スキル指標計算
          const totalTags = Object.keys(tagDifficultyMap).length || 1; // ゼロ除算防止
          const widthScore = stats.tags.size / totalTags * 100;
          
          // 深さスコア計算（難易度の重み付け合計）
          const depthNumerator = l1Tags * 1 + l2Tags * 2 + l3Tags * 3 + l4Tags * 4 + l5Tags * 5;
          const depthDenominator = l1Tags + l2Tags + l3Tags + l4Tags + l5Tags;
          const depthScore = depthDenominator > 0 ? depthNumerator / depthDenominator : 0;
          
          // Rally生産性計算
          let productivityScore = 0;
          if (stats.totalResponseTime > 0) {
            productivityScore = (stats.totalRallyCount / stats.totalResponseTime) * 60; // 1時間あたりのRally数
          }
          
          // チーム一致率
          const teamMatchRate = stats.totalResponses > 0 ? 
                              (stats.primaryTeamResponses + stats.secondaryTeamResponses) / 
                              stats.totalResponses * 100 : 0;
          
          // マクロ使用率
          const macroUsageRate = stats.totalResponses > 0 ?
                               stats.macroUsageCount / stats.totalResponses * 100 : 0;
          
          // 質問率（自分の対応した案件のうち、質問した割合）
          const questionRate = stats.totalResponses > 0 ?
                              stats.questionsAsked / stats.totalResponses * 100 : 0;
          
          // 相談対応率（全体の相談案件のうち、相談に乗った割合）
          const adviserRate = totalConsultationCases > 0 ?
                              stats.adviceProvided / totalConsultationCases * 100 : 0;
          
          // 相談コスト考慮済み効果的生産性の計算
          const effectiveProductivity = UtilsModule.calculateEffectiveProductivity(
            productivityScore, questionRate, adviserRate, settings
          );
          
          // パフォーマータイプ判定
          let performerType;
          
          // 設定からしきい値を取得
          const hyperperformerWidthThreshold = settings.hyperperformerWidthThreshold || 75;
          const hyperperformerProdThreshold = settings.hyperperformerProdThreshold || 10;
          const specialistDepthThreshold = settings.specialistDepthThreshold || 3.5;
          const specialistComplexThreshold = settings.specialistComplexThreshold || 3;
          const allrounderWidthThreshold = settings.allrounderWidthThreshold || 70;
    
          // ハイパフォーマー：幅広いタグに対応し高い生産性を持つ
          if (widthScore >= hyperperformerWidthThreshold && productivityScore >= hyperperformerProdThreshold) {
            performerType = "ハイパフォーマー";
          }
          // スペシャリスト：特定OPチーム/機能領域の高難易度案件に強い
          else if (depthScore >= specialistDepthThreshold && complexCasesTags >= specialistComplexThreshold && (widthScore < 50 || tcL4L5Tags >= 3)) {
            performerType = "スペシャリスト";
          }
          // オールラウンダー：幅広い対応領域を持つが生産性は標準的
          else if (widthScore >= allrounderWidthThreshold) {
            performerType = "オールラウンダー";
          }
          // 標準：主に所属OPチームの標準的な機能に対応
          else {
            performerType = "標準";
          }
          
          // 総合スキルスコア計算
          const totalSkillScore = widthScore * 0.3 + depthScore * 20 * 0.3 + productivityScore * 3 * 0.4;
          
          results.push([
            stats.name,                       // 担当者名
            stats.ihPt,                       // IH/PT区分
            stats.div,                        // 職種
            stats.primaryTeam,                // 主所属OPチーム
            stats.secondaryTeam,              // 兼務OPチーム
            stats.workTimeRatio,              // 稼働時間係数
            stats.tags.size,                  // 対応タグ数
            stats.totalRallyCount,            // 総Rally数
            l1Tags,                           // L1タグ対応数
            l2Tags,                           // L2タグ対応数
            l3Tags,                           // L3タグ対応数
            l4Tags,                           // L4タグ対応数
            l5Tags,                           // L5タグ対応数
            widthScore.toFixed(1),            // 幅スコア
            depthScore.toFixed(2),            // 深さスコア
            productivityScore.toFixed(1),     // 生産性スコア（Rally/時間）
            effectiveProductivity.toFixed(1), // 効果的生産性スコア
            teamMatchRate.toFixed(1),         // チーム一致率(%)
            macroUsageRate.toFixed(1),        // マクロ使用率(%)
            adviserRate.toFixed(1),           // 相談対応率(%)
            questionRate.toFixed(1),          // 質問率(%)
            stats.adviceProvided,             // 相談対応数
            stats.questionsAsked,             // 質問数
            complexCasesTags,                 // 複雑ケース対応タグ数
            totalSkillScore.toFixed(1),       // 総合スキルスコア
            performerType,                    // パフォーマータイプ
            new Date()                        // 更新日時
          ]);
        });
        
        // 結果を書き込み
        skillSheet.clear();
        
        const outputheaders = [
          "担当者名", "IH/PT区分", "職種", "主所属OPチーム", "兼務OPチーム", "稼働時間係数",
          "対応タグ数", "総Rally数", "L1タグ対応数", "L2タグ対応数", "L3タグ対応数", "L4タグ対応数", "L5タグ対応数",
          "幅スコア", "深さスコア", "生産性スコア", "効果的生産性", "チーム一致率(%)", "マクロ使用率(%)", 
          "相談対応率(%)", "質問率(%)", "相談対応数", "質問数", "複雑ケース対応タグ数",
          "総合スキルスコア", "パフォーマータイプ", "更新日時"
        ];
        
        skillSheet.getRange(1, 1, 1, outputheaders.length).setValues([outputheaders])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        if (results.length > 0) {
          skillSheet.getRange(2, 1, results.length, outputheaders.length).setValues(results);
          // 条件付き書式の設定
          this.setIndividualSkillFormatting(skillSheet);
        }
        
        LogModule.logOperation("個人スキル分析", `完了: ${results.length}名の担当者を分析`);
        UtilsModule.showToast(`${results.length}名の個人スキル分析が完了しました！`, '個人スキル分析');
        
        return results;
      } catch (e) {
        LogModule.handleError(e, '個人スキル分析');
        return null;
      }
    },
      
    /**
     * 個人スキルシートの条件付き書式を設定する
     */
    setIndividualSkillFormatting: function(sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return;
      
      // スコア列の書式設定（幅、深さ、生産性、総合）
      const scoreColumns = [14, 15, 16, 17, 25]; // 列インデックス
      
      scoreColumns.forEach(col => {
        const range = sheet.getRange(2, col, lastRow - 1, 1);
        
        // グラデーションでの色分け
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .setGradientMinpoint('#FFFFFF')
          .setGradientMaxpoint('#26C6DA') // 水色
          .setRanges([range])
          .build();
        
        const rules = sheet.getConditionalFormatRules();
        rules.push(rule);
        sheet.setConditionalFormatRules(rules);
      });
      
      // マクロ使用率の書式設定
      const macroRateRange = sheet.getRange(2, 19, lastRow - 1, 1);
      const highMacroRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(50)
        .setBackground('#E8F5E9') // 薄い緑
        .setRanges([macroRateRange])
        .build();
      
      let rules = sheet.getConditionalFormatRules();
      rules.push(highMacroRule);
      sheet.setConditionalFormatRules(rules);
      
      // 相談関連列の書式設定
      // 相談対応率
      const adviserRateRange = sheet.getRange(2, 20, lastRow - 1, 1);
      const highAdviserRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(10)
        .setBackground('#FFF3E0') // 薄いオレンジ
        .setRanges([adviserRateRange])
        .build();
      
      rules = sheet.getConditionalFormatRules();
      rules.push(highAdviserRule);
      sheet.setConditionalFormatRules(rules);
      
      // 質問率
      const questionRateRange = sheet.getRange(2, 21, lastRow - 1, 1);
      const highQuestionRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(20)
        .setBackground('#FFE0B2') // オレンジ
        .setRanges([questionRateRange])
        .build();
      
      rules = sheet.getConditionalFormatRules();
      rules.push(highQuestionRule);
      sheet.setConditionalFormatRules(rules);
      
      // 複雑ケース対応タグ数
      const complexCasesRange = sheet.getRange(2, 24, lastRow - 1, 1);
      const highComplexCasesRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(3)
        .setBackground('#E1F5FE') // 薄い青
        .setRanges([complexCasesRange])
        .build();
      
      rules = sheet.getConditionalFormatRules();
      rules.push(highComplexCasesRule);
      sheet.setConditionalFormatRules(rules);
      
      // パフォーマータイプの書式設定
      const typeRange = sheet.getRange(2, 26, lastRow - 1, 1);
      
      // タイプ別の色分け
      const typeRules = [
        { value: 'ハイパフォーマー', color: '#1565C0' }, // 濃い青
        { value: 'スペシャリスト', color: '#6A1B9A' },   // 紫
        { value: 'オールラウンダー', color: '#2E7D32' }, // 緑
        { value: '標準', color: '#757575' }              // グレー
      ];
      
      typeRules.forEach(type => {
        const rule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(type.value)
          .setBackground(type.color)
          .setFontColor('white')
          .setRanges([typeRange])
          .build();
        
        const rules = sheet.getConditionalFormatRules();
        rules.push(rule);
        sheet.setConditionalFormatRules(rules);
      });
      
      // 列幅の調整
      sheet.setColumnWidth(1, 150);  // 担当者名
      sheet.setColumnWidth(2, 80);   // IH/PT区分
      sheet.setColumnWidth(3, 100);  // 職種
      sheet.setColumnWidth(4, 150);  // 主所属OPチーム
      sheet.setColumnWidth(5, 200);  // 兼務OPチーム
      
      // 残りの列は自動調整
      sheet.autoResizeColumns(7, 20);
    },
  
    /**
     * チームパフォーマンス分析を実行する
     */
    analyzeTeamPerformance: function() {
      LogModule.logOperation("チームパフォーマンス分析", "開始");
      
      try {
        // データ準備
        const data = DataModule.prepareAnalysisData();
        if (!data) {
          throw new Error('分析データの準備に失敗しました');
        }
        
        // 設定値取得
        const settings = SettingsModule.getSettings();
        
        // シート取得
        const performanceSheet = UtilsModule.getOrCreateSheet(SHEETS.TEAM_PERFORMANCE);
        
        // チーム一覧の取得
        const teams = UtilsModule.getTeamList(data.sheets.masterSheet);
        
        // 担当者マスターのマッピング
        const memberMap = data.responderMasterMap;
        
        // チームごとのパフォーマンスデータ
        const teamPerformance = {};
        
        // チームの初期化
        teams.forEach(team => {
          teamPerformance[team] = {
            teamName: team,
            totalTags: 0,
            
            // 対応状況（Rallyベース）
            totalRallyCount: 0,
            
            // 対応区分の細分化
            ihPrimaryResponses: { rallyCount: 0, responders: new Set(), responseTime: 0, questionCount: 0, adviceCount: 0 },
            ihSecondaryResponses: { rallyCount: 0, responders: new Set(), responseTime: 0, questionCount: 0, adviceCount: 0 },
            ptPrimaryResponses: { rallyCount: 0, responders: new Set(), responseTime: 0, questionCount: 0, adviceCount: 0 },
            ptSecondaryResponses: { rallyCount: 0, responders: new Set(), responseTime: 0, questionCount: 0, adviceCount: 0 },
            
            // 時間データ
            totalResponseTime: 0,
            
            // 相談データ
            totalConsultations: 0,
            totalSolveDuration: 0,
            
            // メンバーデータ
            memberPerformance: {},
            
            // メンバー分類
            primaryMembers: [],
            secondaryMembers: []
          };
        });
        
        // タグの所属チーム情報マッピング
        const tagTeamMap = {};
        const diffData = data.difficultyData;
        const diffColIndexes = data.diffColIndexes;
        
        for (let i = 1; i < diffData.length; i++) {
          const tagName = diffData[i][diffColIndexes.tagName];
          const opTeam = diffData[i][diffColIndexes.opTeam];
          if (tagName && opTeam) {
            tagTeamMap[tagName] = opTeam;
            
            // チームごとのタグ数カウント
            if (teamPerformance[opTeam]) {
              teamPerformance[opTeam].totalTags++;
            }
          }
        }
        
        // メンバーの所属チーム情報を設定
        const masterData = data.masterData;
        const masterColIndexes = data.masterColIndexes;
        
        for (let i = 1; i < masterData.length; i++) {
          const name = masterData[i][masterColIndexes.name];
          const primaryTeam = masterData[i][masterColIndexes.primaryTeam];
          const ihPt = masterData[i][masterColIndexes.ihPt];
          
          if (name && primaryTeam && teamPerformance[primaryTeam]) {
            // メンバーの初期化
            teamPerformance[primaryTeam].memberPerformance[name] = {
              name: name,
              ihPt: ihPt,
              rallyCount: 0,
              responseTime: 0,
              productivity: 0,
              questionCount: 0,
              adviceCount: 0
            };
            
            // チームのメンバーリストに追加
            if (ihPt === 'IH') {
              teamPerformance[primaryTeam].primaryMembers.push(name);
            }
          }
          
          // 兼務チームの処理
          const secondaryTeam = masterData[i][masterColIndexes.secondaryTeam];
          if (name && secondaryTeam) {
            const secondaryTeams = secondaryTeam.split(',').map(t => t.trim());
            
            secondaryTeams.forEach(team => {
              if (team && teamPerformance[team]) {
                // 兼務先チームのメンバーパフォーマンスも初期化
                if (!teamPerformance[team].memberPerformance[name]) {
                  teamPerformance[team].memberPerformance[name] = {
                    name: name,
                    ihPt: ihPt,
                    rallyCount: 0,
                    responseTime: 0,
                    productivity: 0,
                    questionCount: 0,
                    adviceCount: 0,
                    isPrimary: false
                  };
                }
                
                // 兼務メンバーリストに追加
                if (ihPt === 'IH') {
                  teamPerformance[team].secondaryMembers.push(name);
                }
              }
            });
          }
        }
        
        // 対応データの集計
        const txData = data.transactionData;
        const txColIndexes = data.txColIndexes;
        
        // 相談ケース総数のカウント（質問率・相談対応率の分母用）
        let totalConsultationCases = 0;
        for (let i = 1; i < txData.length; i++) {
          const row = txData[i];
          if (txColIndexes.soudanUserName !== -1 && row[txColIndexes.soudanUserName]) {
            totalConsultationCases++;
          }
        }
        
        // 対応データの集計
        for (let i = 1; i < txData.length; i++) {
          const row = txData[i];
          const tagName = row[txColIndexes.tagName];
          const userName = row[txColIndexes.userName];
          const ihPt = row[txColIndexes.ihPt];
          const opName = row[txColIndexes.opName] || tagTeamMap[tagName];
          const responseTime = row[txColIndexes.responseTime];
          const rallyCount = txColIndexes.rallyCount !== -1 ? (row[txColIndexes.rallyCount] || 1) : 1;
          
          // OJTタグの除外（オプション）
          const isOjtCase = txColIndexes.hasOjtTag !== -1 && row[txColIndexes.hasOjtTag] === "TRUE";
          if (isOjtCase && settings && settings.excludeOjtData === true) {
            continue;
          }
          
          // 機能タグのみ処理
          if (tagName && UtilsModule.isFunctionTag(tagName) && opName && teams.includes(opName)) {
            const team = opName;
            
            // 対応区分判定
            let category;
            
            // IH/PT区分による分類
            if (ihPt === 'IH') {
              // 所属チーム判定
              const memberInfo = memberMap[userName];
              const primaryTeam = memberInfo ? memberInfo.primaryTeam : "";
              const secondaryTeams = memberInfo && memberInfo.secondaryTeam ? 
                                    memberInfo.secondaryTeam.split(',').map(t => t.trim()) : [];
              
              if (primaryTeam === team) {
                category = 'ihPrimary'; // IH所属チーム
              } else if (secondaryTeams.includes(team)) {
                category = 'ihSecondary'; // IH兼務チーム
              } else {
                category = 'ihSecondary'; // IH他チーム
              }
            } else {
              // PT所属チーム判定
              const memberInfo = memberMap[userName];
              const ptPrimaryTeam = memberInfo ? memberInfo.primaryTeam : "";
              
              if (ptPrimaryTeam === team) {
                category = 'ptPrimary'; // PT所属チーム
              } else {
                category = 'ptSecondary'; // PT他チーム
              }
            }
            
            // 時間データが有効な場合のみ集計
            if (responseTime > 0 && responseTime < 1440) {
              // チーム全体の集計
              teamPerformance[team].totalRallyCount += rallyCount;
              teamPerformance[team].totalResponseTime += responseTime;
              
              // 対応区分別の集計
              if (teamPerformance[team][`${category}Responses`]) {
                teamPerformance[team][`${category}Responses`].rallyCount += rallyCount;
                teamPerformance[team][`${category}Responses`].responders.add(userName);
                teamPerformance[team][`${category}Responses`].responseTime += responseTime;
              }
              
              // メンバー個別のパフォーマンス集計
              if (userName && teamPerformance[team].memberPerformance[userName]) {
                teamPerformance[team].memberPerformance[userName].rallyCount += rallyCount;
                teamPerformance[team].memberPerformance[userName].responseTime += responseTime;
              }
            }
          }
        }
        
        // 相談データの集計（質問と回答）
        for (let i = 1; i < txData.length; i++) {
          const row = txData[i];
          const tagName = row[txColIndexes.tagName];
          const opName = row[txColIndexes.opName] || tagTeamMap[tagName];
          
          // 質問者と回答者の情報取得
          const soudanUserName = txColIndexes.soudanUserName !== -1 ? row[txColIndexes.soudanUserName] : null;
          const adviserUserName = txColIndexes.adviserUserName !== -1 ? row[txColIndexes.adviserUserName] : null;
          const solveDuration = txColIndexes.solveDuration !== -1 ? row[txColIndexes.solveDuration] || 0 : 0;
          
          if (tagName && UtilsModule.isFunctionTag(tagName) && opName && teams.includes(opName)) {
            const team = opName;
            
            // 相談があった場合
            if (soudanUserName && adviserUserName) {
              teamPerformance[team].totalConsultations++;
              teamPerformance[team].totalSolveDuration += solveDuration;
              
              // 質問者の対応区分判定と集計
              let askerCategory;
              if (soudanUserName && memberMap[soudanUserName]) {
                const ihPt = memberMap[soudanUserName].ihPt || "";
                const primaryTeam = memberMap[soudanUserName].primaryTeam || "";
                const secondaryTeams = memberMap[soudanUserName].secondaryTeam ? 
                                      memberMap[soudanUserName].secondaryTeam.split(',').map(t => t.trim()) : [];
                
                if (ihPt === 'IH') {
                  if (primaryTeam === team) {
                    askerCategory = 'ihPrimary';
                  } else if (secondaryTeams.includes(team)) {
                    askerCategory = 'ihSecondary';
                  } else {
                    askerCategory = 'ihSecondary';
                  }
                } else {
                  if (primaryTeam === team) {
                    askerCategory = 'ptPrimary';
                  } else {
                    askerCategory = 'ptSecondary';
                  }
                }
                
                if (teamPerformance[team][`${askerCategory}Responses`]) {
                  teamPerformance[team][`${askerCategory}Responses`].questionCount++;
                }
              }
              
              // 回答者の対応区分判定と集計
              let adviserCategory;
              if (adviserUserName && memberMap[adviserUserName]) {
                const ihPt = memberMap[adviserUserName].ihPt || "";
                const primaryTeam = memberMap[adviserUserName].primaryTeam || "";
                const secondaryTeams = memberMap[adviserUserName].secondaryTeam ? 
                                      memberMap[adviserUserName].secondaryTeam.split(',').map(t => t.trim()) : [];
                
                if (ihPt === 'IH') {
                  if (primaryTeam === team) {
                    adviserCategory = 'ihPrimary';
                  } else if (secondaryTeams.includes(team)) {
                    adviserCategory = 'ihSecondary';
                  } else {
                    adviserCategory = 'ihSecondary';
                  }
                } else {
                  if (primaryTeam === team) {
                    adviserCategory = 'ptPrimary';
                  } else {
                    adviserCategory = 'ptSecondary';
                  }
                }
                
                if (teamPerformance[team][`${adviserCategory}Responses`]) {
                  teamPerformance[team][`${adviserCategory}Responses`].adviceCount++;
                }
              }
              
              // メンバー個別の質問・回答集計
              if (teamPerformance[team].memberPerformance[soudanUserName]) {
                teamPerformance[team].memberPerformance[soudanUserName].questionCount++;
              }
              
              if (teamPerformance[team].memberPerformance[adviserUserName]) {
                teamPerformance[team].memberPerformance[adviserUserName].adviceCount++;
              }
            }
          }
        }
        
        // 集計結果の計算
        Object.values(teamPerformance).forEach(team => {
          // 各対応区分の生産性と割合の計算
          const categories = ['ihPrimary', 'ihSecondary', 'ptPrimary', 'ptSecondary'];
          const totalRally = team.totalRallyCount || 1; // ゼロ除算防止
          
          categories.forEach(category => {
            const responses = team[`${category}Responses`];
            
            // Rally生産性の計算（Rally/時間）
            responses.productivity = responses.responseTime > 0 ? 
                                 (responses.rallyCount / responses.responseTime) * 60 : 0;
            
            // 対応割合の計算
            responses.rallyRate = (responses.rallyCount / totalRally) * 100;
            
            // 質問率の計算
            responses.questionRate = responses.rallyCount > 0 ? 
                                (responses.questionCount / responses.rallyCount) * 100 : 0;
            
            // 相談対応率の計算
            responses.adviceRate = team.totalConsultations > 0 ? 
                               (responses.adviceCount / team.totalConsultations) * 100 : 0;
            
            // 相談コスト係数を考慮した効果的な生産性
            responses.effectiveProductivity = UtilsModule.calculateEffectiveProductivity(
              responses.productivity, responses.questionRate, responses.adviceRate, settings
            );
          });
          
          // 各メンバーの生産性計算
          Object.values(team.memberPerformance).forEach(member => {
            // Rally生産性計算（Rally/時間）
            member.productivity = member.responseTime > 0 ? 
                              (member.rallyCount / member.responseTime) * 60 : 0;
            
            // 質問率の計算
            member.questionRate = member.rallyCount > 0 ? 
                             (member.questionCount / member.rallyCount) * 100 : 0;
            
            // 相談対応率の計算
            member.adviceRate = team.totalConsultations > 0 ? 
                           (member.adviceCount / team.totalConsultations) * 100 : 0;
            
            // 相談コスト係数を考慮した効果的な生産性
            member.effectiveProductivity = UtilsModule.calculateEffectiveProductivity(
              member.productivity, member.questionRate, member.adviceRate, settings
            );
          });
          
          // チーム全体の生産性計算（Rally/時間）
          team.totalProductivity = team.totalResponseTime > 0 ? 
                                 (team.totalRallyCount / team.totalResponseTime) * 60 : 0;
          
          // チーム内完結率（所属IHの対応率）
          team.completionRate = team.ihPrimaryResponses.rallyRate;
          
          // 相談コスト係数を考慮した効果的な生産性
          team.totalQuestionRate = team.totalRallyCount > 0 ? 
                                (team.totalConsultations / team.totalRallyCount) * 100 : 0;
          
          // 効果的生産性の計算
          team.effectiveProductivity = UtilsModule.calculateEffectiveProductivity(
            team.totalProductivity, team.totalQuestionRate, 0, settings
          );
          
          // 月間必要HCの計算
          // 月間稼働時間 = 稼働日数 × 1日あたり時間 × 生産性係数
          const monthlyWorkHours = settings.workDaysPerMonth * settings.hoursPerDay * settings.productivityRatio;
          
          // 月間Rally数推定 (3ヶ月分のデータから1ヶ月分を推定)
          const monthlyRallyCount = team.totalRallyCount / 3;
          
          // チーム内完結目標率
          const targetCompletionRate = settings.targetCompletionRate || 85;
          
          // 目標生産性
          const targetRallyPerHour = settings.targetRallyPerHour || 10;
          
          // チーム内完結分の必要HC
          team.idealHC = (monthlyRallyCount * (targetCompletionRate / 100)) / (targetRallyPerHour * monthlyWorkHours);
          
          // 相談コスト考慮済みの必要HC (オプション)
          team.adjustedIdealHC = team.idealHC * (1 + (settings.consultationCostFactor || 0) * (team.totalQuestionRate / 100));
          
          // HC不足度
          team.hcGap = team.primaryMembers.length - team.idealHC;
        });
        
        // チームパフォーマンスシートの作成
        performanceSheet.clear();
        
        // ヘッダーとタイトル
        performanceSheet.getRange("A1:M1").merge().setValue("チームパフォーマンス分析（Rally反映・相談コスト考慮版）")
          .setFontWeight('bold').setBackground('#E8EAF6').setFontSize(14);
        
        performanceSheet.getRange(2, 1).setValue("更新日: " + new Date().toLocaleDateString());
        
        // チーム概要セクション
        performanceSheet.getRange("A4:M4").merge().setValue("チーム概要")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
        
        const summaryHeaders = [
          "OPチーム", "対応タグ数", "Rally総数", "所属IHメンバー", "兼務IHメンバー",
          "チーム内完結率(%)", "IH他チーム対応率(%)", "PT所属対応率(%)", "PT他チーム対応率(%)",
          "生産性(Rally/h)", "効果的生産性", "必要HC", "HC余剰/不足"
        ];
        
        performanceSheet.getRange(5, 1, 1, summaryHeaders.length).setValues([summaryHeaders])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // チーム概要データ
        const summaryData = teams.map(teamName => {
          const team = teamPerformance[teamName];
          return [
            team.teamName,
            team.totalTags,
            team.totalRallyCount,
            team.primaryMembers.length,
            team.secondaryMembers.length,
            team.ihPrimaryResponses.rallyRate.toFixed(1) + "%",
            team.ihSecondaryResponses.rallyRate.toFixed(1) + "%",
            team.ptPrimaryResponses.rallyRate.toFixed(1) + "%",
            team.ptSecondaryResponses.rallyRate.toFixed(1) + "%",
            team.totalProductivity.toFixed(1),
            team.effectiveProductivity.toFixed(1),
            team.idealHC.toFixed(1),
            team.hcGap.toFixed(1)
          ];
        });
        
        if (summaryData.length > 0) {
          performanceSheet.getRange(6, 1, summaryData.length, summaryHeaders.length)
            .setValues(summaryData);
        }
        
        // 各チームの詳細セクション
        let currentRow = 6 + summaryData.length + 2;
        
        teams.forEach(teamName => {
          const team = teamPerformance[teamName];
          
          // チームセクションのヘッダー
          performanceSheet.getRange(currentRow, 1, 1, 13).merge()
            .setValue(teamName + " 詳細分析")
            .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
          
          currentRow++;
          
          // 対応区分別セクション
          performanceSheet.getRange(currentRow, 1, 1, 13).merge()
            .setValue("対応区分別パフォーマンス");
          
          currentRow++;
          
          const categoryHeaders = [
            "対応区分", "Rally数", "割合(%)", "対応者数", "生産性(Rally/h)", 
            "効果的生産性", "目標生産性", "目標との差", "質問数", "質問率(%)",
            "相談対応数", "相談対応率(%)", "状態"
          ];
          
          performanceSheet.getRange(currentRow, 1, 1, categoryHeaders.length).setValues([categoryHeaders])
            .setFontWeight('bold').setBackground('#f3f3f3');
          
          currentRow++;
          
          // 対応区分別データ
          const categoryData = [
            'ihPrimary', 'ihSecondary', 'ptPrimary', 'ptSecondary'
          ].map(category => {
            const responses = team[`${category}Responses`];
            
            // 目標値の取得
            const targetProd = settings[`${category}TargetProd`] || 10;
            const targetRate = settings[`${category}TargetRate`] || 25;
            
            // 目標との差分
            const prodGap = responses.productivity - targetProd;
            
            // 状態評価
            let status;
            if (prodGap >= 2) {
              status = "余裕あり";
            } else if (prodGap >= -1) {
              status = "適正";
            } else if (prodGap >= -3) {
              status = "やや不足";
            } else {
              status = "不足";
            }
            
            return [
              UtilsModule.getCategoryDisplayName(category),
              responses.rallyCount,
              responses.rallyRate.toFixed(1) + "%",
              responses.responders.size,
              responses.productivity.toFixed(1),
              responses.effectiveProductivity.toFixed(1),
              targetProd.toFixed(1),
              prodGap.toFixed(1),
              responses.questionCount,
              responses.questionRate.toFixed(1) + "%",
              responses.adviceCount,
              responses.adviceRate.toFixed(1) + "%",
              status
            ];
          });
          
          performanceSheet.getRange(currentRow, 1, categoryData.length, categoryHeaders.length)
            .setValues(categoryData);
          
          currentRow += categoryData.length + 2;
          
          // メンバー生産性ランキングセクション
          performanceSheet.getRange(currentRow, 1, 1, 10).merge()
            .setValue("メンバー生産性ランキング（上位10名）");
          
          currentRow++;
          
          const memberHeaders = [
            "担当者名", "IH/PT", "Rally数", "生産性(Rally/h)", "効果的生産性", 
            "質問数", "質問率(%)", "相談対応数", "相談対応率(%)", "貢献度(%)"
          ];
          
          performanceSheet.getRange(currentRow, 1, 1, memberHeaders.length).setValues([memberHeaders])
            .setFontWeight('bold').setBackground('#f3f3f3');
          
          currentRow++;
          
          // メンバー生産性ランキング
          const memberData = Object.values(team.memberPerformance)
            .filter(m => m.rallyCount > 0)
            .sort((a, b) => b.productivity - a.productivity)
            .map(m => {
              // 貢献度の計算
              const contributionRate = (m.rallyCount / team.totalRallyCount) * 100;
              
              return [
                m.name,
                m.ihPt,
                m.rallyCount,
                m.productivity.toFixed(1),
                m.effectiveProductivity.toFixed(1),
                m.questionCount,
                m.questionRate.toFixed(1) + "%",
                m.adviceCount,
                m.adviceRate.toFixed(1) + "%",
                contributionRate.toFixed(1) + "%"
              ];
            });
          
          // 上位10名のみ表示
          const topMembers = memberData.slice(0, 10);
          if (topMembers.length > 0) {
            performanceSheet.getRange(currentRow, 1, topMembers.length, memberHeaders.length)
              .setValues(topMembers);
          }
          
          currentRow += Math.max(topMembers.length, 1) + 3;
        });
        
        // シート書式設定
        this.setTeamPerformanceFormatting(performanceSheet, summaryData.length);
        
        // 列幅調整
        performanceSheet.setColumnWidth(1, 150); // OPチーム/担当者名
        performanceSheet.autoResizeColumns(2, 12); // その他の列は自動調整
        
        LogModule.logOperation("チームパフォーマンス分析", `完了: ${teams.length}チームを分析`);
        UtilsModule.showToast(`${teams.length}チームのパフォーマンス分析が完了しました！`, 'チームパフォーマンス');
        
        return teamPerformance;
      } catch (e) {
        LogModule.handleError(e, 'チームパフォーマンス分析');
        return null;
      }
    },
    
    /**
     * チームパフォーマンスシートの書式を設定する
     */
    setTeamPerformanceFormatting: function(sheet, summaryRows) {
      // チーム概要の書式設定
      if (summaryRows > 0) {
        // チーム内完結率の書式設定
        const completionRange = sheet.getRange(6, 6, summaryRows, 1);
        
        const highCompletionRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("8")
          .setBackground('#C8E6C9') // 緑
          .setRanges([completionRange])
          .build();
        
        const midCompletionRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("5")
          .setBackground('#FFF9C4') // 黄色
          .setRanges([completionRange])
          .build();
        
        const lowCompletionRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("3")
          .setBackground('#FFEBEE') // 薄い赤
          .setRanges([completionRange])
          .build();
        
        let rules = sheet.getConditionalFormatRules();
        rules.push(highCompletionRule, midCompletionRule, lowCompletionRule);
        sheet.setConditionalFormatRules(rules);
        
        // PT対応率の書式設定
        const ptRateRange = sheet.getRange(6, 8, summaryRows, 1);
        
        const highPTRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("5")
          .setBackground('#FFEBEE') // 薄い赤
          .setRanges([ptRateRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(highPTRule);
        sheet.setConditionalFormatRules(rules);
        
        // 生産性の書式設定
        const productivityRanges = [
          sheet.getRange(6, 10, summaryRows, 1),  // チーム内生産性
          sheet.getRange(6, 11, summaryRows, 1), // 効果的生産性
        ];
        
        productivityRanges.forEach(range => {
          const highProdRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberGreaterThan(10)
            .setBackground('#C8E6C9') // 緑
            .setRanges([range])
            .build();
          
          const lowProdRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberLessThan(5)
            .setBackground('#FFECB3') // 黄色
            .setRanges([range])
            .build();
          
          rules = sheet.getConditionalFormatRules();
          rules.push(highProdRule, lowProdRule);
          sheet.setConditionalFormatRules(rules);
        });
        
        // HC余剰/不足の書式設定
        const hcGapRange = sheet.getRange(6, 13, summaryRows, 1);
        
        const highHCRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(0.5)
          .setBackground('#C8E6C9') // 緑
          .setRanges([hcGapRange])
          .build();
        
        const lowHCRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(-0.5)
          .setBackground('#FFEBEE') // 薄い赤
          .setRanges([hcGapRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(highHCRule, lowHCRule);
        sheet.setConditionalFormatRules(rules);
      }
    },
  
    /**
     * 機能領域カバレッジ分析を実行する
     */
    analyzeAreaCoverage: function() {
      LogModule.logOperation("機能領域カバレッジ分析", "開始");
      
      try {
        // データ準備
        const data = DataModule.prepareAnalysisData();
        if (!data) {
          throw new Error('分析データの準備に失敗しました');
        }
        
        // 設定値取得
        const settings = SettingsModule.getSettings();
        
        // シート取得
        const coverageSheet = UtilsModule.getOrCreateSheet(SHEETS.AREA_COVERAGE);
        
        // 機能領域（OPチーム）リストの取得
        const opTeams = UtilsModule.getOPTeamList(data.sheets.dataSheet);
        
        // 担当者マスターのマッピング
        const memberMap = data.responderMasterMap;
        
        // 機能領域ごとのカバレッジデータ
        const areaCoverage = {};
        
        // 機能領域の初期化
        opTeams.forEach(team => {
          areaCoverage[team] = {
            areaName: team,
            totalTags: 0,
            highLevelTags: 0,  // L4, L5タグの数
            
            // 対応状況
            totalRallyCount: 0,
            
            // 対応区分の細分化
            ihPrimaryResponses: { rallyCount: 0, responders: new Set(), responseTime: 0, questionCount: 0, adviceCount: 0 },
            ihSecondaryResponses: { rallyCount: 0, responders: new Set(), responseTime: 0, questionCount: 0, adviceCount: 0 },
            ptPrimaryResponses: { rallyCount: 0, responders: new Set(), responseTime: 0, questionCount: 0, adviceCount: 0 },
            ptSecondaryResponses: { rallyCount: 0, responders: new Set(), responseTime: 0, questionCount: 0, adviceCount: 0 },
            
            // 時間データ
            totalResponseTime: 0,
            
            // 相談データ
            totalConsultations: 0,
            totalSolveDuration: 0,
            
            // 対応者データ
            responders: {}
          };
        });
        
        // タグの機能領域情報マッピング
        const tagAreaMap = {};
        const diffData = data.difficultyData;
        const diffColIndexes = data.diffColIndexes;
        
        for (let i = 1; i < diffData.length; i++) {
          const row = diffData[i];
          const tagName = row[diffColIndexes.tagName];
          const opTeam = row[diffColIndexes.opTeam];
          const level = row[diffColIndexes.level];
          
          if (tagName && opTeam && UtilsModule.isFunctionTag(tagName)) {
            tagAreaMap[tagName] = opTeam;
            
            // 機能領域ごとのタグ数カウント
            if (areaCoverage[opTeam]) {
              areaCoverage[opTeam].totalTags++;
              
              // 高難易度タグのカウント
              if (level === 'L4' || level === 'L5') {
                areaCoverage[opTeam].highLevelTags++;
              }
            }
          }
        }
        
        // 対応データの集計
        const txData = data.transactionData;
        const txColIndexes = data.txColIndexes;
        
        // 相談ケース総数のカウント（質問率・相談対応率の分母用）
        let totalConsultationCases = 0;
        for (let i = 1; i < txData.length; i++) {
          const row = txData[i];
          if (txColIndexes.soudanUserName !== -1 && row[txColIndexes.soudanUserName]) {
            totalConsultationCases++;
          }
        }
        
        // 対応データの集計
        for (let i = 1; i < txData.length; i++) {
          const row = txData[i];
          const tagName = row[txColIndexes.tagName];
          const userName = row[txColIndexes.userName];
          const ihPt = row[txColIndexes.ihPt];
          const opName = row[txColIndexes.opName] || tagAreaMap[tagName];
          const responseTime = row[txColIndexes.responseTime];
          const rallyCount = txColIndexes.rallyCount !== -1 ? (row[txColIndexes.rallyCount] || 1) : 1;
          
          // OJTタグの除外（オプション）
          const isOjtCase = txColIndexes.hasOjtTag !== -1 && row[txColIndexes.hasOjtTag] === "TRUE";
          if (isOjtCase && settings && settings.excludeOjtData === true) {
            continue;
          }
          
          // 選択された機能領域のタグのみ処理
          if (tagName && UtilsModule.isFunctionTag(tagName) && opName && opTeams.includes(opName)) {
            const area = opName;
            
            // 対応区分判定
            let category;
            
            // IH/PT区分による分類
            if (ihPt === 'IH') {
              // 所属チーム判定
              const memberInfo = memberMap[userName];
              const primaryTeam = memberInfo ? memberInfo.primaryTeam : "";
              const secondaryTeams = memberInfo && memberInfo.secondaryTeam ? 
                                    memberInfo.secondaryTeam.split(',').map(t => t.trim()) : [];
              
              if (primaryTeam === area) {
                category = 'ihPrimary'; // IH所属チーム
              } else if (secondaryTeams.includes(area)) {
                category = 'ihSecondary'; // IH兼務チーム
              } else {
                category = 'ihSecondary'; // IH他チーム
              }
            } else {
              // PT所属チーム判定
              const memberInfo = memberMap[userName];
              const ptPrimaryTeam = memberInfo ? memberInfo.primaryTeam : "";
              
              if (ptPrimaryTeam === area) {
                category = 'ptPrimary'; // PT所属チーム
              } else {
                category = 'ptSecondary'; // PT他チーム
              }
            }
            
            // 機能領域全体の集計
            areaCoverage[area].totalRallyCount += rallyCount;
            
            // 時間データが有効な場合のみ集計
            if (responseTime > 0 && responseTime < 1440) {
              areaCoverage[area].totalResponseTime += responseTime;
              
              // 対応区分別の集計
              if (areaCoverage[area][`${category}Responses`]) {
                areaCoverage[area][`${category}Responses`].rallyCount += rallyCount;
                areaCoverage[area][`${category}Responses`].responders.add(userName);
                areaCoverage[area][`${category}Responses`].responseTime += responseTime;
              }
              
              // 対応者別集計
              if (!areaCoverage[area].responders[userName]) {
                areaCoverage[area].responders[userName] = {
                  name: userName,
                  ihPt: ihPt,
                  primaryTeam: memberMap[userName] ? memberMap[userName].primaryTeam : "",
                  rallyCount: 0,
                  responseTime: 0,
                  tagCount: new Set(),
                  questionCount: 0,
                  adviceCount: 0
                };
              }
              
              areaCoverage[area].responders[userName].rallyCount += rallyCount;
              areaCoverage[area].responders[userName].responseTime += responseTime;
              areaCoverage[area].responders[userName].tagCount.add(tagName);
            }
          }
        }
        
        // 相談データの集計（質問と回答）
        for (let i = 1; i < txData.length; i++) {
          const row = txData[i];
          const tagName = row[txColIndexes.tagName];
          const opName = row[txColIndexes.opName] || tagAreaMap[tagName];
          
          // 質問者と回答者の情報取得
          const soudanUserName = txColIndexes.soudanUserName !== -1 ? row[txColIndexes.soudanUserName] : null;
          const adviserUserName = txColIndexes.adviserUserName !== -1 ? row[txColIndexes.adviserUserName] : null;
          const solveDuration = txColIndexes.solveDuration !== -1 ? row[txColIndexes.solveDuration] || 0 : 0;
          
          if (tagName && UtilsModule.isFunctionTag(tagName) && opName && opTeams.includes(opName)) {
            const area = opName;
            
            // 相談があった場合
            if (soudanUserName && adviserUserName) {
              areaCoverage[area].totalConsultations++;
              areaCoverage[area].totalSolveDuration += solveDuration;
              
              // 質問者の対応区分判定と集計
              let askerCategory;
              if (soudanUserName && memberMap[soudanUserName]) {
                const ihPt = memberMap[soudanUserName].ihPt || "";
                const primaryTeam = memberMap[soudanUserName].primaryTeam || "";
                const secondaryTeams = memberMap[soudanUserName].secondaryTeam ? 
                                      memberMap[soudanUserName].secondaryTeam.split(',').map(t => t.trim()) : [];
                
                if (ihPt === 'IH') {
                  if (primaryTeam === area) {
                    askerCategory = 'ihPrimary';
                  } else if (secondaryTeams.includes(area)) {
                    askerCategory = 'ihSecondary';
                  } else {
                    askerCategory = 'ihSecondary';
                  }
                } else {
                  if (primaryTeam === area) {
                    askerCategory = 'ptPrimary';
                  } else {
                    askerCategory = 'ptSecondary';
                  }
                }
                
                if (areaCoverage[area][`${askerCategory}Responses`]) {
                  areaCoverage[area][`${askerCategory}Responses`].questionCount++;
                }
                
                // 対応者別質問数の集計
                if (areaCoverage[area].responders[soudanUserName]) {
                  areaCoverage[area].responders[soudanUserName].questionCount++;
                }
              }
              
              // 回答者の対応区分判定と集計
              let adviserCategory;
              if (adviserUserName && memberMap[adviserUserName]) {
                const ihPt = memberMap[adviserUserName].ihPt || "";
                const primaryTeam = memberMap[adviserUserName].primaryTeam || "";
                const secondaryTeams = memberMap[adviserUserName].secondaryTeam ? 
                                      memberMap[adviserUserName].secondaryTeam.split(',').map(t => t.trim()) : [];
                
                if (ihPt === 'IH') {
                  if (primaryTeam === area) {
                    adviserCategory = 'ihPrimary';
                  } else if (secondaryTeams.includes(area)) {
                    adviserCategory = 'ihSecondary';
                  } else {
                    adviserCategory = 'ihSecondary';
                  }
                } else {
                  if (primaryTeam === area) {
                    adviserCategory = 'ptPrimary';
                  } else {
                    adviserCategory = 'ptSecondary';
                  }
                }
                
                if (areaCoverage[area][`${adviserCategory}Responses`]) {
                  areaCoverage[area][`${adviserCategory}Responses`].adviceCount++;
                }
                
                // 対応者別回答数の集計
                if (areaCoverage[area].responders[adviserUserName]) {
                  areaCoverage[area].responders[adviserUserName].adviceCount++;
                }
              }
            }
          }
        }
        
        // 集計結果の計算
        Object.values(areaCoverage).forEach(area => {
          // 各対応区分の生産性と割合の計算
          const categories = ['ihPrimary', 'ihSecondary', 'ptPrimary', 'ptSecondary'];
          const totalRally = area.totalRallyCount || 1; // ゼロ除算防止
          
          categories.forEach(category => {
            const responses = area[`${category}Responses`];
            
            // Rally生産性の計算（Rally/時間）
            responses.productivity = responses.responseTime > 0 ? 
                                 (responses.rallyCount / responses.responseTime) * 60 : 0;
            
            // 対応割合の計算
            responses.rallyRate = (responses.rallyCount / totalRally) * 100;
            
            // 質問率の計算
            responses.questionRate = responses.rallyCount > 0 ? 
                                (responses.questionCount / responses.rallyCount) * 100 : 0;
            
            // 相談対応率の計計算
            responses.adviceRate = area.totalConsultations > 0 ? 
                               (responses.adviceCount / area.totalConsultations) * 100 : 0;
            
            // 相談コスト係数を考慮した効果的な生産性
            responses.effectiveProductivity = UtilsModule.calculateEffectiveProductivity(
              responses.productivity, responses.questionRate, responses.adviceRate, settings
            );
          });
          
          // 対応者の生産性計算
          Object.values(area.responders).forEach(responder => {
            if (responder.responseTime > 0) {
              responder.productivity = (responder.rallyCount / responder.responseTime) * 60; // 1時間あたりのRally数
              
              // 質問率の計算
              responder.questionRate = responder.rallyCount > 0 ? 
                                 (responder.questionCount / responder.rallyCount) * 100 : 0;
              
              // 相談対応率の計算
              responder.adviceRate = area.totalConsultations > 0 ? 
                                 (responder.adviceCount / area.totalConsultations) * 100 : 0;
              
              // 相談コスト係数を考慮した効果的な生産性
              responder.effectiveProductivity = UtilsModule.calculateEffectiveProductivity(
                responder.productivity, responder.questionRate, responder.adviceRate, settings
              );
            }
          });
          
          // 全体の生産性計算
          area.totalProductivity = area.totalResponseTime > 0 ? 
                                (area.totalRallyCount / area.totalResponseTime) * 60 : 0;
          
          // チーム内完結率（所属IHの対応率）
          area.completionRate = area.ihPrimaryResponses.rallyRate;
          
          // 相談コスト係数を考慮した効果的な生産性
          area.totalQuestionRate = area.totalRallyCount > 0 ? 
                                (area.totalConsultations / area.totalRallyCount) * 100 : 0;
          
          area.effectiveProductivity = UtilsModule.calculateEffectiveProductivity(
            area.totalProductivity, area.totalQuestionRate, 0, settings
          );
          
          // 対応者数の計算
          area.ihPrimaryResponderCount = area.ihPrimaryResponses.responders.size;
          area.ihSecondaryResponderCount = area.ihSecondaryResponses.responders.size;
          area.ptPrimaryResponderCount = area.ptPrimaryResponses.responders.size;
          area.ptSecondaryResponderCount = area.ptSecondaryResponses.responders.size;
          
          // メンバーカバレッジ（所属IHメンバー数/タグ数）
          area.memberCoverage = area.totalTags > 0 ? 
                               area.ihPrimaryResponderCount / area.totalTags * 100 : 0;
          
          // 月間必要HC計算
          // 月間稼働時間 = 稼働日数 × 1日あたり時間 × 生産性係数
          const monthlyWorkHours = settings.workDaysPerMonth * settings.hoursPerDay * settings.productivityRatio;
          
          // 月間Rally数推定 (3ヶ月分のデータから1ヶ月分を推定)
          const monthlyRallyCount = area.totalRallyCount / 3;
          
          // チーム内完結目標率
          const targetCompletionRate = settings.targetCompletionRate || 85;
          
          // 目標生産性
          const targetRallyPerHour = settings.targetRallyPerHour || 10;
          
          // チーム内完結分の必要HC
          area.idealHC = (monthlyRallyCount * (targetCompletionRate / 100)) / (targetRallyPerHour * monthlyWorkHours);
          
          // 相談コスト考慮済みの必要HC (オプション)
          area.adjustedIdealHC = area.idealHC * (1 + (settings.consultationCostFactor || 0) * (area.totalQuestionRate / 100));
        });
        
        // 機能領域カバレッジシートの作成
        coverageSheet.clear();
        
        // ヘッダーとタイトル
        coverageSheet.getRange("A1:M1").merge().setValue("機能領域カバレッジ分析")
          .setFontWeight('bold').setBackground('#E8EAF6').setFontSize(14);
        
        coverageSheet.getRange(2, 1).setValue("更新日: " + new Date().toLocaleDateString());
        
        // 領域概要セクション
        coverageSheet.getRange("A4:M4").merge().setValue("機能領域概要")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
        
        const summaryHeaders = [
          "機能領域", "対応タグ数", "高難度タグ数", "Rally総数", 
          "所属IH対応率(%)", "IH他チーム対応率(%)", "PT所属対応率(%)", "PT他チーム対応率(%)",
          "所属IHメンバー数", "他チームIHメンバー数", "PT所属メンバー数", "PT他チームメンバー数",
          "メンバーカバレッジ(%)"
        ];
        
        coverageSheet.getRange(5, 1, 1, summaryHeaders.length).setValues([summaryHeaders])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // 領域概要データ
        const summaryData = opTeams.map(teamName => {
          const area = areaCoverage[teamName];
          return [
            area.areaName,
            area.totalTags,
            area.highLevelTags,
            area.totalRallyCount,
            area.ihPrimaryResponses.rallyRate.toFixed(1) + "%",
            area.ihSecondaryResponses.rallyRate.toFixed(1) + "%",
            area.ptPrimaryResponses.rallyRate.toFixed(1) + "%",
            area.ptSecondaryResponses.rallyRate.toFixed(1) + "%",
            area.ihPrimaryResponderCount,
            area.ihSecondaryResponderCount,
            area.ptPrimaryResponderCount,
            area.ptSecondaryResponderCount,
            area.memberCoverage.toFixed(1) + "%"
          ];
        });
        
        if (summaryData.length > 0) {
          coverageSheet.getRange(6, 1, summaryData.length, summaryHeaders.length)
            .setValues(summaryData);
        }
        
        // 各領域の詳細セクション
        let currentRow = 6 + summaryData.length + 2;
        
        opTeams.forEach(teamName => {
          const area = areaCoverage[teamName];
          
          // 領域セクションのヘッダー
          coverageSheet.getRange(currentRow, 1, 1, 13).merge()
            .setValue(teamName + " 詳細分析")
            .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
          
          currentRow++;
          
          // 対応区分別セクション
          coverageSheet.getRange(currentRow, 1, 1, 13).merge()
            .setValue("対応区分別パフォーマンス");
          
          currentRow++;
          
          const categoryHeaders = [
            "対応区分", "Rally数", "割合(%)", "対応者数", "生産性(Rally/h)", 
            "効果的生産性", "目標生産性", "目標との差", "質問数", "質問率(%)",
            "相談対応数", "相談対応率(%)", "状態"
          ];
          
          coverageSheet.getRange(currentRow, 1, 1, categoryHeaders.length).setValues([categoryHeaders])
            .setFontWeight('bold').setBackground('#f3f3f3');
          
          currentRow++;
          
          // 対応区分別データ
          const categoryData = [
            'ihPrimary', 'ihSecondary', 'ptPrimary', 'ptSecondary'
          ].map(category => {
            const responses = area[`${category}Responses`];
            
            // 目標値の取得
            const targetProd = settings[`${category}TargetProd`] || 10;
            const targetRate = settings[`${category}TargetRate`] || 25;
            
            // 目標との差分
            const prodGap = responses.productivity - targetProd;
            
            // 状態評価
            let status;
            if (prodGap >= 2) {
              status = "余裕あり";
            } else if (prodGap >= -1) {
              status = "適正";
            } else if (prodGap >= -3) {
              status = "やや不足";
            } else {
              status = "不足";
            }
            
            return [
              UtilsModule.getCategoryDisplayName(category),
              responses.rallyCount,
              responses.rallyRate.toFixed(1) + "%",
              responses.responders.size,
              responses.productivity.toFixed(1),
              responses.effectiveProductivity.toFixed(1),
              targetProd.toFixed(1),
              prodGap.toFixed(1),
              responses.questionCount,
              responses.questionRate.toFixed(1) + "%",
              responses.adviceCount,
              responses.adviceRate.toFixed(1) + "%",
              status
            ];
          });
          
          coverageSheet.getRange(currentRow, 1, categoryData.length, categoryHeaders.length)
            .setValues(categoryData);
          
          currentRow += categoryData.length + 2;
          
          // 対応者詳細ランキング
          coverageSheet.getRange(currentRow, 1, 1, 9).merge()
            .setValue("対応者詳細 (上位15名)");
          
          currentRow++;
          
          const responderHeaders = [
            "担当者名", "IH/PT", "所属チーム", "Rally数", "生産性(Rally/h)", 
            "効果的生産性", "対応タグ数", "質問率(%)", "相談対応率(%)"
          ];
          
          coverageSheet.getRange(currentRow, 1, 1, responderHeaders.length).setValues([responderHeaders])
            .setFontWeight('bold').setBackground('#f3f3f3');
          
          currentRow++;
          
          // 対応者ランキング
          const responderData = Object.values(area.responders)
            .filter(r => r.rallyCount > 0)
            .sort((a, b) => b.rallyCount - a.rallyCount)
            .map(r => [
              r.name,
              r.ihPt,
              r.primaryTeam,
              r.rallyCount,
              r.productivity.toFixed(1),
              r.effectiveProductivity.toFixed(1),
              r.tagCount.size,
              r.questionRate.toFixed(1) + "%",
              r.adviceRate.toFixed(1) + "%"
            ]);
          
          // 上位15名のみ表示
          const topResponders = responderData.slice(0, 15);
          if (topResponders.length > 0) {
            coverageSheet.getRange(currentRow, 1, topResponders.length, responderHeaders.length)
              .setValues(topResponders);
          }
          
          currentRow += Math.max(topResponders.length, 1) + 3;
        });
        
        // シート書式設定
        this.setAreaCoverageFormatting(coverageSheet, summaryData.length);
        
        // 列幅調整
        coverageSheet.setColumnWidth(1, 150); // 機能領域/担当者名
        coverageSheet.setColumnWidth(3, 100); // 所属チーム
        coverageSheet.autoResizeColumns(2, 11); // その他の列は自動調整
        
        LogModule.logOperation("機能領域カバレッジ分析", `完了: ${opTeams.length}領域を分析`);
        UtilsModule.showToast(`${opTeams.length}機能領域のカバレッジ分析が完了しました！`, '機能領域カバレッジ');
        
        return areaCoverage;
      } catch (e) {
        LogModule.handleError(e, '機能領域カバレッジ分析');
        return null;
      }
    },
    
    /**
     * 機能領域カバレッジシートの書式を設定する
     */
    setAreaCoverageFormatting: function(sheet, summaryRows) {
      // 領域概要の書式設定
      if (summaryRows > 0) {
        // 所属IH対応率の書式設定
        const ihPrimaryRange = sheet.getRange(6, 5, summaryRows, 1);
        
        const highPrimaryRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("8")
          .setBackground('#C8E6C9') // 緑
          .setRanges([ihPrimaryRange])
          .build();
        
        const midPrimaryRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("5")
          .setBackground('#FFF9C4') // 黄色
          .setRanges([ihPrimaryRange])
          .build();
        
        const lowPrimaryRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("3")
          .setBackground('#FFEBEE') // 薄い赤
          .setRanges([ihPrimaryRange])
          .build();
        
        let rules = sheet.getConditionalFormatRules();
        rules.push(highPrimaryRule, midPrimaryRule, lowPrimaryRule);
        sheet.setConditionalFormatRules(rules);
        
        // PT対応率の書式設定
        const ptRateRange = sheet.getRange(6, 7, summaryRows, 1);
        
        const highPTRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("5")
          .setBackground('#FFEBEE') // 薄い赤
          .setRanges([ptRateRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(highPTRule);
        sheet.setConditionalFormatRules(rules);
        
        // メンバーカバレッジの書式設定
        const coverageRange = sheet.getRange(6, 13, summaryRows, 1);
        
        const highCoverageRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("8")
          .setBackground('#C8E6C9') // 緑
          .setRanges([coverageRange])
          .build();
        
        const lowCoverageRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("3")
          .setBackground('#FFEBEE') // 薄い赤
          .setRanges([coverageRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(highCoverageRule, lowCoverageRule);
        sheet.setConditionalFormatRules(rules);
      }
    },
    
    /**
     * 相談フロー分析を実行する
     */
    analyzeConsultationFlow: function() {
      LogModule.logOperation("相談フロー分析", "開始");
      
      try {
        // データ準備
        const data = DataModule.prepareAnalysisData();
        if (!data) {
          throw new Error('分析データの準備に失敗しました');
        }
        
        // 設定値取得
        const settings = SettingsModule.getSettings();
        
        // シート取得
        const consultationFlowSheet = UtilsModule.getOrCreateSheet(SHEETS.CONSULTATION_FLOW);
        
        // チーム一覧の取得
        const teams = UtilsModule.getTeamList(data.sheets.masterSheet);
        
        // 担当者マスターのマッピング
        const memberMap = data.responderMasterMap;
        
        // チームごとの相談フローデータ
        const teamConsultationFlows = {};
        
        // チームの初期化
        teams.forEach(team => {
          teamConsultationFlows[team] = {
            teamName: team,
            totalConsultations: 0,
            totalSolveDuration: 0,
            
            // カテゴリ別相談集計
            categoryCounts: {
              ihPrimary: { asked: 0, advised: 0 },
              ihSecondary: { asked: 0, advised: 0 },
              ptPrimary: { asked: 0, advised: 0 },
              ptSecondary: { asked: 0, advised: 0 }
            },
            
            // 相談フロー（カテゴリ間フロー）
            categoryFlows: {},
            
            // 機能タグごとの相談集計
            tagConsultations: {},
            
            // 個人間の相談集計
            personalFlows: {}
          };
          
          // カテゴリフロー初期化
          const categories = ['ihPrimary', 'ihSecondary', 'ptPrimary', 'ptSecondary'];
          categories.forEach(askerCategory => {
            categories.forEach(adviserCategory => {
              const flowKey = `${askerCategory}To${adviserCategory}`;
              teamConsultationFlows[team].categoryFlows[flowKey] = {
                count: 0,
                rallyCount: 0,
                responseTime: 0,
                solveDuration: 0
              };
            });
          });
        });
        
        // 対応データの集計
        const txData = data.transactionData;
        const txColIndexes = data.txColIndexes;
        
        // タグの所属チーム情報マッピング
        const tagTeamMap = {};
        const diffData = data.difficultyData;
        const diffColIndexes = data.diffColIndexes;
        
        for (let i = 1; i < diffData.length; i++) {
          const tagName = diffData[i][diffColIndexes.tagName];
          const opTeam = diffData[i][diffColIndexes.opTeam];
          if (tagName && opTeam) {
            tagTeamMap[tagName] = opTeam;
          }
        }
        
        // 相談データの集計
        for (let i = 1; i < txData.length; i++) {
          const row = txData[i];
          const tagName = row[txColIndexes.tagName];
          const opName = row[txColIndexes.opName] || tagTeamMap[tagName];
          
          // OJTタグの除外（オプション）
          const isOjtCase = txColIndexes.hasOjtTag !== -1 && row[txColIndexes.hasOjtTag] === "TRUE";
          if (isOjtCase && settings && settings.excludeOjtData === true) {
            continue;
          }
          
          // 相談に関するデータを抽出
          const soudanUserName = txColIndexes.soudanUserName !== -1 ? row[txColIndexes.soudanUserName] : null;
          const adviserUserName = txColIndexes.adviserUserName !== -1 ? row[txColIndexes.adviserUserName] : null;
          const solveDuration = txColIndexes.solveDuration !== -1 ? row[txColIndexes.solveDuration] || 0 : 0;
          const responseTime = txColIndexes.responseTime !== -1 ? row[txColIndexes.responseTime] || 0 : 0;
          const rallyCount = txColIndexes.rallyCount !== -1 ? row[txColIndexes.rallyCount] || 1 : 1;
          
          // 有効な相談データとチームが特定できる場合のみ処理
          if (soudanUserName && adviserUserName && opName && teams.includes(opName)) {
            const team = opName;
            teamConsultationFlows[team].totalConsultations++;
            teamConsultationFlows[team].totalSolveDuration += solveDuration;
            
            // カテゴリ判定
            let askerCategory = 'unknown';
            let adviserCategory = 'unknown';
            
            // 質問者のカテゴリ判定
            if (soudanUserName && memberMap[soudanUserName]) {
              const ihPt = memberMap[soudanUserName].ihPt || "";
              const primaryTeam = memberMap[soudanUserName].primaryTeam || "";
              const secondaryTeams = memberMap[soudanUserName].secondaryTeam ? 
                                    memberMap[soudanUserName].secondaryTeam.split(',').map(t => t.trim()) : [];
              
              if (ihPt === 'IH') {
                if (primaryTeam === team) {
                  askerCategory = 'ihPrimary';
                } else if (secondaryTeams.includes(team)) {
                  askerCategory = 'ihSecondary';
                } else {
                  askerCategory = 'ihSecondary';
                }
              } else {
                if (primaryTeam === team) {
                  askerCategory = 'ptPrimary';
                } else {
                  askerCategory = 'ptSecondary';
                }
              }
              
              // カテゴリ別質問数カウント
              teamConsultationFlows[team].categoryCounts[askerCategory].asked++;
            }
            
            // 回答者のカテゴリ判定
            if (adviserUserName && memberMap[adviserUserName]) {
              const ihPt = memberMap[adviserUserName].ihPt || "";
              const primaryTeam = memberMap[adviserUserName].primaryTeam || "";
              const secondaryTeams = memberMap[adviserUserName].secondaryTeam ? 
                                    memberMap[adviserUserName].secondaryTeam.split(',').map(t => t.trim()) : [];
              
              if (ihPt === 'IH') {
                if (primaryTeam === team) {
                  adviserCategory = 'ihPrimary';
                } else if (secondaryTeams.includes(team)) {
                  adviserCategory = 'ihSecondary';
                } else {
                  adviserCategory = 'ihSecondary';
                }
              } else {
                if (primaryTeam === team) {
                  adviserCategory = 'ptPrimary';
                } else {
                  adviserCategory = 'ptSecondary';
                }
              }
              
              // カテゴリ別回答数カウント
              teamConsultationFlows[team].categoryCounts[adviserCategory].advised++;
            }
            
            // カテゴリ間フロー集計
            if (askerCategory !== 'unknown' && adviserCategory !== 'unknown') {
              const flowKey = `${askerCategory}To${adviserCategory}`;
              
              if (teamConsultationFlows[team].categoryFlows[flowKey]) {
                teamConsultationFlows[team].categoryFlows[flowKey].count++;
                teamConsultationFlows[team].categoryFlows[flowKey].rallyCount += rallyCount;
                teamConsultationFlows[team].categoryFlows[flowKey].responseTime += responseTime;
                teamConsultationFlows[team].categoryFlows[flowKey].solveDuration += solveDuration;
              }
            }
            
            // 機能タグごとの相談集計
            if (tagName) {
              if (!teamConsultationFlows[team].tagConsultations[tagName]) {
                teamConsultationFlows[team].tagConsultations[tagName] = {
                  tagName: tagName,
                  count: 0,
                  rallyCount: 0,
                  solveDuration: 0,
                  askerCategories: {},
                  adviserCategories: {}
                };
              }
              
              teamConsultationFlows[team].tagConsultations[tagName].count++;
              teamConsultationFlows[team].tagConsultations[tagName].rallyCount += rallyCount;
              teamConsultationFlows[team].tagConsultations[tagName].solveDuration += solveDuration;
              
              // 質問者カテゴリ集計
              if (askerCategory !== 'unknown') {
                if (!teamConsultationFlows[team].tagConsultations[tagName].askerCategories[askerCategory]) {
                  teamConsultationFlows[team].tagConsultations[tagName].askerCategories[askerCategory] = 0;
                }
                teamConsultationFlows[team].tagConsultations[tagName].askerCategories[askerCategory]++;
              }
              
              // 回答者カテゴリ集計
              if (adviserCategory !== 'unknown') {
                if (!teamConsultationFlows[team].tagConsultations[tagName].adviserCategories[adviserCategory]) {
                  teamConsultationFlows[team].tagConsultations[tagName].adviserCategories[adviserCategory] = 0;
                }
                teamConsultationFlows[team].tagConsultations[tagName].adviserCategories[adviserCategory]++;
              }
            }
            
            // 個人間の相談集計
            const personalFlowKey = `${soudanUserName}To${adviserUserName}`;
            
            if (!teamConsultationFlows[team].personalFlows[personalFlowKey]) {
              teamConsultationFlows[team].personalFlows[personalFlowKey] = {
                asker: soudanUserName,
                adviser: adviserUserName,
                askerCategory: askerCategory,
                adviserCategory: adviserCategory,
                count: 0,
                rallyCount: 0,
                solveDuration: 0,
                tags: new Set()
              };
            }
            
            teamConsultationFlows[team].personalFlows[personalFlowKey].count++;
            teamConsultationFlows[team].personalFlows[personalFlowKey].rallyCount += rallyCount;
            teamConsultationFlows[team].personalFlows[personalFlowKey].solveDuration += solveDuration;
            
            if (tagName) {
              teamConsultationFlows[team].personalFlows[personalFlowKey].tags.add(tagName);
            }
          }
        }
        
        // 相談フロー分析シートの作成
        consultationFlowSheet.clear();
        
        // ヘッダーとタイトル
        consultationFlowSheet.getRange("A1:K1").merge().setValue("相談フロー分析")
          .setFontWeight('bold').setBackground('#E8EAF6').setFontSize(14);
        
        consultationFlowSheet.getRange(2, 1).setValue("更新日: " + new Date().toLocaleDateString());
        
        // チーム概要セクション
        consultationFlowSheet.getRange("A4:K4").merge().setValue("チーム別相談概要")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
        
        const summaryHeaders = [
          "OPチーム", "相談総数", "平均解決時間(分)", "IH所属発信率(%)", 
          "IH所属受信率(%)", "IH他チーム発信率(%)", "IH他チーム受信率(%)",
          "PT所属発信率(%)", "PT所属受信率(%)", "PT他チーム発信率(%)", "PT他チーム受信率(%)"
        ];
        
        consultationFlowSheet.getRange(5, 1, 1, summaryHeaders.length).setValues([summaryHeaders])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // チーム概要データ
        const summaryData = teams.map(teamName => {
          const team = teamConsultationFlows[teamName];
          const totalCount = team.totalConsultations || 1; // ゼロ除算防止
          const avgSolveDuration = team.totalConsultations > 0 ? 
                                team.totalSolveDuration / team.totalConsultations : 0;
          
          // カテゴリごとの発信率と受信率を計算
          const askedRates = {};
          const advisedRates = {};
          
          Object.entries(team.categoryCounts).forEach(([category, counts]) => {
            askedRates[category] = (counts.asked / totalCount) * 100;
            advisedRates[category] = (counts.advised / totalCount) * 100;
          });
          
          return [
            team.teamName,
            team.totalConsultations,
            avgSolveDuration.toFixed(1),
            askedRates.ihPrimary.toFixed(1) + "%",
            advisedRates.ihPrimary.toFixed(1) + "%",
            askedRates.ihSecondary.toFixed(1) + "%",
            advisedRates.ihSecondary.toFixed(1) + "%",
            askedRates.ptPrimary.toFixed(1) + "%",
            advisedRates.ptPrimary.toFixed(1) + "%",
            askedRates.ptSecondary.toFixed(1) + "%",
            advisedRates.ptSecondary.toFixed(1) + "%"
          ];
        });
        
        if (summaryData.length > 0) {
          consultationFlowSheet.getRange(6, 1, summaryData.length, summaryHeaders.length)
            .setValues(summaryData);
        }
        
        // 各チームの詳細セクション
        let currentRow = 6 + summaryData.length + 2;
        
        teams.forEach(teamName => {
          const team = teamConsultationFlows[teamName];
          
          // チームセクションのヘッダー
          consultationFlowSheet.getRange(currentRow, 1, 1, 11).merge()
            .setValue(teamName + " 相談フロー詳細")
            .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
          
          currentRow++;
          
          // カテゴリ間フローセクション
          consultationFlowSheet.getRange(currentRow, 1, 1, 9).merge()
            .setValue("カテゴリ間相談フロー");
          
          currentRow++;
          
          const flowHeaders = [
            "相談元", "相談先", "件数", "割合(%)", "平均解決時間(分)", 
            "Rally総数", "分あたりRally数", "ナレッジ化優先度", "アクション"
          ];
          
          consultationFlowSheet.getRange(currentRow, 1, 1, flowHeaders.length).setValues([flowHeaders])
            .setFontWeight('bold').setBackground('#f3f3f3');
          
          currentRow++;
          
          // カテゴリ間フローデータの作成
          const flowData = Object.entries(team.categoryFlows)
            .filter(([_, flow]) => flow.count > 0)
            .map(([flowKey, flow]) => {
              const [askerCategory, adviserCategory] = flowKey.split('To');
              const percentage = (flow.count / team.totalConsultations) * 100;
              const avgSolveDuration = flow.count > 0 ? flow.solveDuration / flow.count : 0;
              const rallyPerMinute = flow.solveDuration > 0 ? flow.rallyCount / flow.solveDuration : 0;
              
              // ナレッジ化優先度の計算
              let priority = 0;
              if (percentage > 15) priority += 3;
              else if (percentage > 5) priority += 1;
              
              if (avgSolveDuration > 15) priority += 2;
              else if (avgSolveDuration > 5) priority += 1;
              
              if (flow.count > 20) priority += 3;
              else if (flow.count > 5) priority += 1;
              
              let priorityText = "";
              if (priority >= 6) priorityText = "高";
              else if (priority >= 3) priorityText = "中";
              else if (priority > 0) priorityText = "低";
              
              // アクション推奨
              let action = "";
              if (priority >= 6) {
                action = "ナレッジ作成";
              } else if (priority >= 3) {
                if (askerCategory.includes('Secondary') || askerCategory.includes('pt')) {
                  action = "スキル獲得検討";
                } else {
                  action = "研修検討";
                }
              }
              
              return [
                UtilsModule.getCategoryDisplayName(askerCategory),
                UtilsModule.getCategoryDisplayName(adviserCategory),
                flow.count,
                percentage.toFixed(1) + "%",
                avgSolveDuration.toFixed(1),
                flow.rallyCount,
                rallyPerMinute.toFixed(2),
                priorityText,
                action
              ];
            });
          
          // 件数の多い順にソート
          flowData.sort((a, b) => b[2] - a[2]);
          
          // 上位16件までを表示
          const topFlows = flowData.slice(0, 16);
          if (topFlows.length > 0) {
            consultationFlowSheet.getRange(currentRow, 1, topFlows.length, flowHeaders.length)
              .setValues(topFlows);
            
            currentRow += topFlows.length + 2;
          } else {
            consultationFlowSheet.getRange(currentRow, 1).setValue("相談フローデータがありません");
            currentRow += 2;
          }
          
          // 機能タグ別相談セクション
          consultationFlowSheet.getRange(currentRow, 1, 1, 7).merge()
            .setValue("機能タグ別相談状況");
          
          currentRow++;
          
          const tagHeaders = [
            "タグ名", "相談件数", "割合(%)", "平均解決時間(分)", 
            "主な質問者", "主な回答者", "対応改善アクション"
          ];
          
          consultationFlowSheet.getRange(currentRow, 1, 1, tagHeaders.length).setValues([tagHeaders])
            .setFontWeight('bold').setBackground('#f3f3f3');
          
          currentRow++;
          
          // 機能タグ別相談データの作成
          const tagData = Object.values(team.tagConsultations)
            .filter(tag => tag.count > 0)
            .map(tag => {
              const percentage = (tag.count / team.totalConsultations) * 100;
              const avgSolveDuration = tag.count > 0 ? tag.solveDuration / tag.count : 0;
              
              // 主な質問者カテゴリの特定
              let mainAskerCategory = "不明";
              let maxAskerCount = 0;
              
              Object.entries(tag.askerCategories).forEach(([category, count]) => {
                if (count > maxAskerCount) {
                  maxAskerCount = count;
                  mainAskerCategory = category;
                }
              });
              
              // 主な回答者カテゴリの特定
              let mainAdviserCategory = "不明";
              let maxAdviserCount = 0;
              
              Object.entries(tag.adviserCategories).forEach(([category, count]) => {
                if (count > maxAdviserCount) {
                  maxAdviserCount = count;
                  mainAdviserCategory = category;
                }
              });
              
              // 対応改善アクション
              let action = "";
              if (percentage > 5 && avgSolveDuration > 10) {
                if (mainAskerCategory.includes('Secondary') || mainAskerCategory.includes('pt')) {
                  action = "スキル移管検討";
                } else {
                  action = "ナレッジ作成";
                }
              } else if (tag.count > 10) {
                action = "FAQ検討";
              }
              
              return [
                tag.tagName,
                tag.count,
                percentage.toFixed(1) + "%",
                avgSolveDuration.toFixed(1),
                UtilsModule.getCategoryDisplayName(mainAskerCategory),
                UtilsModule.getCategoryDisplayName(mainAdviserCategory),
                action
              ];
            });
          
          // 件数の多い順にソート
          tagData.sort((a, b) => b[1] - a[1]);
          
          // 上位10件までを表示
          const topTags = tagData.slice(0, 10);
          if (topTags.length > 0) {
            consultationFlowSheet.getRange(currentRow, 1, topTags.length, tagHeaders.length)
              .setValues(topTags);
            
            currentRow += topTags.length + 2;
          } else {
            consultationFlowSheet.getRange(currentRow, 1).setValue("タグ別相談データがありません");
            currentRow += 2;
          }
          
          // 個人間相談フローセクション
          consultationFlowSheet.getRange(currentRow, 1, 1, 9).merge()
            .setValue("個人間相談フロー（上位10件）");
          
          currentRow++;
          
          const personalHeaders = [
            "質問者", "質問者区分", "回答者", "回答者区分", 
            "件数", "割合(%)", "平均解決時間(分)", "対応タグ数", "ペア生産性"
          ];
          
          consultationFlowSheet.getRange(currentRow, 1, 1, personalHeaders.length).setValues([personalHeaders])
            .setFontWeight('bold').setBackground('#f3f3f3');
          
          currentRow++;
          
          // 個人間相談フローデータの作成
          const personalData = Object.values(team.personalFlows)
            .filter(flow => flow.count > 0)
            .map(flow => {
              const percentage = (flow.count / team.totalConsultations) * 100;
              const avgSolveDuration = flow.count > 0 ? flow.solveDuration / flow.count : 0;
              const pairProductivity = flow.solveDuration > 0 ? (flow.rallyCount / flow.solveDuration) * 60 : 0;
              
              return [
                flow.asker,
                UtilsModule.getCategoryDisplayName(flow.askerCategory),
                flow.adviser,
                UtilsModule.getCategoryDisplayName(flow.adviserCategory),
                flow.count,
                percentage.toFixed(1) + "%",
                avgSolveDuration.toFixed(1),
                flow.tags.size,
                pairProductivity.toFixed(1)
              ];
            });
          
          // 件数の多い順にソート
          personalData.sort((a, b) => b[4] - a[4]);
          
          // 上位10件までを表示
          const topPersonalFlows = personalData.slice(0, 10);
          if (topPersonalFlows.length > 0) {
            consultationFlowSheet.getRange(currentRow, 1, topPersonalFlows.length, personalHeaders.length)
              .setValues(topPersonalFlows);
            
            currentRow += topPersonalFlows.length + 3;
          } else {
            consultationFlowSheet.getRange(currentRow, 1).setValue("個人間相談データがありません");
            currentRow += 3;
          }
        });
        
        // シート書式設定
        this.setConsultationFlowFormatting(consultationFlowSheet, summaryData.length);
        
        // 列幅調整
        consultationFlowSheet.setColumnWidth(1, 150); // 相談元/タグ名/質問者
        consultationFlowSheet.setColumnWidth(2, 150); // 相談先/質問者区分
        consultationFlowSheet.setColumnWidth(3, 100); // 件数
        consultationFlowSheet.setColumnWidth(9, 150); // アクション
        consultationFlowSheet.autoResizeColumns(4, 5); // その他の列は自動調整
        
        LogModule.logOperation("相談フロー分析", `完了: ${teams.length}チームの相談フローを分析`);
        UtilsModule.showToast(`${teams.length}チームの相談フロー分析が完了しました！`, '相談フロー分析');
        
        return teamConsultationFlows;
      } catch (e) {
        LogModule.handleError(e, '相談フロー分析');
        return null;
      }
    },
  
    /**
     * 相談フロー分析シートの書式を設定する
     */
    setConsultationFlowFormatting: function(sheet, summaryRows) {
      // チーム概要の書式設定
      if (summaryRows > 0) {
        // 相談件数の書式設定
        const consultationCountRange = sheet.getRange(6, 2, summaryRows, 1);
        
        const highCountRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(100)
          .setBackground('#C8E6C9') // 緑
          .setRanges([consultationCountRange])
          .build();
        
        const lowCountRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(10)
          .setBackground('#FFF9C4') // 黄色
          .setRanges([consultationCountRange])
          .build();
        
        let rules = sheet.getConditionalFormatRules();
        rules.push(highCountRule, lowCountRule);
        sheet.setConditionalFormatRules(rules);
        
        // 発信・受信率の書式設定
        const rateRanges = [
          {range: sheet.getRange(6, 4, summaryRows, 1), target: 'ihPrimary'}, // IH所属発信率
          {range: sheet.getRange(6, 5, summaryRows, 1), target: 'ihPrimary'}, // IH所属受信率
          {range: sheet.getRange(6, 6, summaryRows, 1), target: 'ihSecondary'}, // IH他チーム発信率
          {range: sheet.getRange(6, 7, summaryRows, 1), target: 'ihSecondary'}, // IH他チーム受信率
          {range: sheet.getRange(6, 8, summaryRows, 1), target: 'ptPrimary'}, // PT所属発信率
          {range: sheet.getRange(6, 9, summaryRows, 1), target: 'ptPrimary'}, // PT所属受信率
          {range: sheet.getRange(6, 10, summaryRows, 1), target: 'ptSecondary'}, // PT他チーム発信率
          {range: sheet.getRange(6, 11, summaryRows, 1), target: 'ptSecondary'} // PT他チーム受信率
        ];
        
        rateRanges.forEach(item => {
          const highRateRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextContains("50")
            .setBackground('#E3F2FD') // 薄い青
            .setRanges([item.range])
            .build();
          
          rules = sheet.getConditionalFormatRules();
          rules.push(highRateRule);
          sheet.setConditionalFormatRules(rules);
        });
      }
    }
  };