// ユーティリティモジュール
const UtilsModule = {
  /**
   * 機能タグかどうかを判定する
   */
  isFunctionTag: function(tagName) {
    if (!tagName) return false;
    const tag = tagName.toString();
    return tag.startsWith('機能_') || tag.startsWith('koban_');
  },
  
  /**
   * シートを取得する（存在しない場合はnullを返す）
   */
  getSheet: function(sheetName) {
    return SS.getSheetByName(sheetName);
  },

  /**
   * シートを取得または作成する
   */
  getOrCreateSheet: function(sheetName) {
    let sheet = SS.getSheetByName(sheetName);
    if (!sheet) {
      sheet = SS.insertSheet(sheetName);
    }
    return sheet;
  },

  /**
   * トースト通知を表示する
   */
  showToast: function(message, title) {
    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
    } catch (e) {
      console.log('トースト表示エラー: ' + e.toString());
    }
  },

  /**
   * アラートを表示する
   */
  showAlert: function(message) {
    try {
      SpreadsheetApp.getUi().alert(message);
    } catch (e) {
      console.log('アラート表示エラー: ' + e.toString());
    }
  },
  
  /**
   * パーセンタイル値を計算する
   */
  calculatePercentile: function(array, percentile) {
    if (array.length === 0) return 0;
    
    const sortedArray = [...array].sort((a, b) => a - b);
    const index = (percentile / 100) * sortedArray.length;
    
    if (Math.floor(index) === index) {
      return (sortedArray[index - 1] + sortedArray[index]) / 2;
    }
    return sortedArray[Math.floor(index)];
  },
  
  /**
   * 各次元のレベルを判定する
   */
  getDimensionLevel: function(score, dimensionPrefix, settings) {
    const l1Threshold = settings[dimensionPrefix + "L1Threshold"] || 1.8;
    const l2Threshold = settings[dimensionPrefix + "L2Threshold"] || 2.6;
    const l3Threshold = settings[dimensionPrefix + "L3Threshold"] || 3.4;
    const l4Threshold = settings[dimensionPrefix + "L4Threshold"] || 4.2;
    
    if (score < l1Threshold) return "L1";
    if (score < l2Threshold) return "L2";
    if (score < l3Threshold) return "L3";
    if (score < l4Threshold) return "L4";
    return "L5";
  },
  
  /**
   * 対応時間スコアを計算する (1-5)
   */
  calculateTimeScore: function(medianTime) {
    if (medianTime < 5) return 1;       // 5分未満
    if (medianTime < 15) return 2;      // 5-15分
    if (medianTime < 30) return 3;      // 15-30分
    if (medianTime < 60) return 4;      // 30-60分
    return 5;                           // 60分以上
  },
  
  /**
   * カバレッジスコアを計算する (1-5)
   */
  calculateCoverageScore: function(coveragePercent) {
    if (coveragePercent > 80) return 1; // 80%以上
    if (coveragePercent > 60) return 2; // 60-80%
    if (coveragePercent > 40) return 3; // 40-60%
    if (coveragePercent > 20) return 4; // 20-40%
    return 5;                           // 20%未満
  },
  
  /**
   * Rallyスコアを計算する (1-5)
   */
  calculateRallyScore: function(avgRallyCount) {
    if (avgRallyCount < 1.5) return 1;  // 平均1.5未満
    if (avgRallyCount < 2.5) return 2;  // 平均1.5-2.5
    if (avgRallyCount < 3.5) return 3;  // 平均2.5-3.5
    if (avgRallyCount < 5) return 4;    // 平均3.5-5
    return 5;                           // 平均5以上
  },
  
  /**
   * シート内のレベル表示に条件付き書式を適用する
   */
  applyLevelFormatting: function(sheet, range, levelColumn) {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    
    const levelRange = sheet.getRange(2, levelColumn, lastRow - 1, 1);
    
    // L1-L5の色分け
    const blueColors = ['#E3F2FD', '#90CAF9', '#42A5F5', '#1976D2', '#0D47A1'];
    
    for (let i = 1; i <= 5; i++) {
      const level = 'L' + i;
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(level)
        .setBackground(blueColors[i - 1])
        .setFontColor(i >= 4 ? 'white' : 'black')
        .setRanges([levelRange])
        .build();
      
      const rules = sheet.getConditionalFormatRules();
      rules.push(rule);
      sheet.setConditionalFormatRules(rules);
    }
  },
  
  /**
   * マスターデータからチーム一覧を取得する
   */
  getTeamList: function(masterSheet) {
    try {
      const data = masterSheet.getDataRange().getValues();
      const headers = data[0];
      
      const primaryTeamIdx = headers.indexOf('主所属OPチーム');
      const secondaryTeamIdx = headers.indexOf('兼務OPチーム');
      
      if (primaryTeamIdx === -1) {
        return [];
      }
      
      const teams = new Set();
      
      for (let i = 1; i < data.length; i++) {
        const primaryTeam = data[i][primaryTeamIdx];
        if (primaryTeam) teams.add(primaryTeam);
        
        // 兼務チームも追加
        if (secondaryTeamIdx !== -1) {
          const secondaryTeams = data[i][secondaryTeamIdx];
          if (secondaryTeams) {
            secondaryTeams.split(',').forEach(team => {
              const trimmedTeam = team.trim();
              if (trimmedTeam) teams.add(trimmedTeam);
            });
          }
        }
      }
      
      return Array.from(teams).sort();
      
    } catch (e) {
      LogModule.handleError(e, 'チーム一覧取得');
      return [];
    }
  },
  
  /**
   * 取込データからOPチーム（機能領域）一覧を取得する
   */
  getOPTeamList: function(dataSheet) {
    try {
      const data = dataSheet.getDataRange().getValues();
      const headers = data[0];
      
      const opTeamIdx = headers.indexOf('op_name');
      
      if (opTeamIdx === -1) {
        return [];
      }
      
      const opTeams = new Set();
      
      for (let i = 1; i < data.length; i++) {
        const opTeam = data[i][opTeamIdx];
        if (opTeam) opTeams.add(opTeam);
      }
      
      return Array.from(opTeams).sort();
      
    } catch (e) {
      LogModule.handleError(e, 'OPチーム一覧取得');
      return [];
    }
  },
  
  /**
   * PTメンバーの所属チームを判定する
   */
  isPtPrimaryTeam: function(userName, teamName, memberMap) {
    const memberInfo = memberMap[userName];
    if (!memberInfo || memberInfo.ihPt !== 'PT') return false;
    
    // PTの主所属チーム判定
    const ptPrimaryTeam = memberInfo.primaryTeam || "";
    return ptPrimaryTeam === teamName;
  },

  /**
   * 対応区分カテゴリを判定する
   * @param {string} userName ユーザー名
   * @param {string} teamName チーム名
   * @param {object} memberMap メンバーマップ
   * @return {string} 対応区分カテゴリ ('ihPrimary', 'ihSecondary', 'ptPrimary', 'ptSecondary', 'unknown')
   */
  getResponseCategory: function(userName, teamName, memberMap) {
    if (!memberMap || !memberMap[userName]) return "unknown";
    
    const ihPt = memberMap[userName].ihPt || "";
    const primaryTeam = memberMap[userName].primaryTeam || "";
    
    if (ihPt === 'IH') {
      return primaryTeam === teamName ? 'ihPrimary' : 'ihSecondary';
    } else {
      return this.isPtPrimaryTeam(userName, teamName, memberMap) ? 'ptPrimary' : 'ptSecondary';
    }
  },

  /**
   * 対応区分の表示名を取得する
   * @param {string} category 対応区分カテゴリ ('ihPrimary', 'ihSecondary', 'ptPrimary', 'ptSecondary')
   * @return {string} 表示名
   */
  getCategoryDisplayName: function(category) {
    const displayNames = {
      'ihPrimary': 'IH所属チーム',
      'ihSecondary': 'IH他チーム',
      'ptPrimary': 'PT所属チーム',
      'ptSecondary': 'PT他チーム',
      'unknown': '不明'
    };
    
    return displayNames[category] || category;
  },

  /**
   * Rallyベースの生産性を計算する
   * @param {number} rallyCount Rally数
   * @param {number} responseTime 対応時間
   * @return {number} 時間あたりのRally数
   */
  calculateRallyProductivity: function(rallyCount, responseTime) {
    return responseTime > 0 ? (rallyCount / responseTime) * 60 : 0;
  },

  /**
   * 相談コスト係数を適用した生産性を計算する
   * @param {number} baseProductivity 基本生産性
   * @param {number} questionRate 質問率(%)
   * @param {number} adviceRate 相談対応率(%)
   * @param {object} settings 設定値
   * @return {number} 相談コスト考慮済み生産性
   */
  calculateEffectiveProductivity: function(baseProductivity, questionRate, adviceRate, settings) {
    const questionCostFactor = settings.questionCostFactor || 0.3;
    const consultationCostFactor = settings.consultationCostFactor || 0.5;
    
    // 相談コスト係数の計算
    const costFactor = 1 + (questionRate / 100) * questionCostFactor + (adviceRate / 100) * consultationCostFactor;
    
    // 効果的な生産性の計算
    return baseProductivity / costFactor;
  }
};