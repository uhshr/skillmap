// ダッシュボードモジュール
const DashboardModule = {
    /**
     * 個人スキルダッシュボードを初期化する
     */
    initPersonalDashboard: function() {
      LogModule.logOperation("個人スキルダッシュボード初期化", "開始");
      
      try {
        // 必要なシートの取得
        const skillSheet = UtilsModule.getSheet(SHEETS.INDIVIDUAL_SKILL);
        const masterSheet = UtilsModule.getSheet(SHEETS.MASTER);
        const personalDashSheet = UtilsModule.getOrCreateSheet(SHEETS.PERSONAL_DASHBOARD);
        
        // データの確認
        if (!skillSheet || !masterSheet) {
          throw new Error('必要なデータシートが見つかりません。先に分析を実行してください。');
        }
        
        // シートを初期化
        personalDashSheet.clear();
        
        // ヘッダーとタイトル
        personalDashSheet.getRange("A1:G1").merge().setValue("個人スキルダッシュボード")
          .setFontWeight('bold').setBackground('#E8EAF6').setFontSize(14);
        
        personalDashSheet.getRange("A2:G2").merge().setValue("更新日: " + new Date().toLocaleDateString())
          .setFontStyle("italic");
        
        // 操作方法の説明
        personalDashSheet.getRange("A3:G3").merge()
          .setValue("使用方法: 担当者を選択 → 「スキルマップ→ダッシュボード→個人スキルダッシュボード更新」を実行")
          .setBackground('#FFF9C4').setFontWeight('bold');
        
        // 担当者選択セクション
        personalDashSheet.getRange("A5").setValue("担当者:");
        
        // 個人スキルシートから担当者一覧を取得
        const skillData = skillSheet.getDataRange().getValues();
        const skillHeaders = skillData[0];
        const nameIdx = skillHeaders.indexOf('担当者名');
        
        const responders = [];
        for (let i = 1; i < skillData.length; i++) {
          if (skillData[i][nameIdx]) {
            responders.push(skillData[i][nameIdx]);
          }
        }
        
        // ドロップダウンリストの設定
        if (responders.length > 0) {
          const responderValidation = SpreadsheetApp.newDataValidation()
            .requireValueInList(responders, true)
            .build();
          
          personalDashSheet.getRange("B5").setDataValidation(responderValidation);
          
          // デフォルト値として最初の担当者を設定
          if (responders[0]) {
            personalDashSheet.getRange("B5").setValue(responders[0]);
          }
        } else {
          personalDashSheet.getRange("B5").setValue("(担当者が見つかりません)");
        }
        
        // 担当者選択の説明文
        personalDashSheet.getRange("C5").setValue("担当者を選択して「個人スキルダッシュボード更新」を実行してください");
        
        // 基本情報セクション（表示用のプレースホルダー）
        personalDashSheet.getRange("A7:G7").merge().setValue("基本情報")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
        
        const basicInfoHeaders = ["項目", "値", "", "項目", "値"];
        personalDashSheet.getRange(8, 1, 1, 5).setValues([basicInfoHeaders])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        const basicInfoRows = [
          ["担当者名", "", "", "パフォーマータイプ", ""],
          ["IH/PT区分", "", "", "幅スコア", ""],
          ["主所属OPチーム", "", "", "深さスコア", ""],
          ["兼務OPチーム", "", "", "生産性スコア", ""],
          ["対応タグ数", "", "", "総合スキルスコア", ""],
          ["チーム一致率", "", "", "相談対応率", ""],
          ["マクロ使用率", "", "", "質問率", ""]
        ];
        
        personalDashSheet.getRange(9, 1, basicInfoRows.length, 5).setValues(basicInfoRows);
        
        // 機能別スキル詳細セクション
        personalDashSheet.getRange("A18:G18").merge().setValue("機能別スキル詳細")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
        
        // 所属チーム内タグセクション
        personalDashSheet.getRange("A20:G20").merge().setValue("所属チーム内タグ")
          .setFontWeight('bold').setBackground('#E8F5E9');
        
        const tagHeaders = ["タグ名", "対応件数", "難易度", "対応レベル", "複雑案件対応", "マクロ使用率", "生産性"];
        personalDashSheet.getRange(21, 1, 1, 7).setValues([tagHeaders])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // 所属チーム外タグセクション
        personalDashSheet.getRange("A30:G30").merge().setValue("所属チーム外タグ")
          .setFontWeight('bold').setBackground('#E3F2FD');
        
        personalDashSheet.getRange(31, 1, 1, 7).setValues([tagHeaders])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // 列幅調整
        personalDashSheet.setColumnWidth(1, 150);
        personalDashSheet.setColumnWidth(2, 100);
        personalDashSheet.setColumnWidth(3, 80);
        personalDashSheet.setColumnWidth(4, 150);
        personalDashSheet.setColumnWidth(5, 120);
        personalDashSheet.setColumnWidth(6, 120);
        personalDashSheet.setColumnWidth(7, 100);
        
        // シートをアクティブにする
        personalDashSheet.activate();
        
        LogModule.logOperation("個人スキルダッシュボード初期化", "完了");
        UtilsModule.showToast('個人スキルダッシュボードを初期化しました', 'ダッシュボード初期化');
        
        return true;
      } catch (e) {
        LogModule.handleError(e, '個人スキルダッシュボード初期化');
        return false;
      }
    },
  
    /**
     * 個人スキルダッシュボードを更新する
     */
    updatePersonalDashboard: function() {
      LogModule.logOperation("個人スキルダッシュボード更新", "開始");
      
      try {
        // 必要なシートの取得
        const skillSheet = UtilsModule.getSheet(SHEETS.INDIVIDUAL_SKILL);
        const masterSheet = UtilsModule.getSheet(SHEETS.MASTER);
        const difficultySheet = UtilsModule.getSheet(SHEETS.TAG_DIFFICULTY);
        const distributionSheet = UtilsModule.getSheet(SHEETS.TAG_DISTRIBUTION);
        const dataSheet = UtilsModule.getSheet(SHEETS.DATA);
        const personalDashSheet = UtilsModule.getSheet(SHEETS.PERSONAL_DASHBOARD);
        
        if (!skillSheet || !masterSheet || !personalDashSheet || !dataSheet || !difficultySheet) {
          throw new Error('必要なデータシートが見つかりません。先に初期化を実行してください。');
        }
        
        // 担当者名の取得
        const responderName = personalDashSheet.getRange("B5").getValue();
        
        if (!responderName || responderName === "(担当者が見つかりません)") {
          personalDashSheet.getRange("C5").setValue("担当者を選択してください");
          return;
        }
        
        // 個人スキルデータの取得
        const skillData = skillSheet.getDataRange().getValues();
        const skillHeaders = skillData[0];
        
        // 難易度データの取得
        const difficultyData = difficultySheet.getDataRange().getValues();
        const difficultyHeaders = difficultyData[0];
        
        // マスターデータの取得
        const masterData = masterSheet.getDataRange().getValues();
        const masterHeaders = masterData[0];
        
        // 取込データの取得
        const transactionData = dataSheet.getDataRange().getValues();
        const transactionHeaders = transactionData[0];
        
        // ヘッダーインデックスの取得
        const nameIdx = skillHeaders.indexOf('担当者名');
        const ihPtIdx = skillHeaders.indexOf('IH/PT区分');
        const primaryTeamIdx = skillHeaders.indexOf('主所属OPチーム');
        const secondaryTeamIdx = skillHeaders.indexOf('兼務OPチーム');
        const tagCountIdx = skillHeaders.indexOf('対応タグ数');
        const widthScoreIdx = skillHeaders.indexOf('幅スコア');
        const depthScoreIdx = skillHeaders.indexOf('深さスコア');
        const productivityScoreIdx = skillHeaders.indexOf('生産性スコア');
        const teamMatchRateIdx = skillHeaders.indexOf('チーム一致率(%)');
        const macroRateIdx = skillHeaders.indexOf('マクロ使用率(%)');
        const adviserRateIdx = skillHeaders.indexOf('相談対応率(%)');
        const questionRateIdx = skillHeaders.indexOf('質問率(%)');
        const totalSkillScoreIdx = skillHeaders.indexOf('総合スキルスコア');
        const performerTypeIdx = skillHeaders.indexOf('パフォーマータイプ');
        
        // 取込データのインデックス
        const txNameIdx = transactionHeaders.indexOf('user_name');
        const txTagIdx = transactionHeaders.indexOf('tag_name');
        const txOpTeamIdx = transactionHeaders.indexOf('op_name');
        const txRallyIdx = transactionHeaders.indexOf('rally_count');
        const txResponseTimeIdx = transactionHeaders.indexOf('total_response_time');
        const txMacroIdx = transactionHeaders.indexOf('use_macro');
        
        // 難易度データのインデックス
        const diffTagIdx = difficultyHeaders.indexOf('タグ名');
        const diffOpTeamIdx = difficultyHeaders.indexOf('OPチーム名');
        const diffLevelIdx = difficultyHeaders.indexOf('難易度レベル');
        
        // 担当者データの検索
        let responderData = null;
        for (let i = 1; i < skillData.length; i++) {
          if (skillData[i][nameIdx] === responderName) {
            responderData = skillData[i];
            break;
          }
        }
        
        if (!responderData) {
          throw new Error('選択された担当者のデータが見つかりません');
        }
        
        // 基本情報の更新
        personalDashSheet.getRange("B9").setValue(responderName);
        personalDashSheet.getRange("B10").setValue(responderData[ihPtIdx]);
        personalDashSheet.getRange("B11").setValue(responderData[primaryTeamIdx]);
        personalDashSheet.getRange("B12").setValue(responderData[secondaryTeamIdx]);
        personalDashSheet.getRange("B13").setValue(responderData[tagCountIdx]);
        personalDashSheet.getRange("B14").setValue(responderData[teamMatchRateIdx] + "%");
        personalDashSheet.getRange("B15").setValue(responderData[macroRateIdx] + "%");
        
        personalDashSheet.getRange("E9").setValue(responderData[performerTypeIdx]);
        personalDashSheet.getRange("E10").setValue(responderData[widthScoreIdx]);
        personalDashSheet.getRange("E11").setValue(responderData[depthScoreIdx]);
        personalDashSheet.getRange("E12").setValue(responderData[productivityScoreIdx]);
        personalDashSheet.getRange("E13").setValue(responderData[totalSkillScoreIdx]);
        personalDashSheet.getRange("E14").setValue(responderData[adviserRateIdx] + "%");
        personalDashSheet.getRange("E15").setValue(responderData[questionRateIdx] + "%");
        
        // パフォーマータイプに応じた色設定
        const performerType = responderData[performerTypeIdx];
        let typeColor;
        
        switch (performerType) {
          case 'ハイパフォーマー': typeColor = '#1565C0'; break; // 濃い青
          case 'スペシャリスト': typeColor = '#6A1B9A'; break;   // 紫
          case 'オールラウンダー': typeColor = '#2E7D32'; break; // 緑
          default: typeColor = '#757575'; break;              // グレー
        }
        
        personalDashSheet.getRange("E9").setBackground(typeColor).setFontColor('white');
        
        // 説明文の更新
        personalDashSheet.getRange("C5").setValue(responderName + "さんのスキル分析");
        
        // 機能別スキル詳細の取得と表示
        const primaryTeam = responderData[primaryTeamIdx];
        
        // 担当者のタグごとの対応状況を集計
        const tagPerformance = {};
        let totalTeamInCount = 0;
        let totalTeamOutCount = 0;
        
        for (let i = 1; i < transactionData.length; i++) {
          const row = transactionData[i];
          const name = row[txNameIdx];
          const tag = row[txTagIdx];
          const opTeam = row[txOpTeamIdx];
          const rally = row[txRallyIdx] || 1;
          const responseTime = row[txResponseTimeIdx] || 0;
          const useMacro = row[txMacroIdx] === "マクロ使用";
          
          // 担当者が対応したケースのみ
          if (name === responderName && tag && UtilsModule.isFunctionTag(tag)) {
            // タグ情報の初期化
            if (!tagPerformance[tag]) {
              const isTeamIn = (opTeam === primaryTeam);
              
              // 難易度データから情報取得
              let level = "L3"; // デフォルト
              
              for (let j = 1; j < difficultyData.length; j++) {
                if (difficultyData[j][diffTagIdx] === tag) {
                  level = difficultyData[j][diffLevelIdx];
                  break;
                }
              }
              
              tagPerformance[tag] = {
                tag: tag,
                opTeam: opTeam,
                isTeamIn: isTeamIn,
                level: level,
                count: 0,
                rallyCount: 0,
                responseTime: 0,
                macroUsageCount: 0,
                hasComplexCase: false,
                productivity: 0
              };
              
              // カウント
              if (isTeamIn) {
                totalTeamInCount++;
              } else {
                totalTeamOutCount++;
              }
            }
            
            // データ集計
            tagPerformance[tag].count++;
            tagPerformance[tag].rallyCount += rally;
            
            if (responseTime > 0 && responseTime < 1440) {
              tagPerformance[tag].responseTime += responseTime;
              
              // 複雑ケース判定（対応時間が長いものを複雑ケースとみなす）
              if (responseTime >= 60) { // 60分以上は複雑ケースとみなす
                tagPerformance[tag].hasComplexCase = true;
              }
            }
            
            if (useMacro) {
              tagPerformance[tag].macroUsageCount++;
            }
          }
        }
        
        // タグごとの生産性計算
        Object.values(tagPerformance).forEach(tp => {
          // 生産性計算（Rally/時間）
          tp.productivity = tp.responseTime > 0 ? (tp.rallyCount / tp.responseTime) * 60 : 0;
          
          // マクロ使用率
          tp.macroUsageRate = tp.count > 0 ? (tp.macroUsageCount / tp.count) * 100 : 0;
          
          // 対応レベル判定
          if (tp.hasComplexCase) {
            tp.responseLevel = "複雑〜";
          } else if (tp.count >= 5) {
            tp.responseLevel = "標準〜";
          } else {
            tp.responseLevel = "簡易〜";
          }
        });
        
        // 所属チーム内タグと所属チーム外タグに分けてソート
        const teamInTags = Object.values(tagPerformance)
          .filter(tp => tp.isTeamIn)
          .sort((a, b) => b.count - a.count);
        
        const teamOutTags = Object.values(tagPerformance)
          .filter(tp => !tp.isTeamIn)
          .sort((a, b) => b.count - a.count);
        
        // まず全てのセルをクリア
        personalDashSheet.getRange(22, 1, 8, 7).clearContent();
        
        // 所属チーム内タグの表示
        if (teamInTags.length > 0) {
          const teamInData = teamInTags.map(tp => [
            tp.tag,
            tp.count,
            tp.level,
            tp.responseLevel,
            tp.hasComplexCase ? "○" : "",
            tp.macroUsageRate.toFixed(1) + "%",
            tp.productivity.toFixed(1)
          ]);
          
          // 最大8件まで表示
          const displayCount = Math.min(teamInTags.length, 8);
          personalDashSheet.getRange(22, 1, displayCount, 7).setValues(teamInData.slice(0, displayCount));
        } else {
          // メッセージを表示（セル結合しない）
          personalDashSheet.getRange(22, 1).setValue("所属チーム内のタグ対応がありません");
          personalDashSheet.getRange(22, 1).setFontStyle("italic");
        }
        
        // 所属チーム外タグの表示（同様にクリアしてから処理）
        personalDashSheet.getRange(32, 1, 8, 7).clearContent();
        
        if (teamOutTags.length > 0) {
          const teamOutData = teamOutTags.map(tp => [
            tp.tag,
            tp.count,
            tp.level,
            tp.responseLevel,
            tp.hasComplexCase ? "○" : "",
            tp.macroUsageRate.toFixed(1) + "%",
            tp.productivity.toFixed(1)
          ]);
          
          // 最大8件まで表示
          const displayCount = Math.min(teamOutTags.length, 8);
          personalDashSheet.getRange(32, 1, displayCount, 7).setValues(teamOutData.slice(0, displayCount));
        } else {
          // メッセージを表示（セル結合しない）
          personalDashSheet.getRange(32, 1).setValue("所属チーム外のタグ対応がありません");
          personalDashSheet.getRange(32, 1).setFontStyle("italic");
        }
        
        // 機能別スキル詳細の書式設定
        this.setPersonalDashboardFormatting(personalDashSheet, teamInTags.length, teamOutTags.length);
        
        LogModule.logOperation("個人スキルダッシュボード更新", `完了: ${responderName}さんのデータを表示`);
        UtilsModule.showToast(`${responderName}さんのスキルデータを表示しました`, '個人スキルダッシュボード');
        
        return true;
      } catch (e) {
        LogModule.handleError(e, '個人スキルダッシュボード更新');
        return false;
      }
    },
    
    /**
     * 個人スキルダッシュボードの書式を設定する
     */
    setPersonalDashboardFormatting: function(sheet, teamInCount, teamOutCount) {
      // パフォーマータイプには既に色が設定されているため省略
      
      // 難易度レベルの書式設定
      if (teamInCount > 0) {
        const teamInLevelRange = sheet.getRange(22, 3, Math.min(teamInCount, 8), 1);
        
        const levelColors = [
          { level: "L1", color: "#E3F2FD", textColor: "black" },
          { level: "L2", color: "#90CAF9", textColor: "black" },
          { level: "L3", color: "#42A5F5", textColor: "white" },
          { level: "L4", color: "#1976D2", textColor: "white" },
          { level: "L5", color: "#0D47A1", textColor: "white" }
        ];
        
        levelColors.forEach(levelInfo => {
          const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(levelInfo.level)
            .setBackground(levelInfo.color)
            .setFontColor(levelInfo.textColor)
            .setRanges([teamInLevelRange])
            .build();
          
          let rules = sheet.getConditionalFormatRules();
          rules.push(rule);
          sheet.setConditionalFormatRules(rules);
        });
        
        // 複雑案件対応の書式設定
        const complexCaseRange = sheet.getRange(22, 5, Math.min(teamInCount, 8), 1);
        
        const complexRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("○")
          .setBackground("#E8F5E9") // 薄い緑
          .setRanges([complexCaseRange])
          .build();
        
        let rules = sheet.getConditionalFormatRules();
        rules.push(complexRule);
        sheet.setConditionalFormatRules(rules);
      }
      
      // 所属チーム外タグにも同様の書式設定
      if (teamOutCount > 0) {
        const teamOutLevelRange = sheet.getRange(32, 3, Math.min(teamOutCount, 8), 1);
        
        const levelColors = [
          { level: "L1", color: "#E3F2FD", textColor: "black" },
          { level: "L2", color: "#90CAF9", textColor: "black" },
          { level: "L3", color: "#42A5F5", textColor: "white" },
          { level: "L4", color: "#1976D2", textColor: "white" },
          { level: "L5", color: "#0D47A1", textColor: "white" }
        ];
        
        levelColors.forEach(levelInfo => {
          const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(levelInfo.level)
            .setBackground(levelInfo.color)
            .setFontColor(levelInfo.textColor)
            .setRanges([teamOutLevelRange])
            .build();
          
          let rules = sheet.getConditionalFormatRules();
          rules.push(rule);
          sheet.setConditionalFormatRules(rules);
        });
        
        // 複雑案件対応の書式設定
        const complexCaseRange = sheet.getRange(32, 5, Math.min(teamOutCount, 8), 1);
        
        const complexRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("○")
          .setBackground("#E8F5E9") // 薄い緑
          .setRanges([complexCaseRange])
          .build();
        
        let rules = sheet.getConditionalFormatRules();
        rules.push(complexRule);
        sheet.setConditionalFormatRules(rules);
      }
    },
  
      /**
     * チームパフォーマンスダッシュボードを初期化する
     */
    initTeamPerformanceDashboard: function() {
      LogModule.logOperation("チームパフォーマンスダッシュボード初期化", "開始");
      
      try {
        // 必要なシートの取得
        const masterSheet = UtilsModule.getSheet(SHEETS.MASTER);
        const skillSheet = UtilsModule.getSheet(SHEETS.INDIVIDUAL_SKILL);
        const dataSheet = UtilsModule.getSheet(SHEETS.DATA);
        const performanceSheet = UtilsModule.getOrCreateSheet(SHEETS.TEAM_PERFORMANCE_DASHBOARD);
        
        // データの確認
        if (!masterSheet || !skillSheet || !dataSheet) {
          throw new Error('必要なデータシートが見つかりません。先に分析を実行してください。');
        }
        
        // チームパフォーマンスシートをクリア
        performanceSheet.clear();
        
        // タイトルと説明の設定
        performanceSheet.getRange("A1:H1").merge().setValue("チームパフォーマンスダッシュボード")
          .setFontWeight('bold').setBackground('#E8EAF6').setFontSize(16);
        
        const today = new Date();
        performanceSheet.getRange("A2:H2").merge().setValue(`更新日: ${today.toLocaleDateString()}`)
          .setFontStyle("italic");
        
        // 操作方法の説明
        performanceSheet.getRange("A3:H3").merge().setValue("使用方法: チームを選択 → 「スキルマップ→ダッシュボード→チームパフォーマンス更新」を実行")
          .setBackground('#FFF9C4').setFontWeight('bold');
        
        // チーム選択セクション
        performanceSheet.getRange("A5").setValue("チーム選択:");
        
        // マスターデータからチーム一覧を取得
        const teams = UtilsModule.getTeamList(masterSheet);
        
        // ドロップダウンリストの設定
        if (teams.length > 0) {
          const teamValidation = SpreadsheetApp.newDataValidation()
            .requireValueInList(teams, true)
            .build();
          performanceSheet.getRange("B5").setDataValidation(teamValidation);
          
          // デフォルト値として最初のチームを設定
          if (teams[0]) {
            performanceSheet.getRange("B5").setValue(teams[0]);
          }
        } else {
          performanceSheet.getRange("B5").setValue("(チームが見つかりません)");
        }
        
        // チーム概要セクション
        performanceSheet.getRange("A7:H7").merge().setValue("チーム概要")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        // チーム概要ヘッダー
        performanceSheet.getRange("A8:E8").setValues([["チーム名", "所属メンバー数", "チーム内完結率(%)", "実質HC合計", "理想HCとの差分"]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // チーム概要データ領域の確保
        performanceSheet.getRange("A9:E9").setBackground('#f9f9f9');
        
        // チームメンバーパフォーマンスセクション
        performanceSheet.getRange("A11:H11").merge().setValue("チームメンバーパフォーマンス")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        // メンバーパフォーマンスヘッダー
        performanceSheet.getRange("A12:H12").setValues([[
          "担当者名", "IH/PT", "等級", "HC相当", "チーム機能生産性", "全機能生産性", "チーム貢献度", "チーム機能タイプ"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // メンバーパフォーマンスデータ領域の確保（最大10名分）
        performanceSheet.getRange("A13:H22").setBackground('#f9f9f9');
        
        // 外部リソース依存状況セクション
        performanceSheet.getRange("A24:E24").merge().setValue("IH外部リソース依存状況")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        // 外部リソースヘッダー
        performanceSheet.getRange("A25:E25").setValues([[
          "担当者名", "所属チーム", "対応Rally数", "依存度", "生産性"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // 外部リソースデータ領域の確保（最大5名分）
        performanceSheet.getRange("A26:E30").setBackground('#f9f9f9');
        
        // 理想HC設定セクション
        performanceSheet.getRange("A32:F32").merge().setValue("理想HC設定")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        // 理想HC設定ヘッダー
        performanceSheet.getRange("A33:F33").setValues([[
          "パラメータ", "現在値", "目標値", "差分", "調整", "説明"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // パラメータ行
        performanceSheet.getRange("A34:F38").setValues([
          ["Rally/時間 (標準)", "", "10.0", "", "↑↓", "チームの平均生産性の目標値"],
          ["チーム内完結目標 (%)", "", "85.0", "", "↑↓", "チーム内で対応できる案件の目標割合"],
          ["難易度調整係数", "", "1.0", "", "↑↓", "担当機能の難易度による生産性調整"],
          ["必要HC（Rally数ベース）", "", "", "", "【計算値】", "上記パラメータから算出される必要HC"],
          ["実際のHC - 必要HC", "", "", "", "【計算値】", "実際のHC合計と必要HCのギャップ"]
        ]);
        
        // 調整ボタン領域の設定
        // 調整ボタンの代わりに↑↓記号を配置し、メニューからの更新を促す説明を追加
        performanceSheet.getRange("A40:F40").merge().setValue("※ 目標値の調整後は「スキルマップ→ダッシュボード→チームパフォーマンス更新」を実行してください")
          .setBackground('#FFF9C4').setFontStyle("italic");
        
        // 列幅の調整
        performanceSheet.setColumnWidth(1, 150); // 担当者名
        performanceSheet.setColumnWidth(5, 150); // チーム機能生産性
        performanceSheet.setColumnWidth(6, 150); // 全機能生産性
        performanceSheet.setColumnWidth(7, 120); // チーム貢献度
        performanceSheet.setColumnWidth(8, 120); // スキルレベル
        
        // 説明列
        performanceSheet.setColumnWidth(6, 250); // 説明
        
        // チーム設定の保存領域（非表示）
        // 各チームの設定値をシートの下部に保存
        performanceSheet.getRange("A100:E100").setValues([["チームID", "チーム名", "Rally/時間目標", "完結率目標", "難易度係数"]]);
        
        // シートをアクティブにする
        performanceSheet.activate();
        
        LogModule.logOperation("チームパフォーマンス初期化", "完了");
        UtilsModule.showToast('チームパフォーマンスダッシュボードを初期化しました', 'ダッシュボード初期化');
        
        return true;
      } catch (e) {
        LogModule.handleError(e, 'チームパフォーマンスダッシュボード初期化');
        return false;
      }
    },
  
    /**
     * チームパフォーマンスを更新する
     */
    updateTeamPerformance: function() {
      LogModule.logOperation("チームパフォーマンス更新", "開始");
      
      try {
        // 設定値取得
        const settings = SettingsModule.getSettings();
  
        // 必要なシートの取得
        const masterSheet = UtilsModule.getSheet(SHEETS.MASTER);
        const skillSheet = UtilsModule.getSheet(SHEETS.INDIVIDUAL_SKILL);
        const dataSheet = UtilsModule.getSheet(SHEETS.DATA);
        const difficultySheet = UtilsModule.getSheet(SHEETS.TAG_DIFFICULTY); // タグ難易度シートを追加
        const performanceSheet = UtilsModule.getSheet(SHEETS.TEAM_PERFORMANCE_DASHBOARD);
        
        if (!masterSheet || !skillSheet || !dataSheet || !performanceSheet || !difficultySheet) {
          throw new Error('必要なデータシートが見つかりません');
        }
        
        // 選択されたチームを取得
        const selectedTeam = performanceSheet.getRange("B5").getValue();
        
        if (!selectedTeam || selectedTeam === "(チームが見つかりません)") {
          performanceSheet.getRange("A9:E9").setValues([["チームが選択されていません", "", "", "", ""]]);
          
          // メンバーデータと依存データをクリア
          performanceSheet.getRange("A13:H22").clearContent();
          performanceSheet.getRange("A26:E30").clearContent();
          
          return false;
        }
        
        // マスターデータの取得
        const masterData = masterSheet.getDataRange().getValues();
        const masterHeaders = masterData[0];
        
        const masterNameIdx = masterHeaders.indexOf('担当者名');
        const masterIhPtIdx = masterHeaders.indexOf('IH/PT区分');
        const masterRoleIdx = masterHeaders.indexOf('等級');
        const masterPrimaryTeamIdx = masterHeaders.indexOf('主所属OPチーム');
        const masterSecondaryTeamIdx = masterHeaders.indexOf('兼務OPチーム');
        const masterRatioIdx = masterHeaders.indexOf('稼働時間係数');
        
        // スキルデータの取得
        const skillData = skillSheet.getDataRange().getValues();
        const skillHeaders = skillData[0];
        
        const skillNameIdx = skillHeaders.indexOf('担当者名');
        const skillIhPtIdx = skillHeaders.indexOf('IH/PT区分');
        const skillPrimaryTeamIdx = skillHeaders.indexOf('主所属OPチーム');
        const skillProductivityIdx = skillHeaders.indexOf('生産性スコア');
        const skillTotalScoreIdx = skillHeaders.indexOf('総合スキルスコア');
        const skillPerformerTypeIdx = skillHeaders.indexOf('パフォーマータイプ');
        
        // 難易度データの取得
        const difficultyData = difficultySheet.getDataRange().getValues();
        const difficultyHeaders = difficultyData[0];
        
        const diffTagIdx = difficultyHeaders.indexOf('タグ名');
        const diffOpTeamIdx = difficultyHeaders.indexOf('OPチーム名');
        const diffLevelIdx = difficultyHeaders.indexOf('難易度レベル');
        
        // 取込データの取得
        const transactionData = dataSheet.getDataRange().getValues();
        const transactionHeaders = transactionData[0];
        
        const txNameIdx = transactionHeaders.indexOf('user_name');
        const txIhPtIdx = transactionHeaders.indexOf('ih_pt');
        const txTagIdx = transactionHeaders.indexOf('tag_name');
        const txOpTeamIdx = transactionHeaders.indexOf('op_name');
        const txRallyIdx = transactionHeaders.indexOf('rally_count');
        const txResponseTimeIdx = transactionHeaders.indexOf('total_response_time');
        
  
        // チームが担当する機能タグを収集
        const teamFunctions = new Set();
        for (let i = 1; i < difficultyData.length; i++) {
          const tag = difficultyData[i][diffTagIdx];
          const team = difficultyData[i][diffOpTeamIdx];
          
          if (team === selectedTeam && tag && UtilsModule.isFunctionTag(tag)) {
            teamFunctions.add(tag);
          }
        }
        
        // 1. チームメンバーの取得
        const teamMembers = [];
        const teamMemberMap = {};
        
        for (let i = 1; i < masterData.length; i++) {
          const name = masterData[i][masterNameIdx];
          const primaryTeam = masterData[i][masterPrimaryTeamIdx];
          const secondaryTeams = masterData[i][masterSecondaryTeamIdx] || "";
          
          // 所属チームまたは兼務チームが選択チームと一致
          if (primaryTeam === selectedTeam || 
              secondaryTeams.split(',').some(team => team.trim() === selectedTeam)) {
            
            const ihPt = masterData[i][masterIhPtIdx] || "";
            const role = masterData[i][masterRoleIdx] || "";
            let hcRatio = masterData[i][masterRatioIdx] || 1.0;
            
            // 兼務の場合はHC相当値を調整
            if (primaryTeam !== selectedTeam) {
              // 兼務の場合、担当率は低めに設定
              hcRatio = hcRatio * 0.5;
            }
            
            // スキルデータから追加情報を取得
            let teamProductivity = 0;
            let totalProductivity = 0;
            let skillScore = 0;
            let performerType = "標準";
            
            for (let j = 1; j < skillData.length; j++) {
              if (skillData[j][skillNameIdx] === name) {
                totalProductivity = parseFloat(skillData[j][skillProductivityIdx]) || 0;
                skillScore = parseFloat(skillData[j][skillTotalScoreIdx]) || 0;
                performerType = skillData[j][skillPerformerTypeIdx] || "標準";
                break;
              }
            }
            
            // チーム固有機能に対する生産性を計算
            // チーム機能に対する対応データを抽出
            let teamRallyCount = 0;
            let teamResponseTime = 0;
            
            // 対応できるチーム機能タグを収集
            const memberTeamFunctions = new Set();
            let complexCaseCount = 0;
            
            for (let j = 1; j < transactionData.length; j++) {
              const row = transactionData[j];
              const rowName = row[txNameIdx];
              const rowIhPt = row[txIhPtIdx];
              const rowTeam = row[txOpTeamIdx];
              const rowTag = row[txTagIdx];
              const rowRally = row[txRallyIdx] || 1;
              const rowResponseTime = row[txResponseTimeIdx] || 0;
             
              // 機能タグかつチーム担当かつ担当者名が一致
              if (rowName === name && 
                  rowTag && UtilsModule.isFunctionTag(rowTag) && 
                  rowTeam === selectedTeam && 
                  rowResponseTime > 0 && rowResponseTime < 1440) {
                
                teamRallyCount += rowRally;
                teamResponseTime += rowResponseTime;
                memberTeamFunctions.add(rowTag);
                
                // 複雑ケース判定（対応時間が長いものを複雑ケースとみなす）
                if (rowResponseTime >= 60) { // 60分以上は複雑ケースとみなす
                  complexCaseCount++;
                }
              }
            }
            
            // チーム固有の生産性計算（Rally/時間）
            teamProductivity = teamResponseTime > 0 ? (teamRallyCount / teamResponseTime) * 60 : 0;
            
  
            // チーム貢献度の計算
            let totalResponses = 0;
            let teamResponses = 0;
  
            for (let j = 1; j < transactionData.length; j++) {
              const row = transactionData[j];
              if (row[txNameIdx] === name && 
                  row[txTagIdx] && UtilsModule.isFunctionTag(row[txTagIdx])) {
                
                totalResponses++;
                if (row[txOpTeamIdx] === selectedTeam) {
                  teamResponses++;
                }
              }
            }
  
            const contributionRate = totalResponses > 0 ? (teamResponses / totalResponses) * 100 : 0;
            
            // チーム機能カバー率の計算
            const teamFunctionCoverage = teamFunctions.size > 0 ? 
                                      (memberTeamFunctions.size / teamFunctions.size) * 100 : 0;
            
            // チーム機能パフォーマータイプの判定
            let teamFunctionType;
            
            // 設定からしきい値を取得
            const hyperperformerProdThreshold = settings.hyperperformerProdThreshold || 10;
            const specialistComplexThreshold = settings.specialistComplexThreshold || 3;
            
            // チーム内パフォーマータイプの判定
            if (teamProductivity >= hyperperformerProdThreshold && teamFunctionCoverage >= 70) {
              teamFunctionType = "ハイパフォーマー";
            } else if (complexCaseCount >= specialistComplexThreshold && teamFunctionCoverage < 50) {
              teamFunctionType = "スペシャリスト";
            } else if (teamFunctionCoverage >= 70) {
              teamFunctionType = "オールラウンダー";
            } else if (teamFunctionCoverage < 30 || teamProductivity < 1.5) {
              teamFunctionType = "成長中";
            } else {
              teamFunctionType = "標準";
            }
            
            teamMembers.push({
              name: name,
              ihPt: ihPt,
              role: role,
              hcRatio: hcRatio,
              teamProductivity: teamProductivity,
              totalProductivity: totalProductivity,
              contributionRate: contributionRate,
              teamFunctionType: teamFunctionType,
              teamFunctionCoverage: teamFunctionCoverage,
              complexCaseCount: complexCaseCount,
              skillScore: skillScore,
              performerType: performerType
            });
            
            // マップにも追加（後で外部依存チェックに使用）
            teamMemberMap[name] = true;
          }
        }
        
        // 2. チーム概要データの算出
        const memberCount = teamMembers.filter(m => m.ihPt === 'IH').length;
        let totalHcRatio = 0;
        let avgTeamProductivity = 0;
        
        teamMembers.forEach(member => {
          if (member.ihPt === 'IH') { // IHのみ集計
            totalHcRatio += member.hcRatio;
            avgTeamProductivity += member.teamProductivity * member.hcRatio;
          }
        });
        
        avgTeamProductivity = totalHcRatio > 0 ? avgTeamProductivity / totalHcRatio : 0;
        
        // チーム内完結率の計算
        // チームが担当するタグに対する、チームIHメンバーの対応割合
        let teamTagTotalRally = 0;
        let teamMemberRally = 0;
  
        for (let i = 1; i < transactionData.length; i++) {
          const row = transactionData[i];
          const opTeam = row[txOpTeamIdx];
          const tag = row[txTagIdx];
          const name = row[txNameIdx];
          const ihPt = row[txIhPtIdx];
          const rally = row[txRallyIdx] || 1;
          
          if (tag && UtilsModule.isFunctionTag(tag) && opTeam === selectedTeam) {
            teamTagTotalRally += rally;
            
            if (teamMemberMap[name] && ihPt === 'IH') {
              teamMemberRally += rally;
            }
          }
        }
        
        const completionRate = teamTagTotalRally > 0 ? (teamMemberRally / teamTagTotalRally) * 100 : 0;
        
        // 3. IH外部リソース依存状況の分析
        const externalIHResponders = {};
        let totalExternalRally = 0;
        
        for (let i = 1; i < transactionData.length; i++) {
          const row = transactionData[i];
          const name = row[txNameIdx];
          const ihPt = row[txIhPtIdx];
          const opTeam = row[txOpTeamIdx];
          const tag = row[txTagIdx];
          const rally = row[txRallyIdx] || 1;
          const responseTime = row[txResponseTimeIdx] || 0;
          
          // チームのタグでIH外部リソースが対応
          if (tag && UtilsModule.isFunctionTag(tag) && 
              opTeam === selectedTeam && 
              ihPt === 'IH' && 
              !teamMemberMap[name]) {
            
            if (!externalIHResponders[name]) {
              externalIHResponders[name] = {
                name: name,
                rallyCount: 0,
                responseTime: 0,
                primaryTeam: ""
              };
            }
            
            externalIHResponders[name].rallyCount += rally;
            
            if (responseTime > 0 && responseTime < 1440) {
              externalIHResponders[name].responseTime += responseTime;
            }
            
            totalExternalRally += rally;
          }
        }
        
        // マスターデータから所属チーム情報を取得
        Object.keys(externalIHResponders).forEach(name => {
          for (let i = 1; i < masterData.length; i++) {
            if (masterData[i][masterNameIdx] === name) {
              externalIHResponders[name].primaryTeam = masterData[i][masterPrimaryTeamIdx] || "";
              break;
            }
          }
        });
        
        // 依存度と生産性を計算
        const externalDependencies = [];
        
        Object.values(externalIHResponders).forEach(responder => {
          const dependencyRate = totalExternalRally > 0 ? 
                              (responder.rallyCount / totalExternalRally) * 100 : 0;
          
          const productivity = responder.responseTime > 0 ? 
                              (responder.rallyCount / responder.responseTime) * 60 : 0;
          
          externalDependencies.push({
            name: responder.name,
            team: responder.primaryTeam,
            rallyCount: responder.rallyCount,
            dependencyRate: dependencyRate,
            productivity: productivity
          });
        });
        
        // 4. 理想HC設定の取得/計算
        // 保存されている設定を読み込むか、デフォルト値を使用
        const teamSettings = this.loadTeamSettings(performanceSheet, selectedTeam) || {
          rallyPerHour: 10.0,
          completionTarget: 85.0,
          difficultyFactor: 1.0
        };
        
        // 目標値の更新（ユーザーが入力した場合）
        const c34 = performanceSheet.getRange("C34").getValue();
        const c35 = performanceSheet.getRange("C35").getValue();
        const c36 = performanceSheet.getRange("C36").getValue();
        
        if (c34 && !isNaN(c34)) teamSettings.rallyPerHour = parseFloat(c34);
        if (c35 && !isNaN(c35)) teamSettings.completionTarget = parseFloat(c35);
        if (c36 && !isNaN(c36)) teamSettings.difficultyFactor = parseFloat(c36);
        
        // HC計算パラメータの取得
        const workDaysPerMonth = settings.workDaysPerMonth || 20;
        const hoursPerDay = settings.hoursPerDay || 7.5;
        const productivityRatio = settings.productivityRatio || 0.7;
        const targetRallyPerHour = teamSettings.rallyPerHour || 10;
        const targetCompletionRate = teamSettings.completionTarget || 85;
  
        // 理想HC計算
        const monthlyWorkHours = workDaysPerMonth * hoursPerDay * productivityRatio;
        const totalMonthlyRally = teamTagTotalRally / 3; // 3ヶ月分のデータを1ヶ月分に換算（サンプルに応じて調整）
  
        const targetRally = totalMonthlyRally * (targetCompletionRate / 100);
        const idealHC = (targetRally / (targetRallyPerHour * monthlyWorkHours));
        
        // HC差分計算
        const hcGap = totalHcRatio - idealHC;
        
        // チーム概要を表示
        performanceSheet.getRange("A9:E9").setValues([[
          selectedTeam,
          memberCount,
          completionRate.toFixed(1) + "%",
          totalHcRatio.toFixed(1),
          hcGap.toFixed(1)
        ]]);
        
        // 5. チームメンバーパフォーマンスの表示
        // パフォーマータイプ考慮で生産性スコアでソート
        teamMembers.sort((a, b) => {
          // パフォーマータイプ優先順位: ハイパフォーマー > スペシャリスト > オールラウンダー > 標準
          const typeOrder = {
            'ハイパフォーマー': 0,
            'スペシャリスト': 1,
            'オールラウンダー': 2,
            '標準': 3,
            '成長中': 4
          };
          
          const typeDiff = typeOrder[a.teamFunctionType] - typeOrder[b.teamFunctionType];
          if (typeDiff !== 0) return typeDiff;
          
          // 同じタイプならチームの生産性で降順ソート
          return b.teamProductivity - a.teamProductivity;
        });
        
        // メンバーデータの表示（最大10名）
        const displayMemberCount = Math.min(teamMembers.length, 10);
        
        // ヘッダー更新 - スキルレベルをチーム機能タイプに変更
        performanceSheet.getRange("H12").setValue("チーム機能タイプ");
        
        if (displayMemberCount > 0) {
          const memberData = [];
          
          for (let i = 0; i < displayMemberCount; i++) {
            const member = teamMembers[i];
            memberData.push([
              member.name,
              member.ihPt,
              member.role,
              member.hcRatio.toFixed(1),
              member.teamProductivity.toFixed(1),
              member.totalProductivity.toFixed(1),
              member.contributionRate.toFixed(1) + "%",
              member.teamFunctionType // スキルレベルをチーム機能タイプに変更
            ]);
          }
          
          performanceSheet.getRange(13, 1, displayMemberCount, 8).setValues(memberData);
          
          // 空白行のクリア
          if (displayMemberCount < 10) {
            performanceSheet.getRange(13 + displayMemberCount, 1, 10 - displayMemberCount, 8).clearContent();
          }
        } else {
          performanceSheet.getRange("A13:H13").merge()
            .setValue(`チーム「${selectedTeam}」にはメンバーが登録されていません`)
            .setFontStyle("italic");
          
          // 残りの行をクリア
          performanceSheet.getRange(14, 1, 9, 8).clearContent();
        }
        
        // 6. 外部リソース依存状況の表示
        // 依存度順にソート
        externalDependencies.sort((a, b) => b.rallyCount - a.rallyCount);
        
        // 外部依存データの表示（最大5名）
        const displayDepCount = Math.min(externalDependencies.length, 5);
        
        if (displayDepCount > 0) {
          const depData = [];
          
          for (let i = 0; i < displayDepCount; i++) {
            const dep = externalDependencies[i];
            depData.push([
              dep.name,
              dep.team,
              dep.rallyCount,
              dep.dependencyRate.toFixed(1) + "%",
              dep.productivity.toFixed(1)
            ]);
          }
          
          performanceSheet.getRange(26, 1, displayDepCount, 5).setValues(depData);
          
          // 空白行のクリア
          if (displayDepCount < 5) {
            performanceSheet.getRange(26 + displayDepCount, 1, 5 - displayDepCount, 5).clearContent();
          }
        } else {
          performanceSheet.getRange("A26:E26").merge()
            .setValue(`チーム「${selectedTeam}」の機能タグはIH外部リソースに依存していません`)
            .setFontStyle("italic");
          
          // 残りの行をクリア
          performanceSheet.getRange(27, 1, 4, 5).clearContent();
        }
        
        // 7. 理想HC設定の更新
        // 設定値の表示
        performanceSheet.getRange("B34").setValue(avgTeamProductivity.toFixed(1));
        performanceSheet.getRange("B35").setValue(completionRate.toFixed(1));
        performanceSheet.getRange("B36").setValue(teamSettings.difficultyFactor.toFixed(1));
        
        // 目標値を再表示
        performanceSheet.getRange("C34").setValue(teamSettings.rallyPerHour.toFixed(1));
        performanceSheet.getRange("C35").setValue(teamSettings.completionTarget.toFixed(1));
        performanceSheet.getRange("C36").setValue(teamSettings.difficultyFactor.toFixed(1));
        
        // 差分計算
        const rallyDiff = avgTeamProductivity - teamSettings.rallyPerHour;
        const completionDiff = completionRate - teamSettings.completionTarget;
        
        performanceSheet.getRange("D34").setValue(rallyDiff.toFixed(1));
        performanceSheet.getRange("D35").setValue(completionDiff.toFixed(1));
        performanceSheet.getRange("D36").setValue("0.0");
        
        // 必要HC表示
        performanceSheet.getRange("B37").setValue(idealHC.toFixed(1));
        performanceSheet.getRange("B38").setValue(hcGap.toFixed(1));
        
        // 条件付き書式の設定
        this.setTeamPerformanceFormatting(performanceSheet, displayMemberCount, displayDepCount);
        
        // 現在の設定を保存
        this.saveTeamSettings(performanceSheet, selectedTeam, {
          rallyPerHour: teamSettings.rallyPerHour,
          completionTarget: teamSettings.completionTarget,
          difficultyFactor: teamSettings.difficultyFactor
        });
        
        LogModule.logOperation("チームパフォーマンス更新", `完了: "${selectedTeam}" の分析`);
        UtilsModule.showToast(`${selectedTeam} のパフォーマンス分析が完了しました`, 'チームパフォーマンス');
        
        return true;
      } catch (e) {
        LogModule.handleError(e, 'チームパフォーマンス更新');
        return false;
      }
    },
  
    /**
     * チームパフォーマンスシートの書式を設定する
     */
    setTeamPerformanceFormatting: function(sheet, memberCount, depCount) {
      // チーム概要の書式設定
      const overviewRanges = {
        completionRate: sheet.getRange("C9"),
        hcGap: sheet.getRange("E9")
      };
      
      // チーム内完結率の書式設定
      const highCompletionRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("8")
        .setBackground('#C8E6C9') // 緑
        .setRanges([overviewRanges.completionRate])
        .build();
      
      const lowCompletionRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("5")
        .setBackground('#FFECB3') // 黄色
        .setRanges([overviewRanges.completionRate])
        .build();
      
      let rules = sheet.getConditionalFormatRules();
      rules.push(highCompletionRule, lowCompletionRule);
      sheet.setConditionalFormatRules(rules);
      
      // HC差分の書式設定
      const positiveHCRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0.5)
        .setBackground('#C8E6C9') // 緑
        .setRanges([overviewRanges.hcGap])
        .build();
      
      const negativeHCRule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(-0.5)
        .setBackground('#FFCDD2') // 赤
        .setRanges([overviewRanges.hcGap])
        .build();
      
      rules = sheet.getConditionalFormatRules();
      rules.push(positiveHCRule, negativeHCRule);
      sheet.setConditionalFormatRules(rules);
      
      // メンバーパフォーマンスの書式設定
      if (memberCount > 0) {
        // チーム貢献度の書式設定
        const contributionRange = sheet.getRange(13, 7, memberCount, 1);
        
        const highContributionRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("8")
          .setBackground('#C8E6C9') // 緑
          .setRanges([contributionRange])
          .build();
        
        const lowContributionRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("5")
          .setBackground('#FFECB3') // 黄色
          .setRanges([contributionRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(highContributionRule, lowContributionRule);
        sheet.setConditionalFormatRules(rules);
        
        // 生産性の書式設定
        const productivityRange = sheet.getRange(13, 5, memberCount, 1);
        
        const highProductivityRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(10)
          .setBackground('#C8E6C9') // 緑
          .setRanges([productivityRange])
          .build();
        
        const lowProductivityRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(5)
          .setBackground('#FFECB3') // 黄色
          .setRanges([productivityRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(highProductivityRule, lowProductivityRule);
        sheet.setConditionalFormatRules(rules);
        
        // チーム機能タイプの書式設定
        const teamFunctionTypeRange = sheet.getRange(13, 8, memberCount, 1);
        
        const teamFunctionTypeRules = [
          { value: 'ハイパフォーマー', color: '#1565C0', textColor: 'white' }, // 濃い青
          { value: 'スペシャリスト', color: '#6A1B9A', textColor: 'white' },   // 紫
          { value: 'オールラウンダー', color: '#2E7D32', textColor: 'white' }, // 緑
          { value: '標準', color: '#757575', textColor: 'white' },             // グレー
          { value: '成長中', color: '#FF9800', textColor: 'black' }            // オレンジ
        ];
        
        teamFunctionTypeRules.forEach(type => {
          const rule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(type.value)
            .setBackground(type.color)
            .setFontColor(type.textColor)
            .setRanges([teamFunctionTypeRange])
            .build();
          
          rules = sheet.getConditionalFormatRules();
          rules.push(rule);
          sheet.setConditionalFormatRules(rules);
        });
      }
  
      // 外部リソース依存の書式設定
      if (depCount > 0) {
        // 依存度の書式設定
        const dependencyRange = sheet.getRange(26, 4, depCount, 1);
        
        const highDependencyRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("3")
          .setBackground('#FFCDD2') // 赤
          .setRanges([dependencyRange])
          .build();
        
        const medDependencyRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("2")
          .setBackground('#FFECB3') // 黄色
          .setRanges([dependencyRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(highDependencyRule, medDependencyRule);
        sheet.setConditionalFormatRules(rules);
      }
      
      // 理想HC設定セクションの書式設定
      // 差分の書式設定
      const diffRanges = [
        sheet.getRange("D34"), // Rally/時間差分
        sheet.getRange("D35"), // 完結率差分
        sheet.getRange("B38")  // HC差分
      ];
      
      diffRanges.forEach(range => {
        const positiveDiffRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(0)
          .setBackground('#C8E6C9') // 緑
          .setRanges([range])
          .build();
        
        const negativeDiffRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(0)
          .setBackground('#FFCDD2') // 赤
          .setRanges([range])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(positiveDiffRule, negativeDiffRule);
        sheet.setConditionalFormatRules(rules);
      });
    },
  
    /**
     * チーム設定を読み込む
     */
    loadTeamSettings: function(sheet, teamName) {
      // デフォルト値
      const defaultSettings = {
        rallyPerHour: 10.0,
        completionTarget: 85.0,
        difficultyFactor: 1.0
      };
      
      try {
        // 設定保存領域を取得
        const settingsRange = sheet.getRange("A101:E200");
        const settingsData = settingsRange.getValues();
        
        // チーム名で検索
        for (let i = 0; i < settingsData.length; i++) {
          if (settingsData[i][1] === teamName) {
            return {
              rallyPerHour: parseFloat(settingsData[i][2]) || defaultSettings.rallyPerHour,
              completionTarget: parseFloat(settingsData[i][3]) || defaultSettings.completionTarget,
              difficultyFactor: parseFloat(settingsData[i][4]) || defaultSettings.difficultyFactor
            };
          }
        }
        
        // 見つからなければデフォルト値を返す
        return defaultSettings;
        
      } catch (e) {
        LogModule.handleError(e, 'チーム設定読み込み');
        return defaultSettings;
      }
    },
  
    /**
     * チーム設定を保存する
     */
    saveTeamSettings: function(sheet, teamName, settings) {
      try {
        // 設定保存領域を取得
        const settingsRange = sheet.getRange("A101:E200");
        const settingsData = settingsRange.getValues();
        
        // チーム名で検索
        let found = false;
        
        for (let i = 0; i < settingsData.length; i++) {
          if (settingsData[i][1] === teamName) {
            // 既存の設定を更新
            settingsData[i][2] = settings.rallyPerHour;
            settingsData[i][3] = settings.completionTarget;
            settingsData[i][4] = settings.difficultyFactor;
            found = true;
            break;
          }
        }
        
        if (!found) {
          // 新しい設定を追加
          let emptyRow = 0;
          for (let i = 0; i < settingsData.length; i++) {
            if (!settingsData[i][0] && !settingsData[i][1]) {
              emptyRow = i;
              break;
            }
          }
          
          if (emptyRow > 0) {
            settingsData[emptyRow] = [
              Date.now(), // チームID（タイムスタンプ）
              teamName,
              settings.rallyPerHour,
              settings.completionTarget,
              settings.difficultyFactor
            ];
          }
        }
        
        // 設定を書き戻す
        settingsRange.setValues(settingsData);
        
        return true;
        
      } catch (e) {
        LogModule.handleError(e, 'チーム設定保存');
        return false;
      }
    },
  
    /**
     * 機能領域カバレッジダッシュボードを初期化する
     */
    initAreaCoverageDashboard: function() {
      LogModule.logOperation("機能領域カバレッジダッシュボード初期化", "開始");
      
      try {
        // 必要なシートの取得
        const dataSheet = UtilsModule.getSheet(SHEETS.DATA);
        const masterSheet = UtilsModule.getSheet(SHEETS.MASTER);
        const difficultySheet = UtilsModule.getSheet(SHEETS.TAG_DIFFICULTY);
        const coverageSheet = UtilsModule.getOrCreateSheet(SHEETS.AREA_COVERAGE_DASHBOARD);
        
        // データの確認
        if (!dataSheet || !masterSheet || !difficultySheet) {
          throw new Error('必要なデータシートが見つかりません。先に分析を実行してください。');
        }
        
        // 機能領域カバレッジシートをクリア
        coverageSheet.clear();
        
        // タイトルと説明の設定
        coverageSheet.getRange("A1:G1").merge().setValue("機能領域カバレッジダッシュボード")
          .setFontWeight('bold').setBackground('#E8EAF6').setFontSize(16);
        
        const today = new Date();
        coverageSheet.getRange("A2:G2").merge().setValue(`更新日: ${today.toLocaleDateString()}`)
          .setFontStyle("italic");
        
        // 操作方法の説明
        coverageSheet.getRange("A3:G3").merge().setValue("使用方法: 機能領域を選択 → 「スキルマップ→ダッシュボード→機能領域カバレッジ更新」を実行")
          .setBackground('#FFF9C4').setFontWeight('bold');
        
        // 機能領域選択セクション
        coverageSheet.getRange("A5").setValue("機能領域:");
        
        // 取込データからOPチーム（機能領域）一覧を取得
        const opTeams = UtilsModule.getOPTeamList(dataSheet);
        
        // ドロップダウンリストの設定
        if (opTeams.length > 0) {
          const opTeamValidation = SpreadsheetApp.newDataValidation()
            .requireValueInList(opTeams, true)
            .build();
          coverageSheet.getRange("B5").setDataValidation(opTeamValidation);
          
          // デフォルト値として最初のOPチームを設定
          if (opTeams[0]) {
            coverageSheet.getRange("B5").setValue(opTeams[0]);
          }
        } else {
          coverageSheet.getRange("B5").setValue("(機能領域が見つかりません)");
        }
        
        // 対応状況サマリーセクション
        coverageSheet.getRange("A7:G7").merge().setValue("対応状況サマリー")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        // サマリーヘッダー
        coverageSheet.getRange("A8:G8").setValues([[
          "機能領域", "対応タグ数", "タグあたり理想HC", "所属チームIH対応率(%)", "他チームIH対応率(%)", "PT対応率(%)", "メンバーカバレッジ(%)"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // サマリーデータ領域の確保
        coverageSheet.getRange("A9:G9").setBackground('#f9f9f9');
        
        // リスクタグ表示セクション
        coverageSheet.getRange("A11:F11").merge().setValue("リスクタグ分析")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        // リスクタグヘッダー
        coverageSheet.getRange("A12:F12").setValues([[
          "タグ名", "難易度", "月間Rally数", "対応可能者数", "必要対応者数", "アクション"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // リスクタグデータ領域の確保（最大8件）
        coverageSheet.getRange("A13:F20").setBackground('#f9f9f9');
        
        // 対応状況内訳セクション
        coverageSheet.getRange("A22:E22").merge().setValue("対応状況内訳")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        // 分布ヘッダー
        coverageSheet.getRange("A23:E23").setValues([[
          "対応区分", "Rally数", "割合(%)", "対応者数", "平均生産性"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // 分布データ用に3行確保（所属チームIH、他チームIH、PT）
        coverageSheet.getRange("A24:E26").setBackground('#f9f9f9');
        
        // 対応者分布セクション
        coverageSheet.getRange("A28:F28").merge().setValue("対応者詳細 (上位15名)")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        // 対応者ヘッダー
        coverageSheet.getRange("A29:F29").setValues([[
          "担当者名", "IH/PT", "所属チーム", "対応Rally数", "対応タグ数", "生産性"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // 対応者データ領域の確保（最大15件）
        coverageSheet.getRange("A30:F44").setBackground('#f9f9f9');
        
        // 列幅の調整
        coverageSheet.setColumnWidth(1, 150); // タグ名/担当者名
        coverageSheet.setColumnWidth(3, 120); // 月間Rally数/所属チーム
        coverageSheet.setColumnWidth(6, 150); // アクション/生産性
        
        // シートをアクティブにする
        coverageSheet.activate();
        
        LogModule.logOperation("機能領域カバレッジダッシュボード初期化", "完了");
        UtilsModule.showToast('機能領域カバレッジダッシュボードを初期化しました', 'ダッシュボード初期化');
        
        return true;
      } catch (e) {
        LogModule.handleError(e, '機能領域カバレッジダッシュボード初期化');
        return false;
      }
    },
    
      /**
     * 機能領域カバレッジダッシュボードを更新する
     */
    updateAreaCoverage: function() {
      LogModule.logOperation("機能領域カバレッジ更新", "開始");
      
      try {
        // 必要なシートの取得
        const dataSheet = UtilsModule.getSheet(SHEETS.DATA);
        const masterSheet = UtilsModule.getSheet(SHEETS.MASTER);
        const difficultySheet = UtilsModule.getSheet(SHEETS.TAG_DIFFICULTY);
        const coverageSheet = UtilsModule.getSheet(SHEETS.AREA_COVERAGE_DASHBOARD);
        
        if (!dataSheet || !masterSheet || !difficultySheet || !coverageSheet) {
          throw new Error('必要なデータシートが見つかりません');
        }
        
        // 選択された機能領域を取得
        const selectedArea = coverageSheet.getRange("B5").getValue();
        
        if (!selectedArea || selectedArea === "(機能領域が見つかりません)") {
          coverageSheet.getRange("A9:G9").setValues([["機能領域が選択されていません", "", "", "", "", "", ""]]);
          
          // 他のデータをクリア
          coverageSheet.getRange("A13:F20").clearContent();
          coverageSheet.getRange("A24:E26").clearContent();
          coverageSheet.getRange("A30:F44").clearContent();
          
          return;
        }
        
        // データ準備
        const data = DataModule.prepareAnalysisData();
        if (!data) {
          throw new Error('分析データの準備に失敗しました');
        }
        
        // 機能領域（OPチーム）リストの取得
        const opTeams = UtilsModule.getOPTeamList(dataSheet);
        
        // タグの機能領域情報マッピング
        const tagAreaMap = {};
        const diffData = data.difficultyData;
        const diffColIndexes = data.diffColIndexes;
        
        for (let i = 1; i < diffData.length; i++) {
          const tagName = diffData[i][diffColIndexes.tagName];
          const opTeam = diffData[i][diffColIndexes.opTeam];
          const level = diffData[i][diffColIndexes.level];
          
          if (tagName && opTeam) {
            tagAreaMap[tagName] = {
              opTeam: opTeam,
              level: level
            };
          }
        }
        
        // 対応データの集計
        const txData = data.transactionData;
        const txColIndexes = data.txColIndexes;
        const memberMap = data.responderMasterMap;
        
        // 機能領域の詳細データ構造の初期化
        const areaDetail = {
          name: selectedArea,
          totalTags: 0,
          highLevelTags: 0, // L4, L5タグ
          monthlyRally: 0,  // 月間Rally数（推定）
          
          // タグデータ
          tags: {},
          
          // 対応者データ
          responders: {},
          
          // 対応区分データ
          ihPrimaryResponses: { rallyCount: 0, responders: new Set(), responseTime: 0 },
          ihSecondaryResponses: { rallyCount: 0, responders: new Set(), responseTime: 0 },
          ptResponses: { rallyCount: 0, responders: new Set(), responseTime: 0 }
        };
        
        // タグ情報の集計
        for (let i = 1; i < diffData.length; i++) {
          const row = diffData[i];
          const tagName = row[diffColIndexes.tagName];
          const opTeam = row[diffColIndexes.opTeam];
          const level = row[diffColIndexes.level];
          
          if (opTeam === selectedArea && UtilsModule.isFunctionTag(tagName)) {
            areaDetail.totalTags++;
            
            // 高難易度タグ（L4, L5）のカウント
            if (level === "L4" || level === "L5") {
              areaDetail.highLevelTags++;
            }
            
            // タグ情報の保存
            areaDetail.tags[tagName] = {
              name: tagName,
              level: level,
              rallyCount: 0,
              responders: new Set(),
              isHighRisk: false
            };
          }
        }
        
        // 対応データの集計
        const threeMonthsAgo = new Date();
        threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);
        
        let totalRallyCount = 0;
        
        for (let i = 1; i < txData.length; i++) {
          const row = txData[i];
          const tagName = row[txColIndexes.tagName];
          const userName = row[txColIndexes.userName];
          const ihPt = row[txColIndexes.ihPt];
          const opName = row[txColIndexes.opName] || (tagAreaMap[tagName] ? tagAreaMap[tagName].opTeam : "");
          const responseTime = row[txColIndexes.responseTime];
          const rallyCount = row[txColIndexes.rallyCount] || 1;
          
          // 選択された機能領域のタグのみ処理
          if (tagName && UtilsModule.isFunctionTag(tagName) && opName === selectedArea) {
            // タグの対応データ更新
            if (areaDetail.tags[tagName]) {
              areaDetail.tags[tagName].rallyCount += rallyCount;
              areaDetail.tags[tagName].responders.add(userName);
            }
            
            totalRallyCount += rallyCount;
            
            // 対応者データの更新
            if (!areaDetail.responders[userName]) {
              areaDetail.responders[userName] = {
                name: userName,
                ihPt: ihPt,
                primaryTeam: memberMap[userName] ? memberMap[userName].primaryTeam : "",
                rallyCount: 0,
                responseTime: 0,
                tags: new Set()
              };
            }
            
            areaDetail.responders[userName].rallyCount += rallyCount;
            areaDetail.responders[userName].tags.add(tagName);
            
            if (responseTime > 0 && responseTime < 1440) {
              areaDetail.responders[userName].responseTime += responseTime;
            }
            
            // 対応区分別の集計
            if (ihPt === 'IH') {
              const memberPrimaryTeam = memberMap[userName] ? memberMap[userName].primaryTeam : "";
              
              if (memberPrimaryTeam === selectedArea) {
                // 所属チームIHの対応
                areaDetail.ihPrimaryResponses.rallyCount += rallyCount;
                areaDetail.ihPrimaryResponses.responders.add(userName);
                
                if (responseTime > 0 && responseTime < 1440) {
                  areaDetail.ihPrimaryResponses.responseTime += responseTime;
                }
              } else {
                // 他チームIHの対応
                areaDetail.ihSecondaryResponses.rallyCount += rallyCount;
                areaDetail.ihSecondaryResponses.responders.add(userName);
                
                if (responseTime > 0 && responseTime < 1440) {
                  areaDetail.ihSecondaryResponses.responseTime += responseTime;
                }
              }
            } else {
              // PTの対応
              areaDetail.ptResponses.rallyCount += rallyCount;
              areaDetail.ptResponses.responders.add(userName);
              
              if (responseTime > 0 && responseTime < 1440) {
                areaDetail.ptResponses.responseTime += responseTime;
              }
            }
          }
        }
        
        // 月間Rally数の推定（3ヶ月分のデータを元に）
        areaDetail.monthlyRally = Math.round(totalRallyCount / 3);
        
        // 対応区分別の割合と生産性計算
        const totalRally = totalRallyCount || 1; // ゼロ除算防止
        
        areaDetail.ihPrimaryRallyRate = (areaDetail.ihPrimaryResponses.rallyCount / totalRally) * 100;
        areaDetail.ihSecondaryRallyRate = (areaDetail.ihSecondaryResponses.rallyCount / totalRally) * 100;
        areaDetail.ptRallyRate = (areaDetail.ptResponses.rallyCount / totalRally) * 100;
        
        // 生産性計算（Rally/時間）
        areaDetail.ihPrimaryProductivity = areaDetail.ihPrimaryResponses.responseTime > 0 ? 
                                          (areaDetail.ihPrimaryResponses.rallyCount / areaDetail.ihPrimaryResponses.responseTime) * 60 : 0;
        
        areaDetail.ihSecondaryProductivity = areaDetail.ihSecondaryResponses.responseTime > 0 ? 
                                            (areaDetail.ihSecondaryResponses.rallyCount / areaDetail.ihSecondaryResponses.responseTime) * 60 : 0;
        
        areaDetail.ptProductivity = areaDetail.ptResponses.responseTime > 0 ? 
                                  (areaDetail.ptResponses.rallyCount / areaDetail.ptResponses.responseTime) * 60 : 0;
        
        // 各対応者の生産性計算
        Object.values(areaDetail.responders).forEach(responder => {
          responder.productivity = responder.responseTime > 0 ? 
                                (responder.rallyCount / responder.responseTime) * 60 : 0;
        });
        
        // リスクタグ分析
        const riskTags = [];
        
        Object.values(areaDetail.tags).forEach(tag => {
          // 難易度に応じた必要対応者数
          let requiredResponders;
          switch (tag.level) {
            case "L5": requiredResponders = 5; break;
            case "L4": requiredResponders = 4; break;
            case "L3": requiredResponders = 3; break;
            default: requiredResponders = 2;
          }
          
          const currentResponders = tag.responders.size;
          const isRisk = (tag.level === "L4" || tag.level === "L5") && currentResponders < requiredResponders;
          
          // リスクタグを追加
          if (isRisk) {
            let action;
            if (tag.level === "L5" && currentResponders <= 1) {
              action = "緊急スキル獲得";
            } else if (currentResponders <= requiredResponders - 2) {
              action = "スキル獲得";
            } else {
              action = "計画的育成";
            }
            
            riskTags.push([
              tag.name,
              tag.level,
              Math.round(tag.rallyCount / 3), // 月間Rally数
              currentResponders,
              requiredResponders,
              action
            ]);
          }
        });
        
        // 優先度順にソート（L5優先、次に対応者不足の大きいもの）
        riskTags.sort((a, b) => {
          if (a[1] === "L5" && b[1] !== "L5") return -1;
          if (a[1] !== "L5" && b[1] === "L5") return 1;
          
          const aGap = a[4] - a[3];
          const bGap = b[4] - b[3];
          return bGap - aGap;
        });
        
        // 対応状況サマリーセクションの更新
        // タグあたり理想HC計算
        const tagsCount = areaDetail.totalTags || 1; // ゼロ除算防止
        const monthlyWorkHours = 20 * 7.5 * 0.7; // 20営業日 × 7.5時間 × 生産性70%
        const idealHcPerTag = areaDetail.monthlyRally / (tagsCount * 10 * monthlyWorkHours); // 10は理想Rally/時間
        
        // メンバーカバレッジ計算
        const memberCoverage = tagsCount > 0 ? 
                            areaDetail.ihPrimaryResponses.responders.size / tagsCount * 100 : 0;
        
        coverageSheet.getRange("A9:G9").setValues([[
          selectedArea,
          areaDetail.totalTags,
          idealHcPerTag.toFixed(2),
          areaDetail.ihPrimaryRallyRate.toFixed(1) + "%",
          areaDetail.ihSecondaryRallyRate.toFixed(1) + "%",
          areaDetail.ptRallyRate.toFixed(1) + "%",
          memberCoverage.toFixed(1) + "%"
        ]]);
        
        // リスクタグ分析セクションの更新
        if (riskTags.length > 0) {
          const riskTagCount = Math.min(riskTags.length, 8);
          coverageSheet.getRange(13, 1, riskTagCount, 6).setValues(riskTags.slice(0, riskTagCount));
          
          // 空白行のクリア
          if (riskTagCount < 8) {
            coverageSheet.getRange(13 + riskTagCount, 1, 8 - riskTagCount, 6).clearContent();
          }
        } else {
          coverageSheet.getRange("A13:F13").merge()
            .setValue(`機能領域「${selectedArea}」には対応者不足のリスクタグがありません`)
            .setFontStyle("italic");
          
          // 残りの行をクリア
          coverageSheet.getRange(14, 1, 7, 6).clearContent();
        }
        
        // 対応状況内訳の更新
        const distData = [
          ["所属チームIH", areaDetail.ihPrimaryResponses.rallyCount, areaDetail.ihPrimaryRallyRate.toFixed(1) + "%", 
          areaDetail.ihPrimaryResponses.responders.size, areaDetail.ihPrimaryProductivity.toFixed(1)],
          ["他チームIH", areaDetail.ihSecondaryResponses.rallyCount, areaDetail.ihSecondaryRallyRate.toFixed(1) + "%", 
          areaDetail.ihSecondaryResponses.responders.size, areaDetail.ihSecondaryProductivity.toFixed(1)],
          ["PT", areaDetail.ptResponses.rallyCount, areaDetail.ptRallyRate.toFixed(1) + "%", 
          areaDetail.ptResponses.responders.size, areaDetail.ptProductivity.toFixed(1)]
        ];
        
        coverageSheet.getRange(24, 1, 3, 5).setValues(distData);
        
        // 対応者詳細の更新
        const responderList = Object.values(areaDetail.responders)
          .sort((a, b) => b.rallyCount - a.rallyCount)
          .map(r => [
            r.name,
            r.ihPt,
            r.primaryTeam,
            r.rallyCount,
            r.tags.size,
            r.productivity.toFixed(1)
          ]);
        
        if (responderList.length > 0) {
          const responderCount = Math.min(responderList.length, 15);
          coverageSheet.getRange(30, 1, responderCount, 6).setValues(responderList.slice(0, responderCount));
          
          // 空白行のクリア
          if (responderCount < 15) {
            coverageSheet.getRange(30 + responderCount, 1, 15 - responderCount, 6).clearContent();
          }
        } else {
          coverageSheet.getRange("A30:F30").merge()
            .setValue(`機能領域「${selectedArea}」には対応者データがありません`)
            .setFontStyle("italic");
          
          // 残りの行をクリア
          coverageSheet.getRange(31, 1, 14, 6).clearContent();
        }
        
        // 書式設定
        this.setAreaCoverageDashboardFormatting(coverageSheet, riskTags.length, responderList.length);
        
        LogModule.logOperation("機能領域カバレッジ更新", `完了: "${selectedArea}" の分析`);
        UtilsModule.showToast(`${selectedArea} の機能領域カバレッジ分析が完了しました`, '機能領域カバレッジ');
        
        return true;
      } catch (e) {
        LogModule.handleError(e, '機能領域カバレッジ更新');
        return false;
      }
    },
  
    /**
     * 機能領域カバレッジダッシュボードの書式を設定する
     */
    setAreaCoverageDashboardFormatting: function(sheet, riskTagCount, responderCount) {
      // 対応状況サマリーの書式設定
      const summaryRanges = {
        ihPrimaryRate: sheet.getRange("D9"),
        ihSecondaryRate: sheet.getRange("E9"),
        ptRate: sheet.getRange("F9"),
        memberCoverage: sheet.getRange("G9")
      };
      
      // 所属チームIH対応率の書式設定
      const highPrimaryRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("8")
        .setBackground('#C8E6C9') // 緑
        .setRanges([summaryRanges.ihPrimaryRate])
        .build();
      
      const midPrimaryRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("5")
        .setBackground('#FFF9C4') // 黄色
        .setRanges([summaryRanges.ihPrimaryRate])
        .build();
      
      const lowPrimaryRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("3")
        .setBackground('#FFEBEE') // 薄い赤
        .setRanges([summaryRanges.ihPrimaryRate])
        .build();
      
      let rules = sheet.getConditionalFormatRules();
      rules.push(highPrimaryRule, midPrimaryRule, lowPrimaryRule);
      sheet.setConditionalFormatRules(rules);
      
      // PT対応率の書式設定
      const highPTRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("5")
        .setBackground('#FFEBEE') // 薄い赤
        .setRanges([summaryRanges.ptRate])
        .build();
      
      rules = sheet.getConditionalFormatRules();
      rules.push(highPTRule);
      sheet.setConditionalFormatRules(rules);
      
      // メンバーカバレッジの書式設定
      const highCoverageRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("8")
        .setBackground('#C8E6C9') // 緑
        .setRanges([summaryRanges.memberCoverage])
        .build();
      
      const lowCoverageRule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("3")
        .setBackground('#FFEBEE') // 薄い赤
        .setRanges([summaryRanges.memberCoverage])
        .build();
      
      rules = sheet.getConditionalFormatRules();
      rules.push(highCoverageRule, lowCoverageRule);
      sheet.setConditionalFormatRules(rules);
      
      // リスクタグの書式設定
      if (riskTagCount > 0) {
        // 難易度列の書式設定
        const levelRange = sheet.getRange(13, 2, riskTagCount, 1);
        
        const l5Rule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("L5")
          .setBackground('#0D47A1') // 濃い青
          .setFontColor('white')
          .setRanges([levelRange])
          .build();
        
        const l4Rule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("L4")
          .setBackground('#1976D2') // 青
          .setFontColor('white')
          .setRanges([levelRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(l5Rule, l4Rule);
        sheet.setConditionalFormatRules(rules);
        
        // アクション列の書式設定
        const actionRange = sheet.getRange(13, 6, riskTagCount, 1);
        
        const urgentRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("緊急")
          .setBackground('#FFCDD2') // 赤
          .setRanges([actionRange])
          .build();
        
        const acquireRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("スキル獲得")
          .setBackground('#FFECB3') // 黄色
          .setRanges([actionRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(urgentRule, acquireRule);
        sheet.setConditionalFormatRules(rules);
      }
      
      // 対応者詳細の書式設定
      if (responderCount > 0) {
        // IH/PT列の書式設定
        const ihPtRange = sheet.getRange(30, 2, responderCount, 1);
        
        const ihRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("IH")
          .setBackground('#E3F2FD') // 薄い青
          .setRanges([ihPtRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(ihRule);
        sheet.setConditionalFormatRules(rules);
        
        // 所属チーム列の書式設定
        const teamRange = sheet.getRange(30, 3, responderCount, 1);
        
        const selectedTeamRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo(sheet.getRange("B5").getValue())
          .setBackground('#E8F5E9') // 薄い緑
          .setRanges([teamRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(selectedTeamRule);
        sheet.setConditionalFormatRules(rules);
      }
    },
  
    /**
     * 相談フローダッシュボードを初期化する
     */
    initConsultationFlowDashboard: function() {
      LogModule.logOperation("相談フローダッシュボード初期化", "開始");
      
      try {
        // 必要なシートの取得
        const masterSheet = UtilsModule.getSheet(SHEETS.MASTER);
        const dataSheet = UtilsModule.getSheet(SHEETS.DATA);
        const consultationFlowDashSheet = UtilsModule.getOrCreateSheet(SHEETS.CONSULTATION_FLOW_DASHBOARD);
        
        // データの確認
        if (!masterSheet || !dataSheet) {
          throw new Error('必要なデータシートが見つかりません。先に分析を実行してください。');
        }
        
        // シートをクリア
        consultationFlowDashSheet.clear();
        
        // タイトルと説明の設定
        consultationFlowDashSheet.getRange("A1:I1").merge().setValue("相談フローダッシュボード")
          .setFontWeight('bold').setBackground('#E8EAF6').setFontSize(16);
        
        const today = new Date();
        consultationFlowDashSheet.getRange("A2:I2").merge().setValue(`更新日: ${today.toLocaleDateString()}`)
          .setFontStyle("italic");
        
        // 操作方法の説明
        consultationFlowDashSheet.getRange("A3:I3").merge().setValue("使用方法: チームを選択 → 「スキルマップ→ダッシュボード→相談フローダッシュボード更新」を実行")
          .setBackground('#FFF9C4').setFontWeight('bold');
        
        // チーム選択セクション
        consultationFlowDashSheet.getRange("A5").setValue("チーム選択:");
        
        // マスターデータからチーム一覧を取得
        const teams = UtilsModule.getTeamList(masterSheet);
        
        // ドロップダウンリストの設定
        if (teams.length > 0) {
          const teamValidation = SpreadsheetApp.newDataValidation()
            .requireValueInList(teams, true)
            .build();
          consultationFlowDashSheet.getRange("B5").setDataValidation(teamValidation);
          
          // デフォルト値として最初のチームを設定
          if (teams[0]) {
            consultationFlowDashSheet.getRange("B5").setValue(teams[0]);
          }
        } else {
          consultationFlowDashSheet.getRange("B5").setValue("(チームが見つかりません)");
        }
        
        // 相談フロー概要セクション
        consultationFlowDashSheet.getRange("A7:I7").merge().setValue("相談フロー概要")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        // 概要ヘッダー
        consultationFlowDashSheet.getRange("A8:E8").setValues([[
          "相談総数", "平均解決時間(分)", "所属チーム内解決率(%)", "他チーム依存率(%)", "ナレッジ化候補数"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // 概要データ領域の確保
        consultationFlowDashSheet.getRange("A9:E9").setBackground('#f9f9f9');
        
        // カテゴリ間フローセクション
        consultationFlowDashSheet.getRange("A11:I11").merge().setValue("カテゴリ間相談フロー（上位8件）")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        // フローヘッダー
        consultationFlowDashSheet.getRange("A12:I12").setValues([[
          "相談元", "相談先", "件数", "割合(%)", "平均解決時間(分)", 
          "Rally総数", "分あたりRally数", "ナレッジ化優先度", "アクション"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // フローデータ領域の確保
        consultationFlowDashSheet.getRange("A13:I20").setBackground('#f9f9f9');
        
        // 機能タグ別相談セクション
        consultationFlowDashSheet.getRange("A22:G22").merge().setValue("機能タグ別相談状況（上位8件）")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        // タグヘッダー
        consultationFlowDashSheet.getRange("A23:G23").setValues([[
          "タグ名", "相談件数", "割合(%)", "平均解決時間(分)", 
          "主な質問者", "主な回答者", "対応改善アクション"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // タグデータ領域の確保
        consultationFlowDashSheet.getRange("A24:G31").setBackground('#f9f9f9');
        
        // 対応者ペアセクション
        consultationFlowDashSheet.getRange("A33:F33").merge().setValue("相談頻度の高い対応者ペア（上位5件）")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        // ペアヘッダー
        consultationFlowDashSheet.getRange("A34:F34").setValues([[
          "質問者", "回答者", "件数", "平均解決時間(分)", "対応タグ数", "ペア生産性"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // ペアデータ領域の確保
        consultationFlowDashSheet.getRange("A35:F39").setBackground('#f9f9f9');
        
        // リソース最適化提案セクション
        consultationFlowDashSheet.getRange("A41:C41").merge().setValue("リソース最適化提案")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        // 提案ヘッダー
        consultationFlowDashSheet.getRange("A42:C42").setValues([[
          "提案項目", "詳細", "期待効果"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // 提案データ領域の確保
        consultationFlowDashSheet.getRange("A43:C47").setBackground('#f9f9f9');
        
        // 列幅の調整
        consultationFlowDashSheet.setColumnWidth(1, 150); // 相談元/タグ名/質問者
        consultationFlowDashSheet.setColumnWidth(2, 150); // 相談先/主な質問者/回答者
        consultationFlowDashSheet.setColumnWidth(3, 80);  // 件数
        consultationFlowDashSheet.setColumnWidth(4, 80);  // 割合(%)
        consultationFlowDashSheet.setColumnWidth(5, 120); // 平均解決時間(分)
        consultationFlowDashSheet.setColumnWidth(7, 120); // 分あたりRally数
        consultationFlowDashSheet.setColumnWidth(9, 150); // アクション
        
        // シートをアクティブにする
        consultationFlowDashSheet.activate();
        
        LogModule.logOperation("相談フローダッシュボード初期化", "完了");
        UtilsModule.showToast('相談フローダッシュボードを初期化しました', 'ダッシュボード初期化');
        
        return true;
      } catch (e) {
        LogModule.handleError(e, '相談フローダッシュボード初期化');
        return false;
      }
    },
  
    /**
     * 相談フローダッシュボードを更新する
     */
    updateConsultationFlow: function() {
      LogModule.logOperation("相談フローダッシュボード更新", "開始");
      
      try {
        // 必要なシートの取得
        const consultationFlowSheet = UtilsModule.getSheet(SHEETS.CONSULTATION_FLOW);
        const consultationFlowDashSheet = UtilsModule.getSheet(SHEETS.CONSULTATION_FLOW_DASHBOARD);
        
        if (!consultationFlowSheet || !consultationFlowDashSheet) {
          throw new Error('必要なシートが見つかりません。先に分析と初期化を実行してください。');
        }
        
        // 選択されたチームを取得
        const selectedTeam = consultationFlowDashSheet.getRange("B5").getValue();
        
        if (!selectedTeam || selectedTeam === "(チームが見つかりません)") {
          consultationFlowDashSheet.getRange("A9").setValue("チームが選択されていません");
          return false;
        }
        
        // 相談フロー分析シートからデータを抽出
        const consultationData = this.extractConsultationFlowData(consultationFlowSheet, selectedTeam);
        
        if (!consultationData) {
          throw new Error(`選択されたチーム "${selectedTeam}" の相談フローデータが見つかりません`);
        }
        
        // 相談フロー概要データの表示
        consultationFlowDashSheet.getRange("A9:E9").setValues([[
          consultationData.totalConsultations,
          consultationData.avgSolveDuration,
          consultationData.internalResolutionRate + "%",
          consultationData.externalDependencyRate + "%",
          consultationData.knowledgeCandidateCount
        ]]);
        
        // カテゴリ間フローデータの表示
        if (consultationData.categoryFlows.length > 0) {
          const displayCount = Math.min(consultationData.categoryFlows.length, 8);
          
          consultationFlowDashSheet.getRange(13, 1, displayCount, 9)
            .setValues(consultationData.categoryFlows.slice(0, displayCount));
          
          // 空白行のクリア
          if (displayCount < 8) {
            consultationFlowDashSheet.getRange(13 + displayCount, 1, 8 - displayCount, 9).clearContent();
          }
        } else {
          consultationFlowDashSheet.getRange("A13:I13").merge()
            .setValue(`チーム "${selectedTeam}" にはカテゴリ間フローデータがありません`)
            .setFontStyle("italic");
          
          consultationFlowDashSheet.getRange(14, 1, 7, 9).clearContent();
        }
        
        // 機能タグ別相談データの表示
        if (consultationData.tagConsultations.length > 0) {
          const displayCount = Math.min(consultationData.tagConsultations.length, 8);
          
          consultationFlowDashSheet.getRange(24, 1, displayCount, 7)
            .setValues(consultationData.tagConsultations.slice(0, displayCount));
          
          // 空白行のクリア
          if (displayCount < 8) {
            consultationFlowDashSheet.getRange(24 + displayCount, 1, 8 - displayCount, 7).clearContent();
          }
        } else {
          consultationFlowDashSheet.getRange("A24:G24").merge()
            .setValue(`チーム "${selectedTeam}" には機能タグ別相談データがありません`)
            .setFontStyle("italic");
          
          consultationFlowDashSheet.getRange(25, 1, 7, 7).clearContent();
        }
        
        // 対応者ペアデータの表示
        if (consultationData.personalFlows.length > 0) {
          const displayCount = Math.min(consultationData.personalFlows.length, 5);
          
          // カテゴリ表示を省略した簡易版をダッシュボードに表示
          const pairData = consultationData.personalFlows.slice(0, displayCount).map(flow => [
            flow[0], // 質問者
            flow[2], // 回答者
            flow[4], // 件数
            flow[6], // 平均解決時間
            flow[7], // 対応タグ数
            flow[8]  // ペア生産性
          ]);
          
          consultationFlowDashSheet.getRange(35, 1, displayCount, 6)
            .setValues(pairData);
          
          // 空白行のクリア
          if (displayCount < 5) {
            consultationFlowDashSheet.getRange(35 + displayCount, 1, 5 - displayCount, 6).clearContent();
          }
        } else {
          consultationFlowDashSheet.getRange("A35:F35").merge()
            .setValue(`チーム "${selectedTeam}" には対応者ペアデータがありません`)
            .setFontStyle("italic");
          
          consultationFlowDashSheet.getRange(36, 1, 4, 6).clearContent();
        }
        
        // リソース最適化提案の生成と表示
        const optimizationProposals = this.generateOptimizationProposals(consultationData, selectedTeam);
        
        if (optimizationProposals.length > 0) {
          const displayCount = Math.min(optimizationProposals.length, 5);
          
          consultationFlowDashSheet.getRange(43, 1, displayCount, 3)
            .setValues(optimizationProposals.slice(0, displayCount));
          
          // 空白行のクリア
          if (displayCount < 5) {
            consultationFlowDashSheet.getRange(43 + displayCount, 1, 5 - displayCount, 3).clearContent();
          }
        } else {
          consultationFlowDashSheet.getRange("A43:C43").merge()
            .setValue("最適化提案を生成するためのデータが不足しています")
            .setFontStyle("italic");
          
          consultationFlowDashSheet.getRange(44, 1, 4, 3).clearContent();
        }
        
        // 書式設定
        this.setConsultationFlowDashboardFormatting(consultationFlowDashSheet);
        
        LogModule.logOperation("相談フローダッシュボード更新", `完了: "${selectedTeam}" の分析`);
        UtilsModule.showToast(`${selectedTeam} の相談フローダッシュボードが更新されました`, '相談フローダッシュボード');
        
        return true;
      } catch (e) {
        LogModule.handleError(e, '相談フローダッシュボード更新');
        return false;
      }
    },
  
    /**
     * 相談フロー分析シートからデータを抽出する
     */
    extractConsultationFlowData: function(sheet, teamName) {
      try {
        const data = sheet.getDataRange().getValues();
        
        // チーム概要を探す
        let teamRow = -1;
        for (let i = 5; i < data.length; i++) {
          if (data[i][0] === teamName) {
            teamRow = i;
            break;
          }
        }
        
        if (teamRow === -1) {
          return null; // チームが見つからない
        }
        
        // 概要データの抽出
        const totalConsultations = data[teamRow][1];
        const avgSolveDuration = parseFloat(data[teamRow][2]);
        
        // IH所属チーム内での解決率
        const internalResolutionRate = parseFloat(data[teamRow][4]);
        
        // 他チーム依存率（IH他チーム受信率＋PT受信率）
        const externalDependencyRate = parseFloat(data[teamRow][6]) + parseFloat(data[teamRow][8]) + parseFloat(data[teamRow][10]);
        
        // カテゴリ間フローデータを探す
        let categoryFlowStart = -1;
        let categoryFlowEnd = -1;
        
        for (let i = teamRow + 1; i < data.length; i++) {
          if (data[i][0] === "相談元" && data[i][1] === "相談先") {
            categoryFlowStart = i + 1;
          } else if (categoryFlowStart !== -1 && (data[i][0] === "" || i === data.length - 1 || data[i][0].includes("機能タグ別"))) {
            categoryFlowEnd = i;
            break;
          }
        }
        
        const categoryFlows = [];
        if (categoryFlowStart !== -1 && categoryFlowEnd !== -1) {
          for (let i = categoryFlowStart; i < categoryFlowEnd; i++) {
            if (data[i][0] && data[i][2] > 0) {
              categoryFlows.push(data[i].slice(0, 9));
            }
          }
        }
        
        // ナレッジ化候補数のカウント
        const knowledgeCandidateCount = categoryFlows.filter(flow => 
          flow[7] === "高" || (flow[8] && flow[8].includes("ナレッジ"))
        ).length;
        
        // 機能タグ別相談データを探す
        let tagConsultStart = -1;
        let tagConsultEnd = -1;
        
        for (let i = categoryFlowEnd || 0; i < data.length; i++) {
          if (data[i][0] === "タグ名" && data[i][1] === "相談件数") {
            tagConsultStart = i + 1;
          } else if (tagConsultStart !== -1 && (data[i][0] === "" || i === data.length - 1 || data[i][0].includes("個人間相談"))) {
            tagConsultEnd = i;
            break;
          }
        }
        
        const tagConsultations = [];
        if (tagConsultStart !== -1 && tagConsultEnd !== -1) {
          for (let i = tagConsultStart; i < tagConsultEnd; i++) {
            if (data[i][0] && data[i][1] > 0) {
              tagConsultations.push(data[i].slice(0, 7));
            }
          }
        }
        
        // 個人間相談フローデータを探す
        let personalFlowStart = -1;
        let personalFlowEnd = -1;
        
        for (let i = tagConsultEnd || 0; i < data.length; i++) {
          if (data[i][0] === "質問者" && data[i][2] === "回答者") {
            personalFlowStart = i + 1;
          } else if (personalFlowStart !== -1 && (data[i][0] === "" || i === data.length - 1 || i > personalFlowStart + 15)) {
            personalFlowEnd = i;
            break;
          }
        }
        
        const personalFlows = [];
        if (personalFlowStart !== -1 && personalFlowEnd !== -1) {
          for (let i = personalFlowStart; i < personalFlowEnd; i++) {
            if (data[i][0] && data[i][4] > 0) {
              personalFlows.push(data[i].slice(0, 9));
            }
          }
        }
        
        return {
          teamName: teamName,
          totalConsultations: totalConsultations,
          avgSolveDuration: avgSolveDuration,
          internalResolutionRate: internalResolutionRate.toFixed(1),
          externalDependencyRate: externalDependencyRate.toFixed(1),
          knowledgeCandidateCount: knowledgeCandidateCount,
          categoryFlows: categoryFlows,
          tagConsultations: tagConsultations,
          personalFlows: personalFlows
        };
        
      } catch (e) {
        LogModule.handleError(e, '相談フローデータ抽出');
        return null;
      }
    },
  
    /**
     * リソース最適化提案を生成する
     */
    generateOptimizationProposals: function(consultationData, teamName) {
      const proposals = [];
      
      if (!consultationData) return proposals;
      
      // 提案1: ナレッジ化
      if (consultationData.knowledgeCandidateCount > 0) {
        const tagSuggestions = consultationData.tagConsultations
          .filter(tag => tag[6] && tag[6].includes("ナレッジ"))
          .map(tag => tag[0])
          .slice(0, 3)
          .join(", ");
        
        proposals.push([
          "ナレッジ作成",
          `頻繁に相談が発生している ${consultationData.knowledgeCandidateCount} 件のフローについてナレッジ化。特に ${tagSuggestions} などのタグを優先`,
          "相談対応コストの削減、解決時間の短縮"
        ]);
      }
      
      // 提案2: スキル移管
      const externalDependency = parseFloat(consultationData.externalDependencyRate);
      if (externalDependency > 30) {
        const topExternalFlows = consultationData.categoryFlows
          .filter(flow => !flow[0].includes("IH所属チーム") && flow[1].includes("IH所属チーム"))
          .slice(0, 2);
        
        if (topExternalFlows.length > 0) {
          const flowDetails = topExternalFlows.map(flow => 
            `${flow[0]}→${flow[1]} (${flow[2]}件)`
          ).join(", ");
          
          proposals.push([
            "スキル移管計画",
            `他チームからの相談依存度が高い (${consultationData.externalDependencyRate}%)。${flowDetails} などのフローを分析しスキル移管を検討`,
            "チーム内完結率の向上、依存度の低減"
          ]);
        }
      }
      
      // 提案3: 研修企画
      const internalConsults = consultationData.categoryFlows
        .filter(flow => flow[0].includes("IH所属チーム") && flow[1].includes("IH所属チーム"))
        .reduce((sum, flow) => sum + flow[2], 0);
      
      const internalConsultRatio = consultationData.totalConsultations > 0 ?
        (internalConsults / consultationData.totalConsultations) * 100 : 0;
      
      if (internalConsultRatio > 40) {
        proposals.push([
          "チーム内研修企画",
          `チーム内の相談が多い (全体の${internalConsultRatio.toFixed(1)}%)。頻出タグに関する定期研修の実施を検討`,
          "チーム全体のスキル底上げ、相談数の削減"
        ]);
      }
      
      // 提案4: 最適なペア編成
      if (consultationData.personalFlows.length >= 3) {
        const topPairs = consultationData.personalFlows
          .filter(flow => flow[4] >= 5) // 5件以上相談があるペア
          .slice(0, 2)
          .map(flow => `${flow[0]}-${flow[2]}`).join(", ");
        
        if (topPairs) {
          proposals.push([
            "ペア制導入検討",
            `相談頻度が高いペア (${topPairs}) を活かした業務ペア制の導入`,
            "知識移転の促進、チーム内完結率の向上"
          ]);
        }
      }
      
      // 提案5: PTリソース最適化
      const ptInvolvement = consultationData.categoryFlows
        .filter(flow => flow[0].includes("PT") || flow[1].includes("PT"))
        .reduce((sum, flow) => sum + flow[2], 0);
      
      const ptRatio = consultationData.totalConsultations > 0 ?
        (ptInvolvement / consultationData.totalConsultations) * 100 : 0;
      
      if (ptRatio > 25) {
        proposals.push([
          "PTリソース最適化",
          `PT関連の相談が多い (全体の${ptRatio.toFixed(1)}%)。PTに対するナレッジ提供や研修を強化`,
          "PT対応率の維持・向上、IHリソースの効率化"
        ]);
      }
      
      return proposals;
    },
  
    /**
     * 相談フローダッシュボードの書式を設定する
     */
    setConsultationFlowDashboardFormatting: function(sheet) {
      try {
        // 概要セクションの書式設定
        // 相談総数
        const consultationCountRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(100)
          .setBackground('#C8E6C9') // 緑
          .setRanges([sheet.getRange("A9")])
          .build();
        
        let rules = sheet.getConditionalFormatRules();
        rules.push(consultationCountRule);
        sheet.setConditionalFormatRules(rules);
        
        // 平均解決時間
        const longDurationRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(15)
          .setBackground('#FFEBEE') // 薄い赤
          .setRanges([sheet.getRange("B9")])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(longDurationRule);
        sheet.setConditionalFormatRules(rules);
        
        // 所属チーム内解決率
        const highResolutionRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("7")
          .setBackground('#C8E6C9') // 緑
          .setRanges([sheet.getRange("C9")])
          .build();
        
        const lowResolutionRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("5")
          .setBackground('#FFECB3') // 黄色
          .setRanges([sheet.getRange("C9")])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(highResolutionRule, lowResolutionRule);
        sheet.setConditionalFormatRules(rules);
        
        // カテゴリフローセクションの書式設定
        // ナレッジ化優先度
        const highPriorityRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("高")
          .setBackground('#FB8C00') // オレンジ
          .setFontColor('white')
          .setRanges([sheet.getRange("H13:H20")])
          .build();
        
        const mediumPriorityRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("中")
          .setBackground('#FFB74D') // 薄いオレンジ
          .setRanges([sheet.getRange("H13:H20")])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(highPriorityRule, mediumPriorityRule);
        sheet.setConditionalFormatRules(rules);
        
        // アクション列
        const knowledgeActionRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("ナレッジ")
          .setBackground('#C8E6C9') // 緑
          .setRanges([sheet.getRange("I13:I20"), sheet.getRange("G24:G31")])
          .build();
        
        const skillActionRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextContains("スキル")
          .setBackground('#E3F2FD') // 薄い青
          .setRanges([sheet.getRange("I13:I20"), sheet.getRange("G24:G31")])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(knowledgeActionRule, skillActionRule);
        sheet.setConditionalFormatRules(rules);
        
        return true;
      } catch (e) {
        LogModule.handleError(e, '相談フローダッシュボード書式設定');
        return false;
      }
    },
  
    /**
     * 統合パフォーマンスダッシュボードを初期化する
     */
    initIntegratedDashboard: function() {
      LogModule.logOperation("統合パフォーマンスダッシュボード初期化", "開始");
      
      try {
        // 必要なシートの取得
        const masterSheet = UtilsModule.getSheet(SHEETS.MASTER);
        const skillSheet = UtilsModule.getSheet(SHEETS.INDIVIDUAL_SKILL);
        const dataSheet = UtilsModule.getSheet(SHEETS.DATA);
        const integratedDashSheet = UtilsModule.getOrCreateSheet(SHEETS.INTEGRATED_DASHBOARD);
        
        // データの確認
        if (!masterSheet || !skillSheet || !dataSheet) {
          throw new Error('必要なデータシートが見つかりません。先に分析を実行してください。');
        }
        
        // シートをクリア
        integratedDashSheet.clear();
        
        // タイトルと説明の設定
        integratedDashSheet.getRange("A1:J1").merge().setValue("統合パフォーマンスダッシュボード")
          .setFontWeight('bold').setBackground('#E8EAF6').setFontSize(16);
        
        const today = new Date();
        integratedDashSheet.getRange("A2:J2").merge().setValue(`更新日: ${today.toLocaleDateString()}`)
          .setFontStyle("italic");
        
        // 操作方法の説明
        integratedDashSheet.getRange("A3:J3").merge().setValue("使用方法: タイプと対象を選択 → 「スキルマップ→ダッシュボード→統合パフォーマンス更新」を実行")
          .setBackground('#FFF9C4').setFontWeight('bold');
        
        // 分析タイプ選択セクション
        integratedDashSheet.getRange("A5").setValue("分析タイプ:");
        
        // 分析タイプのドロップダウンリスト設定
        const typeValidation = SpreadsheetApp.newDataValidation()
          .requireValueInList(["チームパフォーマンス", "機能領域カバレッジ"], true)
          .build();
        integratedDashSheet.getRange("B5").setDataValidation(typeValidation);
        integratedDashSheet.getRange("B5").setValue("チームパフォーマンス"); // デフォルト値
        
        // チーム/機能領域選択セクション
        integratedDashSheet.getRange("D5").setValue("対象:");
        
        // チーム一覧をマスターから取得
        const teams = UtilsModule.getTeamList(masterSheet);
        
        // 機能領域一覧を取得
        const opTeams = UtilsModule.getOPTeamList(dataSheet);
        
        // チームのドロップダウンリスト設定
        if (teams.length > 0) {
          const teamValidation = SpreadsheetApp.newDataValidation()
            .requireValueInList(teams, true)
            .build();
          integratedDashSheet.getRange("E5").setDataValidation(teamValidation);
          
          // デフォルト値として最初のチームを設定
          if (teams[0]) {
            integratedDashSheet.getRange("E5").setValue(teams[0]);
          }
        }
        
        // 計算モード選択セクション
        integratedDashSheet.getRange("G5").setValue("計算モード:");
        
        // 計算モードのドロップダウンリスト設定
        const modeValidation = SpreadsheetApp.newDataValidation()
          .requireValueInList(["標準", "相談コスト考慮", "実稼働時間基準"], true)
          .build();
        integratedDashSheet.getRange("H5").setDataValidation(modeValidation);
        integratedDashSheet.getRange("H5").setValue("標準"); // デフォルト値
        
        // 対応状況サマリーセクション
        integratedDashSheet.getRange("A7:J7").merge().setValue("対応状況サマリー")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        // サマリー基本情報
        integratedDashSheet.getRange("A8:E8").setValues([[
          "対象名", "対応タグ数", "Rally総数", "チーム内完結率(%)", "効果的生産性"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        integratedDashSheet.getRange("A9:E9").setBackground('#f9f9f9');
        
        // 対応区分セクション
        integratedDashSheet.getRange("A11:J11").merge().setValue("対応区分別パフォーマンス")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        integratedDashSheet.getRange("A12:J12").setValues([[
          "対応区分", "Rally数", "割合(%)", "対応者数", "生産性(Rally/h)", 
          "効果的生産性", "目標生産性", "目標との差", "質問率(%)", "状態"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        integratedDashSheet.getRange("A13:J16").setBackground('#f9f9f9');
        
        // 人員計画セクション
        integratedDashSheet.getRange("A18:H18").merge().setValue("人員計画")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        integratedDashSheet.getRange("A19:H19").setValues([[
          "パラメータ", "現在値", "目標値", "差分", "必要HC", "実際HC", "HC差分", "HC調整"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        // HC計算パラメータ行
        integratedDashSheet.getRange("A20:H25").setValues([
          ["Rally/時間", "", "10.0", "", "", "", "", "▲ ▼"],
          ["チーム内完結率(%)", "", "85.0", "", "", "", "", "▲ ▼"],
          ["相談コスト係数", "", "1.0", "", "", "", "", "▲ ▼"],
          ["リソース効率係数", "", "1.0", "", "", "", "", "▲ ▼"],
          ["必要HC（計算値）", "", "", "", "", "", "", "【計算値】"],
          ["最適HC配分", "", "", "", "", "", "", "【計算値】"]
        ]);
        
        // 対応者パフォーマンスセクション
        integratedDashSheet.getRange("A27:J27").merge().setValue("対応者パフォーマンス（上位8名）")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        integratedDashSheet.getRange("A28:J28").setValues([[
          "担当者名", "IH/PT", "対応Rally数", "生産性(Rally/h)", "効果的生産性", 
          "質問率(%)", "相談対応率(%)", "貢献度(%)", "パフォーマータイプ", "稼働最適度"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        integratedDashSheet.getRange("A29:J36").setBackground('#f9f9f9');
        
        // 改善提案セクション
        integratedDashSheet.getRange("A38:E38").merge().setValue("改善提案")
          .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(14);
        
        integratedDashSheet.getRange("A39:E39").setValues([[
          "提案項目", "詳細", "影響度", "期待効果", "実施難易度"
        ]])
          .setFontWeight('bold').setBackground('#f3f3f3');
        
        integratedDashSheet.getRange("A40:E45").setBackground('#f9f9f9');
        
        // チーム/機能領域変更時の動的変更ハンドラ設定
        integratedDashSheet.getRange("B5").setNote("分析タイプを変更すると、対象リストも更新されます");
        
        // 列幅の調整
        integratedDashSheet.setColumnWidth(1, 150); // 対応区分/担当者名
        integratedDashSheet.setColumnWidth(2, 100); // Rally数/IH/PT
        integratedDashSheet.setColumnWidth(3, 80);  // 割合(%)
        integratedDashSheet.setColumnWidth(5, 120); // 生産性(Rally/h)
        integratedDashSheet.setColumnWidth(6, 120); // 効果的生産性
        integratedDashSheet.setColumnWidth(9, 120); // パフォーマータイプ
        integratedDashSheet.setColumnWidth(10, 120); // 稼働最適度/状態
        
        // シートをアクティブにする
        integratedDashSheet.activate();
        
        LogModule.logOperation("統合パフォーマンスダッシュボード初期化", "完了");
        UtilsModule.showToast('統合パフォーマンスダッシュボードを初期化しました', 'ダッシュボード初期化');
        
        return true;
      } catch (e) {
        LogModule.handleError(e, '統合パフォーマンスダッシュボード初期化');
        return false;
      }
    },
  
    /**
     * 統合パフォーマンスダッシュボードを更新する
     */
    updateIntegratedDashboard: function() {
      LogModule.logOperation("統合パフォーマンスダッシュボード更新", "開始");
      
      try {
        // 必要なシートの取得
        const integratedDashSheet = UtilsModule.getSheet(SHEETS.INTEGRATED_DASHBOARD);
        const teamPerformanceSheet = UtilsModule.getSheet(SHEETS.TEAM_PERFORMANCE);
        const areaCoverageSheet = UtilsModule.getSheet(SHEETS.AREA_COVERAGE);
        const skillSheet = UtilsModule.getSheet(SHEETS.INDIVIDUAL_SKILL);
        const consultationFlowSheet = UtilsModule.getSheet(SHEETS.CONSULTATION_FLOW);
        
        if (!integratedDashSheet || !teamPerformanceSheet || !areaCoverageSheet || !skillSheet) {
          throw new Error('必要なシートが見つかりません。先に分析と初期化を実行してください。');
        }
        
        // 設定取得
        const settings = SettingsModule.getSettings();
        
        // 選択された分析タイプと対象を取得
        const analysisType = integratedDashSheet.getRange("B5").getValue();
        const targetName = integratedDashSheet.getRange("E5").getValue();
        const calculationMode = integratedDashSheet.getRange("H5").getValue();
        
        if (!targetName || !analysisType) {
          integratedDashSheet.getRange("A9").setValue("分析タイプと対象を選択してください");
          return false;
        }
        
        // データ取得元の決定
        let sourceSheet, sourceData;
        if (analysisType === "チームパフォーマンス") {
          sourceSheet = teamPerformanceSheet;
        } else {
          sourceSheet = areaCoverageSheet;
        }
        
        // 対象データの抽出
        const extractedData = this.extractPerformanceData(sourceSheet, targetName, analysisType);
        
        if (!extractedData) {
          integratedDashSheet.getRange("A9").setValue(`${analysisType}データに ${targetName} が見つかりません`);
          return false;
        }
        
        // 相談フローデータの取得（存在する場合）
        const consultationData = consultationFlowSheet ? 
                              this.extractConsultationFlowData(consultationFlowSheet, targetName) : null;
        
        // 個人スキルデータの取得
        const memberSkillData = this.extractMemberSkillData(skillSheet, targetName, analysisType);
        
        // 対応状況サマリー表示
        integratedDashSheet.getRange("A9:E9").setValues([[
          targetName,
          extractedData.totalTags,
          extractedData.totalRallyCount,
          extractedData.completionRate + "%",
          extractedData.effectiveProductivity
        ]]);
        
        // 対応区分別パフォーマンス表示
        const categoryData = [
          ['ihPrimary', 'IH所属チーム'],
          ['ihSecondary', 'IH他チーム'],
          ['ptPrimary', 'PT所属チーム'],
          ['ptSecondary', 'PT他チーム']
        ].map(([category, displayName], index) => {
          const responses = extractedData[`${category}Responses`] || {};
          
          // 目標値の取得
          const targetProd = settings[`${category}TargetProd`] || 10;
          const targetRate = settings[`${category}TargetRate`] || 25;
          
          // 目標との差分
          const prodGap = (responses.productivity || 0) - targetProd;
          
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
            displayName,
            responses.rallyCount || 0,
            (responses.rallyRate || 0).toFixed(1) + "%",
            responses.responders ? responses.responders.size : 0,
            (responses.productivity || 0).toFixed(1),
            (responses.effectiveProductivity || 0).toFixed(1),
            targetProd.toFixed(1),
            prodGap.toFixed(1),
            (responses.questionRate || 0).toFixed(1) + "%",
            status
          ];
        });
        
        integratedDashSheet.getRange(13, 1, 4, 10).setValues(categoryData);
        
        // 人員計画セクション更新
        // 現在値の表示
        integratedDashSheet.getRange("B20").setValue(extractedData.totalProductivity.toFixed(1));
        integratedDashSheet.getRange("B21").setValue(extractedData.completionRate.toFixed(1));
        
        const consultationCostFactor = consultationData ? 
                                    (1 + consultationData.avgSolveDuration / 60 * 0.5) : 1.0;
        
        integratedDashSheet.getRange("B22").setValue(consultationCostFactor.toFixed(2));
        
        // リソース効率の計算（パフォーマータイプ分布から）
        const resourceEfficiency = this.calculateResourceEfficiency(memberSkillData);
        integratedDashSheet.getRange("B23").setValue(resourceEfficiency.toFixed(2));
        
        // 目標値の取得（ユーザー入力値またはデフォルト）
        const targetRallyPerHour = integratedDashSheet.getRange("C20").getValue() || 10.0;
        const targetCompletionRate = integratedDashSheet.getRange("C21").getValue() || 85.0;
        const targetCostFactor = integratedDashSheet.getRange("C22").getValue() || 1.0;
        const targetEfficiency = integratedDashSheet.getRange("C23").getValue() || 1.0;
        
        // 差分計算
        integratedDashSheet.getRange("D20").setValue((extractedData.totalProductivity - targetRallyPerHour).toFixed(1));
        integratedDashSheet.getRange("D21").setValue((extractedData.completionRate - targetCompletionRate).toFixed(1));
        integratedDashSheet.getRange("D22").setValue((consultationCostFactor - targetCostFactor).toFixed(2));
        integratedDashSheet.getRange("D23").setValue((resourceEfficiency - targetEfficiency).toFixed(2));
        
        // 必要HC計算
        const monthlyWorkHours = settings.workDaysPerMonth * settings.hoursPerDay * settings.productivityRatio;
        
        // 月間Rally数の推定（3ヶ月分のデータを元に）
        const monthlyRallyCount = extractedData.totalRallyCount / 3;
        
        // 目標Rally生産性に基づく必要HC計算
        const targetRally = monthlyRallyCount * (targetCompletionRate / 100);
        
        const baseIdealHC = (targetRally / (targetRallyPerHour * monthlyWorkHours));
        
        // 調整係数を適用した理想HC
        let adjustedIdealHC;
        
        switch (calculationMode) {
          case "相談コスト考慮":
            adjustedIdealHC = baseIdealHC * targetCostFactor;
            break;
          case "実稼働時間基準":
            adjustedIdealHC = baseIdealHC / targetEfficiency;
            break;
          default: // 標準モード
            adjustedIdealHC = baseIdealHC;
        }
        
        // 必要HC表示
        integratedDashSheet.getRange("E24").setValue(adjustedIdealHC.toFixed(2));
        
        // 実際のHC計算
        const actualHC = extractedData.totalHcRatio || 0;
        integratedDashSheet.getRange("F24").setValue(actualHC.toFixed(2));
        
        // HC差分
        const hcGap = actualHC - adjustedIdealHC;
        integratedDashSheet.getRange("G24").setValue(hcGap.toFixed(2));
        
        // HC配分の詳細計算
        const allocations = this.calculateHCAllocations(
          extractedData, 
          {
            targetRallyPerHour,
            targetCompletionRate,
            baseIdealHC,
            adjustedIdealHC
          }
        );
        
        integratedDashSheet.getRange("F25").setValue(actualHC.toFixed(2));
        integratedDashSheet.getRange("E25").setValue(adjustedIdealHC.toFixed(2));
        integratedDashSheet.getRange("G25").setValue(hcGap.toFixed(2));
        integratedDashSheet.getRange("A25").setValue("区分別必要HC");
        integratedDashSheet.getRange("C25").setValue(allocations.map(a => 
          `${a.category}: ${a.idealHC.toFixed(1)}`
        ).join(", "));
        
        // 対応者パフォーマンス表示
        if (memberSkillData && memberSkillData.length > 0) {
          const memberCount = Math.min(memberSkillData.length, 8);
          
          const memberRows = memberSkillData.slice(0, memberCount).map(member => {
            // 稼働最適度の計算
            const optimalityScore = this.calculateOptimalityScore(member, extractedData);
            let optimalityStatus;
            
            if (optimalityScore >= 0.9) {
              optimalityStatus = "最適";
            } else if (optimalityScore >= 0.7) {
              optimalityStatus = "良好";
            } else if (optimalityScore >= 0.5) {
              optimalityStatus = "調整検討";
            } else {
              optimalityStatus = "再配置検討";
            }
            
            return [
              member.name,
              member.ihPt,
              member.rallyCount,
              member.productivity.toFixed(1),
              member.effectiveProductivity.toFixed(1),
              member.questionRate.toFixed(1) + "%",
              member.adviceRate.toFixed(1) + "%",
              member.contributionRate.toFixed(1) + "%",
              member.performerType,
              optimalityStatus
            ];
          });
          
          integratedDashSheet.getRange(29, 1, memberCount, 10).setValues(memberRows);
          
          // 空白行のクリア
          if (memberCount < 8) {
            integratedDashSheet.getRange(29 + memberCount, 1, 8 - memberCount, 10).clearContent();
          }
        } else {
          integratedDashSheet.getRange("A29:J29").merge()
            .setValue(`${targetName} に関連する対応者データがありません`)
            .setFontStyle("italic");
          
          integratedDashSheet.getRange(30, 1, 7, 10).clearContent();
        }
        
        // 改善提案の生成と表示
        const proposals = this.generateImprovedProposals(
          extractedData, 
          consultationData, 
          memberSkillData, 
          analysisType, 
          targetName,
          {
            hcGap,
            resourceEfficiency
          }
        );
        
        if (proposals.length > 0) {
          const proposalCount = Math.min(proposals.length, 6);
          
          integratedDashSheet.getRange(40, 1, proposalCount, 5).setValues(
            proposals.slice(0, proposalCount)
          );
          
          // 空白行のクリア
          if (proposalCount < 6) {
            integratedDashSheet.getRange(40 + proposalCount, 1, 6 - proposalCount, 5).clearContent();
          }
        } else {
          integratedDashSheet.getRange("A40:E40").merge()
            .setValue("提案を生成するためのデータが不足しています")
            .setFontStyle("italic");
          
          integratedDashSheet.getRange(41, 1, 5, 5).clearContent();
        }
        
      // 書式設定
        this.setIntegratedDashboardFormatting(integratedDashSheet);
        
        LogModule.logOperation("統合パフォーマンスダッシュボード更新", `完了: "${targetName}" の分析`);
        UtilsModule.showToast(`${targetName} の統合パフォーマンスダッシュボードが更新されました`, '統合ダッシュボード');
        
        return true;
      } catch (e) {
        LogModule.handleError(e, '統合パフォーマンスダッシュボード更新');
        return false;
      }
    },
  
    /**
     * パフォーマンスデータの抽出
     */
    extractPerformanceData: function(sheet, targetName, analysisType) {
      try {
        const data = sheet.getDataRange().getValues();
        
        // 対象行を探す
        let targetRow = -1;
        for (let i = 5; i < data.length; i++) {
          if (data[i][0] === targetName) {
            targetRow = i;
            break;
          }
        }
        
        if (targetRow === -1) {
          return null; // 対象が見つからない
        }
        
        // 基本データの抽出（チームパフォーマンスと機能領域カバレッジで共通の部分）
        const basicData = {
          targetName: targetName,
          totalTags: data[targetRow][1],
          totalRallyCount: data[targetRow][3], // チームパフォーマンスデータの場合
          completionRate: parseFloat(data[targetRow][data[0].indexOf("チーム内完結率(%)") || 3]),
          totalProductivity: 0,
          effectiveProductivity: 0,
          
          // 対応区分データ（初期化）
          ihPrimaryResponses: { rallyCount: 0, rallyRate: 0, responders: new Set(), productivity: 0, effectiveProductivity: 0, questionRate: 0 },
          ihSecondaryResponses: { rallyCount: 0, rallyRate: 0, responders: new Set(), productivity: 0, effectiveProductivity: 0, questionRate: 0 },
          ptPrimaryResponses: { rallyCount: 0, rallyRate: 0, responders: new Set(), productivity: 0, effectiveProductivity: 0, questionRate: 0 },
          ptSecondaryResponses: { rallyCount: 0, rallyRate: 0, responders: new Set(), productivity: 0, effectiveProductivity: 0, questionRate: 0 }
        };
        
        // 詳細データの検索
        // チーム詳細セクションを探す
        let detailHeaderRow = -1;
        for (let i = targetRow + 1; i < Math.min(data.length, targetRow + 20); i++) {
          if (data[i][0] && data[i][0].includes(targetName) && data[i][0].includes("詳細")) {
            detailHeaderRow = i;
            break;
          }
        }
        
        // 対応区分別セクションを探す
        let categoryHeaderRow = -1;
        if (detailHeaderRow !== -1) {
          for (let i = detailHeaderRow + 1; i < Math.min(data.length, detailHeaderRow + 10); i++) {
            if (data[i][0] === "対応区分" || (data[i][0] === "カテゴリ" && data[i][1] === "Rally数")) {
              categoryHeaderRow = i;
              break;
            }
          }
        }
        
        // 対応区分データの抽出
        if (categoryHeaderRow !== -1) {
          const categories = [
            { name: "IH所属チーム", key: "ihPrimary" },
            { name: "IH他チーム", key: "ihSecondary" },
            { name: "PT所属チーム", key: "ptPrimary" },
            { name: "PT他チーム", key: "ptSecondary" }
          ];
          
          // 各カテゴリのデータを探して抽出
          for (let i = categoryHeaderRow + 1; i < Math.min(data.length, categoryHeaderRow + 10); i++) {
            if (!data[i][0]) break; // 空の行に達したら終了
            
            // カテゴリ名からキーを特定
            const category = categories.find(c => data[i][0].includes(c.name));
            if (category) {
              basicData[`${category.key}Responses`] = {
                rallyCount: data[i][1],
                rallyRate: parseFloat(data[i][2]),
                responders: { size: data[i][3] }, // セットではなく数値のみ
                productivity: parseFloat(data[i][4]),
                effectiveProductivity: parseFloat(data[i][5] || data[i][4]), // 効果的生産性がなければ通常生産性
                questionRate: parseFloat(data[i][8] || 0)
              };
            }
          }
        }
        
        // 全体生産性データの抽出（列のインデックスが異なる可能性に対応）
        let prodColIdx = -1;
        for (let i = 0; i < data[targetRow].length; i++) {
          if (data[4][i] && (data[4][i].includes("生産性") || data[4][i].includes("生産性(Rally/h)"))) {
            prodColIdx = i;
            break;
          }
        }
        
        if (prodColIdx !== -1) {
          basicData.totalProductivity = parseFloat(data[targetRow][prodColIdx]);
        }
        
        // 効果的生産性データの抽出
        let effectiveProdColIdx = -1;
        for (let i = 0; i < data[targetRow].length; i++) {
          if (data[4][i] && data[4][i].includes("効果的生産性")) {
            effectiveProdColIdx = i;
            break;
          }
        }
        
        if (effectiveProdColIdx !== -1) {
          basicData.effectiveProductivity = parseFloat(data[targetRow][effectiveProdColIdx]);
        } else {
          basicData.effectiveProductivity = basicData.totalProductivity;
        }
        
        // HC関連データの抽出
        let totalHcRatio = 0;
        
        if (analysisType === "チームパフォーマンス") {
          for (let i = 0; i < data[targetRow].length; i++) {
            if (data[4][i] && data[4][i].includes("実質HC合計")) {
              totalHcRatio = parseFloat(data[targetRow][i]);
              break;
            }
          }
        }
        
        basicData.totalHcRatio = totalHcRatio;
        
        return basicData;
      } catch (e) {
        LogModule.handleError(e, 'パフォーマンスデータ抽出');
        return null;
      }
    },
  
    /**
     * リソース効率係数を計算する
     */
    calculateResourceEfficiency: function(memberSkillData) {
      if (!memberSkillData || memberSkillData.length === 0) return 1.0;
      
      // パフォーマータイプによる効率係数
      const typeEfficiencies = {
        'ハイパフォーマー': 1.2,
        'スペシャリスト': 1.1,
        'オールラウンダー': 1.0,
        '標準': 0.9,
        '成長中': 0.7
      };
      
      // 各メンバーの効率を集計
      let totalEfficiency = 0;
      let totalWeight = 0;
      
      memberSkillData.forEach(member => {
        const efficiency = typeEfficiencies[member.performerType] || 0.9;
        const weight = member.contributionRate || 1;
        
        totalEfficiency += efficiency * weight;
        totalWeight += weight;
      });
      
      // 平均効率を計算（貢献度で重み付け）
      return totalWeight > 0 ? totalEfficiency / totalWeight : 1.0;
    },
  
    /**
     * 区分別HC配分を計算する
     */
    calculateHCAllocations: function(extractedData, params) {
      const categories = [
        { key: 'ihPrimary', name: 'IH所属チーム', targetRate: params.targetCompletionRate },
        { key: 'ihSecondary', name: 'IH他チーム', targetRate: 5 },
        { key: 'ptPrimary', name: 'PT所属チーム', targetRate: 5 },
        { key: 'ptSecondary', name: 'PT他チーム', targetRate: 5 }
      ];
      
      return categories.map(category => {
        const responses = extractedData[`${category.key}Responses`] || {};
        const currentRate = responses.rallyRate || 0;
        
        // 対応率が目標に達していない場合、必要HCを計算
        const rateRatio = category.targetRate / Math.max(currentRate, 0.1); // ゼロ除算防止
        const categoryIdealHC = params.adjustedIdealHC * (category.targetRate / 100);
        
        return {
          category: category.name,
          currentRate: currentRate,
          targetRate: category.targetRate,
          rateRatio: rateRatio,
          idealHC: categoryIdealHC
        };
      });
    },
  
    /**
     * 担当者の稼働最適度を計算する
     */
    calculateOptimalityScore: function(member, teamData) {
      // 稼働最適度の計算要素
      // 1. 生産性（目標との比較）
      const productivityOptimality = Math.min(1.0, member.productivity / 10);
      
      // 2. 質問率（低いほど良い）
      const questionOptimality = 1.0 - Math.min(0.5, member.questionRate / 100);
      
      // 3. 相談対応率（役割に応じて評価）
      let adviceOptimality = 0.5;
      if (member.performerType === 'スペシャリスト' || member.performerType === 'ハイパフォーマー') {
        // スペシャリストとハイパフォーマーは相談対応率が高いほど良い
        adviceOptimality = Math.min(1.0, member.adviceRate / 20);
      } else {
        // それ以外は中程度が理想
        adviceOptimality = 1.0 - Math.abs(member.adviceRate - 10) / 20;
      }
      
      // 4. パフォーマータイプの適合度
      const typeWeights = {
        'ハイパフォーマー': 1.0,
        'スペシャリスト': 0.9,
        'オールラウンダー': 0.8,
        '標準': 0.7,
        '成長中': 0.6
      };
      
      const typeOptimality = typeWeights[member.performerType] || 0.7;
      
      // 5. チーム貢献度（適切な範囲内か）
      let contributionOptimality = 0.5;
      if (member.ihPt === 'IH') {
        // IHメンバーは貢献度10-30%が適切
        if (member.contributionRate >= 10 && member.contributionRate <= 30) {
          contributionOptimality = 1.0;
        } else if (member.contributionRate < 10) {
          contributionOptimality = member.contributionRate / 10;
        } else { // > 30%
          contributionOptimality = 1.0 - Math.min(0.5, (member.contributionRate - 30) / 70);
        }
      } else {
        // PTメンバーは貢献度5-15%が適切
        if (member.contributionRate >= 5 && member.contributionRate <= 15) {
          contributionOptimality = 1.0;
        } else if (member.contributionRate < 5) {
          contributionOptimality = member.contributionRate / 5;
        } else { // > 15%
          contributionOptimality = 1.0 - Math.min(0.5, (member.contributionRate - 15) / 85);
        }
      }
      
      // 総合評価（重み付き平均）
      return productivityOptimality * 0.3 + 
            questionOptimality * 0.1 + 
            adviceOptimality * 0.2 + 
            typeOptimality * 0.2 + 
            contributionOptimality * 0.2;
    },
  
    /**
     * 改善提案を生成する
     */
    generateImprovedProposals: function(extractedData, consultationData, memberSkillData, analysisType, targetName, params) {
      const proposals = [];
      
      // 提案1: HC調整
      if (params.hcGap < -0.5) {
        // HC不足
        proposals.push([
          "リソース増強",
          `現在のHCは必要量より${Math.abs(params.hcGap).toFixed(1)}人分不足しています。特にIH所属チームのリソース強化を検討してください。`,
          "高",
          "必要HC確保、チーム内完結率向上",
          "中"
        ]);
      } else if (params.hcGap > 0.5) {
        // HC余剰
        proposals.push([
          "リソース最適化",
          `現在のHCは必要量より${params.hcGap.toFixed(1)}人分多い状態です。他チームへの応援やスキル拡張を検討してください。`,
          "中",
          "コスト削減、リソース有効活用",
          "低"
        ]);
      }
      
      // 提案2: スキル分布最適化
      if (memberSkillData && memberSkillData.length > 0) {
        const performerTypes = memberSkillData.reduce((acc, member) => {
          acc[member.performerType] = (acc[member.performerType] || 0) + 1;
          return acc;
        }, {});
        
        const totalMembers = memberSkillData.length;
        
        // スペシャリスト比率の確認
        const specialistRatio = ((performerTypes['スペシャリスト'] || 0) / totalMembers) * 100;
        
        if (specialistRatio < 15) {
          proposals.push([
            "スペシャリスト育成",
            `スペシャリスト比率が低い (${specialistRatio.toFixed(1)}%)。特定領域の専門家を育成してください。`,
            "中",
            "複雑案件対応力強化、チーム知識レベル向上",
            "高"
          ]);
        }
        
        // 成長中メンバー比率の確認
        const growingRatio = ((performerTypes['成長中'] || 0) / totalMembers) * 100;
        
        if (growingRatio > 30) {
          proposals.push([
            "育成プログラム強化",
            `成長中メンバーの比率が高い (${growingRatio.toFixed(1)}%)。体系的な育成プログラムとメンター制度の導入を検討してください。`,
            "高",
            "新人立ち上げ期間短縮、チーム全体の生産性向上",
            "中"
          ]);
        }
      }
      
      // 提案3: 生産性向上
      const avgProductivity = extractedData.totalProductivity || 0;
      
      if (avgProductivity < 8) {
        proposals.push([
          "生産性向上施策",
          `チーム生産性が目標を下回っています (${avgProductivity.toFixed(1)} Rally/h)。マクロ活用促進とナレッジ共有の強化を検討してください。`,
          "高",
          "必要HC削減、対応スピード向上",
          "中"
        ]);
      }
      
      // 提案4: 相談コスト最適化
      if (consultationData && consultationData.avgSolveDuration > 10) {
        proposals.push([
          "相談プロセス効率化",
          `相談解決時間が長い (平均${consultationData.avgSolveDuration.toFixed(1)}分)。FAQ強化と相談テンプレート導入を検討してください。`,
          "中",
          "相談コスト削減、対応時間短縮",
          "低"
        ]);
      }
      
      // 提案5: リソース効率向上
      if (params.resourceEfficiency < 0.9) {
        proposals.push([
          "メンバー配置最適化",
          `リソース効率が低い (${params.resourceEfficiency.toFixed(2)})。各メンバーの強みを活かした業務配分の見直しを行ってください。`,
          "中",
          "同じHCでより高いパフォーマンス実現",
          "中"
        ]);
      }
      
      // 提案6: チーム完結率向上
      const completionRate = extractedData.completionRate || 0;
      
      if (completionRate < 70) {
        proposals.push([
          "チーム内完結率向上",
          `チーム内完結率が低い (${completionRate.toFixed(1)}%)。コアタグのスキル移管と研修強化を検討してください。`,
          "高",
          "外部依存度低減、対応スピード向上",
          "高"
        ]);
      }
      
      return proposals;
    },
  
    /**
     * 個人スキルデータを抽出する
     */
    extractMemberSkillData: function(sheet, targetName, analysisType) {
      try {
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        
        const nameIdx = headers.indexOf('担当者名');
        const ihPtIdx = headers.indexOf('IH/PT区分');
        const primaryTeamIdx = headers.indexOf('主所属OPチーム');
        const secondaryTeamIdx = headers.indexOf('兼務OPチーム');
        const rallyCountIdx = -1; // 個人スキルシートにはRally数がない可能性
        const productivityIdx = headers.indexOf('生産性スコア');
        const questionRateIdx = headers.indexOf('質問率(%)');
        const adviceRateIdx = headers.indexOf('相談対応率(%)');
        const performerTypeIdx = headers.indexOf('パフォーマータイプ');
        
        if (nameIdx === -1 || primaryTeamIdx === -1) {
          return [];
        }
        
        const memberData = [];
        
        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          let isTargetMember = false;
          
          if (analysisType === "チームパフォーマンス") {
            // チーム分析の場合、該当チームに所属または兼務しているメンバーを抽出
            isTargetMember = (row[primaryTeamIdx] === targetName);
            
            if (!isTargetMember && secondaryTeamIdx !== -1 && row[secondaryTeamIdx]) {
              isTargetMember = row[secondaryTeamIdx].split(',').some(team => team.trim() === targetName);
            }
          } else {
            // 機能領域分析の場合、該当領域のタグを担当しているメンバーを抽出
            // （この判定はデータからは難しいため、一旦すべてのメンバーを対象とする）
            isTargetMember = true;
          }
          
          if (isTargetMember) {
            memberData.push({
              name: row[nameIdx],
              ihPt: row[ihPtIdx],
              primaryTeam: row[primaryTeamIdx],
              secondaryTeam: secondaryTeamIdx !== -1 ? row[secondaryTeamIdx] : "",
              rallyCount: rallyCountIdx !== -1 ? row[rallyCountIdx] : 0,
              productivity: productivityIdx !== -1 ? parseFloat(row[productivityIdx]) : 0,
              questionRate: questionRateIdx !== -1 ? parseFloat(row[questionRateIdx]) : 0,
              adviceRate: adviceRateIdx !== -1 ? parseFloat(row[adviceRateIdx]) : 0,
              performerType: performerTypeIdx !== -1 ? row[performerTypeIdx] : "標準",
              contributionRate: 100 / data.length // 仮の値、後で調整
            });
          }
        }
        
        return memberData;
      } catch (e) {
        LogModule.handleError(e, '個人スキルデータ抽出');
        return [];
      }
    },
  
    /**
     * 統合ダッシュボードの書式を設定する
     */
    setIntegratedDashboardFormatting: function(sheet) {
      try {
        // 対応状況サマリーの書式設定
        // チーム内完結率
        const completionRateRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(80)
          .setBackground('#C8E6C9') // 緑
          .setRanges([sheet.getRange("D9")])
          .build();
        
        let rules = sheet.getConditionalFormatRules();
        rules.push(completionRateRule);
        sheet.setConditionalFormatRules(rules);
        
        // 生産性
        const productivityRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(10)
          .setBackground('#C8E6C9') // 緑
          .setRanges([sheet.getRange("E9")])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(productivityRule);
        sheet.setConditionalFormatRules(rules);
        
        // 対応区分別パフォーマンスの書式設定
        // 状態列
        const statusRange = sheet.getRange("J13:J16");
        
        const optimalStatusRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("適正")
          .setBackground('#C8E6C9') // 緑
          .setRanges([statusRange])
          .build();
        
        const surplusStatusRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("余裕あり")
          .setBackground('#E3F2FD') // 青
          .setRanges([statusRange])
          .build();
        
        const shortageStatusRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("不足")
          .setBackground('#FFEBEE') // 赤
          .setRanges([statusRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(optimalStatusRule, surplusStatusRule, shortageStatusRule);
        sheet.setConditionalFormatRules(rules);
        
        // 人員計画セクションの書式設定
        // HC差分
        const hcGapRange = sheet.getRange("G24:G25");
        
        const positiveHcRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberGreaterThan(0.5)
          .setBackground('#C8E6C9') // 緑
          .setRanges([hcGapRange])
          .build();
        
        const negativeHcRule = SpreadsheetApp.newConditionalFormatRule()
          .whenNumberLessThan(-0.5)
          .setBackground('#FFEBEE') // 赤
          .setRanges([hcGapRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(positiveHcRule, negativeHcRule);
        sheet.setConditionalFormatRules(rules);
        
        // 対応者パフォーマンスセクションの書式設定
        // パフォーマータイプ
        const typeRange = sheet.getRange("I29:I36");
        
        const typeColors = [
          { type: "ハイパフォーマー", color: "#1565C0", textColor: "white" }, // 濃い青
          { type: "スペシャリスト", color: "#6A1B9A", textColor: "white" }, // 紫
          { type: "オールラウンダー", color: "#2E7D32", textColor: "white" }, // 緑
          { type: "標準", color: "#757575", textColor: "white" }, // グレー
          { type: "成長中", color: "#FF9800", textColor: "black" } // オレンジ
        ];
        
        typeColors.forEach(typeInfo => {
          const typeRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(typeInfo.type)
            .setBackground(typeInfo.color)
            .setFontColor(typeInfo.textColor)
            .setRanges([typeRange])
            .build();
          
          rules = sheet.getConditionalFormatRules();
          rules.push(typeRule);
          sheet.setConditionalFormatRules(rules);
        });
        
        // 稼働最適度
        const optimalityRange = sheet.getRange("J29:J36");
        
        const optimalRules = [
          { status: "最適", color: "#43A047" }, // 緑
          { status: "良好", color: "#7CB342" }, // 薄い緑
          { status: "調整検討", color: "#FFB74D" }, // オレンジ
          { status: "再配置検討", color: "#EF5350" } // 赤
        ];
        
        optimalRules.forEach(rule => {
          const optRule = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(rule.status)
            .setBackground(rule.color)
            .setFontColor("white")
            .setRanges([optimalityRange])
            .build();
          
          rules = sheet.getConditionalFormatRules();
          rules.push(optRule);
          sheet.setConditionalFormatRules(rules);
        });
        
        // 改善提案セクションの書式設定
        // 影響度
        const impactRange = sheet.getRange("C40:C45");
        
        const highImpactRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("高")
          .setBackground('#EF5350') // 赤
          .setFontColor("white")
          .setRanges([impactRange])
          .build();
        
        const mediumImpactRule = SpreadsheetApp.newConditionalFormatRule()
          .whenTextEqualTo("中")
          .setBackground('#FFB74D') // オレンジ
          .setRanges([impactRange])
          .build();
        
        rules = sheet.getConditionalFormatRules();
        rules.push(highImpactRule, mediumImpactRule);
        sheet.setConditionalFormatRules(rules);
        
        return true;
      } catch (e) {
        LogModule.handleError(e, '統合ダッシュボード書式設定');
        return false;
      }
    }
  };