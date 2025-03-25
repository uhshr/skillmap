// 設定管理モジュール
const SettingsModule = {
  // デフォルト設定値
  DEFAULT_SETTINGS: {
    // 基本設定
    ihDefaultHours: 4.5,     // IHデフォルト稼働時間
    ptDefaultHours: 5.5,     // PTデフォルト稼働時間
    coreSkillThreshold: 20,  // コアスキル閾値(%)
    rareCaseThreshold: 5,    // レアケース閾値(%)
    minCaseCount: 5,         // 最小ケース数
    
    // 難易度閾値（技術的複雑性）
    tcL1Threshold: 1.8,      // 技術的複雑性L1上限
    tcL2Threshold: 2.6,      // 技術的複雑性L2上限
    tcL3Threshold: 3.4,      // 技術的複雑性L3上限
    tcL4Threshold: 4.2,      // 技術的複雑性L4上限
    
    // 難易度閾値（解決難易度）
    rdL1Threshold: 1.8,      // 解決難易度L1上限
    rdL2Threshold: 2.6,      // 解決難易度L2上限
    rdL3Threshold: 3.4,      // 解決難易度L3上限
    rdL4Threshold: 4.2,      // 解決難易度L4上限
    
    // 難易度閾値（知識要求度）
    krL1Threshold: 1.8,      // 知識要求度L1上限
    krL2Threshold: 2.6,      // 知識要求度L2上限
    krL3Threshold: 3.4,      // 知識要求度L3上限
    krL4Threshold: 4.2,      // 知識要求度L4上限
    
    // パフォーマータイプ判定
    hyperperformerWidthThreshold: 75, // ハイパフォーマー幅閾値(%)
    hyperperformerProdThreshold: 10,  // ハイパフォーマー生産性閾値
    specialistDepthThreshold: 3.5,    // スペシャリスト深さ閾値
    specialistComplexThreshold: 3,    // スペシャリスト複雑ケース閾値
    allrounderWidthThreshold: 70,     // オールラウンダー幅閾値(%)
    
    // HC計算パラメータ
    workDaysPerMonth: 20,             // 月間稼働日数
    hoursPerDay: 7.5,                 // 1日あたり稼働時間
    productivityRatio: 0.7,           // 生産性係数
    targetRallyPerHour: 10,           // 目標Rally/時間
    targetCompletionRate: 85,         // 目標チーム内完結率(%)
    
    // 新規追加：対応区分ごとの目標値
    ihPrimaryTargetRate: 70,          // IH所属チーム対応目標(%)
    ihSecondaryTargetRate: 15,        // IH他チーム対応目標(%)
    ptPrimaryTargetRate: 10,          // PT所属チーム対応目標(%)
    ptSecondaryTargetRate: 5,         // PT他チーム対応目標(%)
    
    ihPrimaryTargetProd: 12,          // IH所属チーム生産性目標
    ihSecondaryTargetProd: 10,        // IH他チーム生産性目標
    ptPrimaryTargetProd: 8,           // PT所属チーム生産性目標
    ptSecondaryTargetProd: 7,         // PT他チーム生産性目標
    
    // 新規追加：相談コスト係数
    consultationCostFactor: 0.5,      // 相談コスト影響係数
    questionCostFactor: 0.3           // 質問コスト影響係数
  },

  /**
   * 設定シートを作成する
   */
  createSettingsSheet: function() {
    try {
      const sheet = UtilsModule.getOrCreateSheet(SHEETS.SETTINGS);
      
      // ヘッダー設定
      sheet.getRange("A1:C1").setValues([["設定名", "設定値", "説明"]])
        .setFontWeight('bold').setBackground('#f3f3f3');
      
      // 設定カテゴリと値を構造化
      const settingsData = [
        // 基本設定
        {
          category: "基本設定", 
          settings: [
            ["ihDefaultHours", this.DEFAULT_SETTINGS.ihDefaultHours, "社内メンバーの標準稼働時間（デフォルト4.5時間）- 担当者マスターで個別調整可能"],
            ["ptDefaultHours", this.DEFAULT_SETTINGS.ptDefaultHours, "社外メンバーの標準稼働時間（デフォルト5.5時間）- 担当者マスターで個別調整可能"],
            ["coreSkillThreshold", this.DEFAULT_SETTINGS.coreSkillThreshold, "コアスキルと判定する頻度閾値（%）"],
            ["rareCaseThreshold", this.DEFAULT_SETTINGS.rareCaseThreshold, "レアケースと判定する頻度閾値（%）"],
            ["minCaseCount", this.DEFAULT_SETTINGS.minCaseCount, "信頼性のある分析に必要な最小ケース数"]
          ]
        },
        
        // 難易度閾値（技術的複雑性）
        {
          category: "技術的複雑性の閾値", 
          settings: [
            ["tcL1Threshold", this.DEFAULT_SETTINGS.tcL1Threshold, "技術的複雑性L1上限（この値未満はL1）"],
            ["tcL2Threshold", this.DEFAULT_SETTINGS.tcL2Threshold, "技術的複雑性L2上限（この値未満はL2）"],
            ["tcL3Threshold", this.DEFAULT_SETTINGS.tcL3Threshold, "技術的複雑性L3上限（この値未満はL3）"],
            ["tcL4Threshold", this.DEFAULT_SETTINGS.tcL4Threshold, "技術的複雑性L4上限（この値未満はL4、以上はL5）"]
          ]
        },
        
        // 難易度閾値（解決難易度）
        {
          category: "解決難易度の閾値", 
          settings: [
            ["rdL1Threshold", this.DEFAULT_SETTINGS.rdL1Threshold, "解決難易度L1上限（この値未満はL1）"],
            ["rdL2Threshold", this.DEFAULT_SETTINGS.rdL2Threshold, "解決難易度L2上限（この値未満はL2）"],
            ["rdL3Threshold", this.DEFAULT_SETTINGS.rdL3Threshold, "解決難易度L3上限（この値未満はL3）"],
            ["rdL4Threshold", this.DEFAULT_SETTINGS.rdL4Threshold, "解決難易度L4上限（この値未満はL4、以上はL5）"]
          ]
        },
        
        // 難易度閾値（知識要求度）
        {
          category: "知識要求度の閾値", 
          settings: [
            ["krL1Threshold", this.DEFAULT_SETTINGS.krL1Threshold, "知識要求度L1上限（この値未満はL1）"],
            ["krL2Threshold", this.DEFAULT_SETTINGS.krL2Threshold, "知識要求度L2上限（この値未満はL2）"],
            ["krL3Threshold", this.DEFAULT_SETTINGS.krL3Threshold, "知識要求度L3上限（この値未満はL3）"],
            ["krL4Threshold", this.DEFAULT_SETTINGS.krL4Threshold, "知識要求度L4上限（この値未満はL4、以上はL5）"]
          ]
        },
        
        // パフォーマータイプ判定
        {
          category: "パフォーマータイプ判定", 
          settings: [
            ["hyperperformerWidthThreshold", this.DEFAULT_SETTINGS.hyperperformerWidthThreshold, "ハイパフォーマー幅閾値(%) - これ以上で幅広いと判定"],
            ["hyperperformerProdThreshold", this.DEFAULT_SETTINGS.hyperperformerProdThreshold, "ハイパフォーマー生産性閾値 - これ以上で高生産性と判定"],
            ["specialistDepthThreshold", this.DEFAULT_SETTINGS.specialistDepthThreshold, "スペシャリスト深さ閾値 - これ以上で専門性が高いと判定"],
            ["specialistComplexThreshold", this.DEFAULT_SETTINGS.specialistComplexThreshold, "スペシャリスト複雑ケース閾値 - これ以上の複雑案件対応でスペシャリスト判定"],
            ["allrounderWidthThreshold", this.DEFAULT_SETTINGS.allrounderWidthThreshold, "オールラウンダー幅閾値(%) - これ以上で幅広いと判定"]
          ]
        },
        
        // HC計算パラメータ
        {
          category: "HC計算パラメータ", 
          settings: [
            ["workDaysPerMonth", this.DEFAULT_SETTINGS.workDaysPerMonth, "月間稼働日数"],
            ["hoursPerDay", this.DEFAULT_SETTINGS.hoursPerDay, "1日あたり稼働時間"],
            ["productivityRatio", this.DEFAULT_SETTINGS.productivityRatio, "生産性係数（0-1）"],
            ["targetRallyPerHour", this.DEFAULT_SETTINGS.targetRallyPerHour, "目標Rally/時間"],
            ["targetCompletionRate", this.DEFAULT_SETTINGS.targetCompletionRate, "目標チーム内完結率(%)"]
          ]
        },
        
        // 新規追加：対応区分ごとの目標値
        {
          category: "対応区分ごとの目標値", 
          settings: [
            ["ihPrimaryTargetRate", this.DEFAULT_SETTINGS.ihPrimaryTargetRate, "IH所属チーム対応目標(%)"],
            ["ihSecondaryTargetRate", this.DEFAULT_SETTINGS.ihSecondaryTargetRate, "IH他チーム対応目標(%)"],
            ["ptPrimaryTargetRate", this.DEFAULT_SETTINGS.ptPrimaryTargetRate, "PT所属チーム対応目標(%)"],
            ["ptSecondaryTargetRate", this.DEFAULT_SETTINGS.ptSecondaryTargetRate, "PT他チーム対応目標(%)"],
            ["ihPrimaryTargetProd", this.DEFAULT_SETTINGS.ihPrimaryTargetProd, "IH所属チーム生産性目標"],
            ["ihSecondaryTargetProd", this.DEFAULT_SETTINGS.ihSecondaryTargetProd, "IH他チーム生産性目標"],
            ["ptPrimaryTargetProd", this.DEFAULT_SETTINGS.ptPrimaryTargetProd, "PT所属チーム生産性目標"],
            ["ptSecondaryTargetProd", this.DEFAULT_SETTINGS.ptSecondaryTargetProd, "PT他チーム生産性目標"]
          ]
        },
        
        // 新規追加：相談コスト係数
        {
          category: "相談コスト係数", 
          settings: [
            ["consultationCostFactor", this.DEFAULT_SETTINGS.consultationCostFactor, "相談コスト影響係数 - 相談対応コストの影響度（0-1）"],
            ["questionCostFactor", this.DEFAULT_SETTINGS.questionCostFactor, "質問コスト影響係数 - 質問コストの影響度（0-1）"]
          ]
        }
      ];
      
      // データ書き込み
      let rowIndex = 2;
      
      settingsData.forEach(categoryData => {
        // カテゴリヘッダー
        sheet.getRange(rowIndex, 1, 1, 3).merge()
          .setValue(categoryData.category)
          .setBackground('#e3f2fd')
          .setFontWeight('bold');
        rowIndex++;
        
        // カテゴリ内の設定値
        categoryData.settings.forEach(setting => {
          sheet.getRange(rowIndex, 1, 1, 3).setValues([setting]);
          rowIndex++;
        });
        
        // カテゴリ間の余白
        rowIndex++;
      });
      
      // ヘルプとステータス情報
      const helpRowIndex = rowIndex + 1;
      sheet.getRange(helpRowIndex, 1, 1, 3).merge()
        .setValue("設定を変更した後は、「スキルマップ→設定値更新」を実行してください")
        .setBackground('#FFF9C4').setFontWeight('bold');
      
      sheet.getRange(helpRowIndex + 2, 1, 1, 3).merge()
        .setValue("最終更新日時: " + new Date().toLocaleString())
        .setFontStyle("italic");
      
      // 列幅調整
      sheet.setColumnWidth(1, 200);
      sheet.setColumnWidth(2, 100);
      sheet.setColumnWidth(3, 400);
      
      return sheet;
    } catch (e) {
      LogModule.handleError(e, '設定シート作成');
      return null;
    }
  },
  
  /**
   * 設定値を更新する（メニューから実行）
   */
  updateSettings: function() {
    try {
      LogModule.logOperation("設定値更新", "開始");
      
      // 設定シートを取得
      const settingsSheet = SS.getSheetByName(SHEETS.SETTINGS);
      if (!settingsSheet) {
        this.createSettingsSheet();
        UtilsModule.showToast("設定シートを新規作成しました", "設定更新");
        return;
      }
      
      // 設定値を読み込み、キャッシュをクリア
      this._settingsCache = null;
      const settings = this.getSettings();
      
      // 更新日時の更新
      const lastRow = settingsSheet.getLastRow();
      for (let i = lastRow; i > 0; i--) {
        const value = settingsSheet.getRange(i, 1).getValue();
        if (value && value.toString().includes("最終更新日時")) {
          settingsSheet.getRange(i, 1, 1, 3).merge()
            .setValue("最終更新日時: " + new Date().toLocaleString())
            .setFontStyle("italic");
          break;
        }
      }
      
      LogModule.logOperation("設定値更新", "完了");
      UtilsModule.showToast("設定値を更新しました", "設定更新");
    } catch (e) {
      LogModule.handleError(e, '設定値更新');
    }
  },

  /**
   * 設定値をデフォルトにリセットする
   */
  resetSettingsToDefault: function() {
    try {
      LogModule.logOperation("設定値リセット", "開始");
      
      // 設定シートを再作成
      const sheet = this.createSettingsSheet();
      
      // キャッシュをクリア
      this._settingsCache = null;
      
      LogModule.logOperation("設定値リセット", "完了");
      UtilsModule.showToast("設定値をデフォルトに戻しました", "設定リセット");
      
      return true;
    } catch (e) {
      LogModule.handleError(e, '設定値リセット');
      return false;
    }
  },

  // 設定値キャッシュ
  _settingsCache: null,

  /**
   * 設定値を取得する
   */
  getSettings: function() {
    // キャッシュがあれば返す
    if (this._settingsCache) {
      return this._settingsCache;
    }
    
    // 設定シートから取得を試みる
    const settingsSheet = SS.getSheetByName(SHEETS.SETTINGS);
    if (settingsSheet) {
      try {
        const data = settingsSheet.getDataRange().getValues();
        const settings = Object.assign({}, this.DEFAULT_SETTINGS); // デフォルト値でベースを作成
        
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] && data[i][1] !== undefined && !this.isCategory(data[i][0])) {
            // 数値または真偽値に変換
            const value = data[i][1];
            const key = data[i][0];
            
            if (typeof this.DEFAULT_SETTINGS[key] === 'number') {
              settings[key] = parseFloat(value) || this.DEFAULT_SETTINGS[key];
            } else if (typeof this.DEFAULT_SETTINGS[key] === 'boolean') {
              settings[key] = value === true || value === 'true';
            } else {
              settings[key] = value;
            }
          }
        }
        
        // キャッシュに保存
        this._settingsCache = settings;
        
        return settings;
      } catch (e) {
        LogModule.handleError(e, '設定値取得');
      }
    }
    
    // 設定シートがない場合は、設定シートを作成してデフォルト値を設定
    this.createSettingsSheet();
    
    // デフォルト値を返す
    this._settingsCache = Object.assign({}, this.DEFAULT_SETTINGS);
    return this._settingsCache;
  },
  
  /**
   * 設定値がカテゴリか判断する
   */
  isCategory: function(settingName) {
    // カテゴリ名のパターン（例：基本設定、技術的複雑性の閾値など）
    return typeof settingName === 'string' && 
           (settingName.includes('設定') || 
            settingName.includes('閾値') || 
            settingName.includes('パラメータ') ||
            settingName.includes('判定'));
  }
};