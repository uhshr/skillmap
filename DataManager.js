// データ管理モジュール
const DataModule = {
    /**
     * 担当者マスターデータの自動生成
     */
    generateMasterData: function() {
      LogModule.logOperation("マスターデータ生成", "開始");
      
      try {
        // シート取得
        const dataSheet = UtilsModule.getSheet(SHEETS.DATA);
        const masterSheet = UtilsModule.getOrCreateSheet(SHEETS.MASTER);
        
        if (!dataSheet) {
          throw new Error('取込データシートが見つかりません');
        }
        
        // 取込データからユニークな担当者情報を抽出
        const dataRange = dataSheet.getDataRange();
        const data = dataRange.getValues();
        const headers = data[0];
        
        const colIndexes = this.getColumnIndexes(headers, [
          'user_name', 'ih_pt', 'div'
        ]);
        
        if (colIndexes.user_name === -1 || colIndexes.ih_pt === -1) {
          throw new Error('必要な列（user_name, ih_pt）が見つかりません');
        }
        
        // ユニークな担当者情報を抽出
        const uniqueUsers = this.extractUniqueUsers(data, colIndexes);
        
        // マスターシートをクリア
        masterSheet.clear();
        
        // マスターデータの作成と書き込み
        this.writeMasterData(masterSheet, uniqueUsers, headers, colIndexes);
        
        LogModule.logOperation("マスターデータ生成", `完了: ${Object.keys(uniqueUsers).length}人の担当者情報を抽出`);
        UtilsModule.showToast(`${Object.keys(uniqueUsers).length}人の担当者マスターデータを生成しました！`, 'マスターデータ生成');
        
        // メインメニューに戻る前に確認メッセージを表示
        UtilsModule.showAlert('担当者マスターデータを生成しました。主所属OPチーム、兼務OPチーム、稼働時間係数は手動で入力してください。');
        
        return true;
      } catch (e) {
        return LogModule.handleError(e, 'マスターデータ生成');
      }
    },
    
    /**
     * 列インデックスを取得する
     */
    getColumnIndexes: function(headers, columnNames) {
      const indexes = {};
      
      columnNames.forEach(name => {
        indexes[name] = headers.indexOf(name);
      });
      
      return indexes;
    },
    
    /**
     * ユニークな担当者情報を抽出する
     */
    extractUniqueUsers: function(data, colIndexes) {
      const uniqueUsers = {};
      
      for (let i = 1; i < data.length; i++) {
        const userName = data[i][colIndexes.user_name];
        if (userName && !uniqueUsers[userName]) {
          uniqueUsers[userName] = {
            name: userName,
            ihPt: data[i][colIndexes.ih_pt] || '',
            div: colIndexes.div !== -1 ? data[i][colIndexes.div] || '' : '',
          };
        }
      }
      
      return uniqueUsers;
    },
    
    /**
     * マスターデータをシートに書き込む
     */
    writeMasterData: function(sheet, uniqueUsers, headers, colIndexes) {
      // ヘッダー設定
      const outputheaders = [
        "担当者ID", "担当者名", "IH/PT区分", "職種", "等級", "役職", 
        "主所属OPチーム", "兼務OPチーム", "稼働時間係数"
      ];
      
      sheet.getRange(1, 1, 1, outputheaders.length).setValues([outputheaders])
        .setFontWeight('bold').setBackground('#f3f3f3');
      
      // 担当者データの作成
      const settings = SettingsModule.getSettings();
      const masterData = Object.values(uniqueUsers).map((user, index) => [
        index + 1,              // 担当者ID
        user.name,              // 担当者名
        user.ihPt,              // IH/PT区分
        user.div,               // 職種
        "",                     // 等級 (手動入力用)
        "",                     // 役職 (手動入力用)
        "",                     // 主所属OPチーム (手動入力用)
        "",                     // 兼務OPチーム (手動入力用)
        user.ihPt === "IH" ? settings.ihDefaultHours : settings.ptDefaultHours  // 稼働時間係数
      ]);
      
      // データをシートに書き込み
      if (masterData.length > 0) {
        sheet.getRange(2, 1, masterData.length, outputheaders.length).setValues(masterData);
      }
      
      // 列幅の調整
      sheet.setColumnWidth(1, 80);  // 担当者ID
      sheet.setColumnWidth(2, 150); // 担当者名
      sheet.setColumnWidth(7, 150); // 主所属OPチーム
      sheet.setColumnWidth(8, 200); // 兼務OPチーム
      
      // ヘルプ情報の追加
      const helpRowIndex = masterData.length + 3;
      sheet.getRange(helpRowIndex, 1, 1, 9).merge()
        .setValue("【注意】主所属OPチーム、兼務OPチーム、稼働時間係数は手動で入力してください")
        .setBackground('#FFF9C4').setFontWeight('bold');
      
      // OPチームのデータ検証設定
      this.setupOPTeamValidation(sheet, headers, colIndexes, masterData.length);
    },
    
    /**
     * OPチームのデータ検証（ドロップダウン）を設定
     */
    setupOPTeamValidation: function(sheet, headers, colIndexes, dataLength) {
      // 主所属OPチームのドロップダウンリスト用に既存のOPチーム情報を収集
      const opTeams = new Set();
      const opTeamIdx = headers.indexOf('op_name');
      
      if (opTeamIdx !== -1) {
        const dataSheet = UtilsModule.getSheet(SHEETS.DATA);
        const data = dataSheet.getDataRange().getValues();
        
        for (let i = 1; i < data.length; i++) {
          const opTeam = data[i][opTeamIdx];
          if (opTeam) opTeams.add(opTeam);
        }
        
        // OPチームリストがあれば、データ検証（ドロップダウン）を設定
        if (opTeams.size > 0) {
          const opTeamList = Array.from(opTeams).sort();
          const validation = SpreadsheetApp.newDataValidation()
            .requireValueInList(opTeamList, true)
            .build();
          
          // 主所属OPチーム列にデータ検証を設定
          sheet.getRange(2, 7, dataLength, 1).setDataValidation(validation);
        }
      }
  },
  
  /**
   * 分析用のデータを前処理して返す
   * @return {Object} 前処理済みのデータセット
   */
    prepareAnalysisData: function() {
      try {
        // シート取得
        const dataSheet = UtilsModule.getSheet(SHEETS.DATA);
        const masterSheet = UtilsModule.getSheet(SHEETS.MASTER);
        const difficultySheet = UtilsModule.getSheet(SHEETS.TAG_DIFFICULTY);
        const distributionSheet = UtilsModule.getSheet(SHEETS.TAG_DISTRIBUTION);
        const adjustmentSheet = UtilsModule.getSheet(SHEETS.ADJUSTMENT);
        
        // 必要なシートが存在するか確認
        if (!dataSheet) {
          throw new Error('取込データシートが見つかりません');
        }
        
        // 各シートからデータを取得
        const transactionData = dataSheet.getDataRange().getValues();
        const txHeaders = transactionData[0];  // 変数名を修正: transactionHeaders → txHeaders
        
        LogModule.debugData(transactionData, 'データ行数');
        LogModule.debugData(txHeaders, 'ヘッダー行');
        
        // 機能タグがあるかチェック
        const tagColIdx = txHeaders.indexOf('tag_name');
        if (tagColIdx !== -1) {
          const functionTags = transactionData.filter((row, idx) => 
            idx > 0 && row[tagColIdx] && UtilsModule.isFunctionTag(row[tagColIdx])
          );
          LogModule.debugData(functionTags.length, '機能タグ数');
        }
        
        // マスターデータの取得と整形
        let masterData = [];
        let masterHeaders = [];
        if (masterSheet) {
          masterData = masterSheet.getDataRange().getValues();
          masterHeaders = masterData[0];
        }
        
        // 難易度データの取得と整形
        let difficultyData = [];
        let difficultyHeaders = [];
        if (difficultySheet) {
          difficultyData = difficultySheet.getDataRange().getValues();
          difficultyHeaders = difficultyData[0];
        }
        
        // タグ内難易度分布データの取得と整形
        let distributionData = [];
        let distributionHeaders = [];
        if (distributionSheet) {
          distributionData = distributionSheet.getDataRange().getValues();
          distributionHeaders = distributionData[0];
        }
        
        // 難易度調整データの取得と整形
        let adjustmentData = {};
        if (adjustmentSheet) {
          const adjustmentValues = adjustmentSheet.getDataRange().getValues();
          const adjustmentHeaders = adjustmentValues[0];
          
          const tagColIdx = adjustmentHeaders.indexOf('タグ名');
          const finalLevelColIdx = adjustmentHeaders.indexOf('最終難易度');
          
          if (tagColIdx !== -1 && finalLevelColIdx !== -1) {
            for (let i = 1; i < adjustmentValues.length; i++) {
              const tag = adjustmentValues[i][tagColIdx];
              const level = adjustmentValues[i][finalLevelColIdx];
              if (tag && level) {
                adjustmentData[tag] = level;
              }
            }
          }
        }
        
        // 列インデックスの取得
        const txColIndexes = this.getTransactionColumnIndexes(txHeaders);
        const masterColIndexes = this.getMasterColumnIndexes(masterHeaders);
        const diffColIndexes = this.getDifficultyColumnIndexes(difficultyHeaders);
        const distColIndexes = this.getDistributionColumnIndexes(distributionHeaders);
        
        // 各種マッピングデータの作成
        const mappings = this.createDataMappings(
          transactionData, masterData, difficultyData, distributionData, 
          txColIndexes, masterColIndexes, diffColIndexes, distColIndexes,
          adjustmentData
        );
        
        return {
          // 元データ
          transactionData: transactionData,
          masterData: masterData,
          difficultyData: difficultyData,
          distributionData: distributionData,
          
          // ヘッダー
          txHeaders: txHeaders,
          masterHeaders: masterHeaders,
          difficultyHeaders: difficultyHeaders,
          distributionHeaders: distributionHeaders,
          
          // 列インデックス
          txColIndexes: txColIndexes,
          masterColIndexes: masterColIndexes,
          diffColIndexes: diffColIndexes,
          distColIndexes: distColIndexes,
          
          // マッピングデータ
          tagDifficultyMap: mappings.tagDifficultyMap,
          tagDistributionMap: mappings.tagDistributionMap,
          responderMasterMap: mappings.responderMasterMap,
          
          // 調整データ
          adjustmentData: adjustmentData,
          
          // シート参照
          sheets: {
            dataSheet: dataSheet,
            masterSheet: masterSheet,
            difficultySheet: difficultySheet,
            distributionSheet: distributionSheet,
            adjustmentSheet: adjustmentSheet
          }
        };
      } catch (e) {
        LogModule.handleError(e, 'データ前処理');
        return null;
      }
    },
    
    /**
     * 取込データの列インデックスを取得
     */
    getTransactionColumnIndexes: function(headers) {
      const indices = {
        tagName: headers.indexOf('tag_name'),
        userName: headers.indexOf('user_name'),
        ihPt: headers.indexOf('ih_pt'),
        div: headers.indexOf('div'),
        opName: headers.indexOf('op_name'),
        responseTime: headers.indexOf('total_response_time'),
        rallyCount: headers.indexOf('rally_count'),
        useMacro: headers.indexOf('use_macro'),
        postTimestamp: headers.indexOf('post_timestamp'),
        soudanUserName: headers.indexOf('soudan_user_name'),
        adviserUserName: headers.indexOf('adviser_user_name'),
        // 新規追加：相談関連
        solveDuration: headers.indexOf('solve_duration'),
        soudanUserDiv: headers.indexOf('soudan_user_div'),
        soudanUserRole: headers.indexOf('soudan_user_role'),
        adviserUserDiv: headers.indexOf('adviser_user_div'),
        adviserUserRole: headers.indexOf('adviser_user_role'),
        // OJT関連
        hasOjtTag: headers.indexOf('has_ojt_tag')
      };
  
      // デバッグログ
      LogModule.debugData(indices, '取得したカラムインデックス');
      return indices;
    },
  
    /**
     * マスターデータの列インデックスを取得
     */
    getMasterColumnIndexes: function(headers) {
      if (!headers || headers.length === 0) {
        return {
          name: -1,
          ihPt: -1,
          primaryTeam: -1,
          secondaryTeam: -1,
          workTimeRatio: -1
        };
      }
      
      // デバッグのためにヘッダー情報を出力
      LogModule.debugData(headers, 'マスターヘッダー');
      
      // 実際の列名とのマッピング
      const indices = {
        name: headers.indexOf('担当者名'),
        ihPt: headers.indexOf('IH/PT区分'),
        primaryTeam: headers.indexOf('主所属OPチーム'),
        secondaryTeam: headers.indexOf('兼務OPチーム'),
        workTimeRatio: headers.indexOf('稼働時間係数')
      };
      
      // 取得したインデックスのデバッグ出力
      LogModule.debugData(indices, 'マスターカラムインデックス');
      
      return indices;
    },
  
    /**
     * 難易度データの列インデックスを取得
     */
    getDifficultyColumnIndexes: function(headers) {
      if (!headers || headers.length === 0) {
        return {
          tagName: -1,
          opTeam: -1,
          level: -1,
          score: -1,
          tcLevel: -1,
          rdLevel: -1,
          krLevel: -1
        };
      }
      
      return {
        tagName: headers.indexOf('タグ名'),
        opTeam: headers.indexOf('OPチーム名'),
        level: headers.indexOf('難易度レベル'),
        score: headers.indexOf('総合難易度スコア(1-5)'),
        tcLevel: headers.indexOf('技術的複雑性レベル'),
        rdLevel: headers.indexOf('解決難易度レベル'),
        krLevel: headers.indexOf('知識要求度レベル'),
        macroRate: headers.indexOf('マクロ利用率(%)'),
        consultRate: headers.indexOf('相談発生率(%)')
      };
    },
    
    /**
     * タグ内難易度分布データの列インデックスを取得
     */
    getDistributionColumnIndexes: function(headers) {
      if (!headers || headers.length === 0) {
        return {
          tagName: -1,
          simpleThreshold: -1,
          complexThreshold: -1,
          simpleLevel: -1,
          standardLevel: -1,
          complexLevel: -1
        };
      }
      
      return {
        tagName: headers.indexOf('タグ名'),
        simpleThreshold: headers.indexOf('簡易ケース基準'),
        complexThreshold: headers.indexOf('複雑ケース基準'),
        simpleLevel: headers.indexOf('簡易ケースレベル'),
        standardLevel: headers.indexOf('標準ケースレベル'),
        complexLevel: headers.indexOf('複雑ケースレベル')
      };
    },
    
    /**
     * 分析に使用する各種マッピングデータを作成
     */
    createDataMappings: function(
      transactionData, masterData, difficultyData, distributionData,
      txColIndexes, masterColIndexes, diffColIndexes, distColIndexes,
      adjustmentData
    ) {
      // タグ難易度マッピング
      const tagDifficultyMap = {};
      if (difficultyData.length > 1 && diffColIndexes.tagName !== -1) {
        for (let i = 1; i < difficultyData.length; i++) {
          const row = difficultyData[i];
          const tagName = row[diffColIndexes.tagName];
          
          if (tagName) {
            // 調整後の難易度レベルがあれば使用
            const adjustedLevel = adjustmentData[tagName];
            
            tagDifficultyMap[tagName] = {
              level: adjustedLevel || row[diffColIndexes.level],
              score: parseFloat(row[diffColIndexes.score] || 3),
              tcLevel: row[diffColIndexes.tcLevel] || "L3",
              rdLevel: row[diffColIndexes.rdLevel] || "L3",
              krLevel: row[diffColIndexes.krLevel] || "L3",
              opTeam: row[diffColIndexes.opTeam] || "",
              macroRate: parseFloat(row[diffColIndexes.macroRate] || 0),
              consultRate: parseFloat(row[diffColIndexes.consultRate] || 0)
            };
          }
        }
      }
      
      // タグ内難易度分布マッピング
      const tagDistributionMap = {};
      if (distributionData.length > 1 && distColIndexes.tagName !== -1) {
        for (let i = 1; i < distributionData.length; i++) {
          const row = distributionData[i];
          const tagName = row[distColIndexes.tagName];
          
          if (tagName) {
            tagDistributionMap[tagName] = {
              simpleThreshold: row[distColIndexes.simpleThreshold] || "",
              complexThreshold: row[distColIndexes.complexThreshold] || "",
              simpleLevel: row[distColIndexes.simpleLevel] || "L1",
              standardLevel: row[distColIndexes.standardLevel] || "L2",
              complexLevel: row[distColIndexes.complexLevel] || "L3"
            };
          }
        }
      }
      
      // 担当者マスターマッピング
      const responderMasterMap = {};
      if (masterData.length > 1 && masterColIndexes.name !== -1) {
        // デバッグ情報
        LogModule.debugData(masterData.length, 'マスターデータ行数');
        LogModule.debugData(masterColIndexes, 'マスターカラム位置');
        
        for (let i = 1; i < masterData.length; i++) {
          const row = masterData[i];
          if (masterColIndexes.name !== -1 && row[masterColIndexes.name]) {
            const name = row[masterColIndexes.name];
            
            // primaryTeam, secondaryTeam, workTimeRatioの値をデバッグ
            const primaryTeam = masterColIndexes.primaryTeam !== -1 ? row[masterColIndexes.primaryTeam] : "";
            const secondaryTeam = masterColIndexes.secondaryTeam !== -1 ? row[masterColIndexes.secondaryTeam] : "";
            const workTimeRatio = masterColIndexes.workTimeRatio !== -1 ? row[masterColIndexes.workTimeRatio] : 0;
            
            // 最初の数件のマスターデータをデバッグ出力
            if (i <= 3) {
              LogModule.debugData({
                name: name,
                primaryTeam: primaryTeam, 
                secondaryTeam: secondaryTeam,
                workTimeRatio: workTimeRatio
              }, `マスターデータサンプル ${i}`);
            }
            
            responderMasterMap[name] = {
              primaryTeam: primaryTeam,
              secondaryTeam: secondaryTeam,
              ihPtRatio: workTimeRatio
            };
          }
        }
        
        // マッピング結果の最初の数件をデバッグ表示
        const sampleNames = Object.keys(responderMasterMap).slice(0, 3);
        for (const name of sampleNames) {
          LogModule.debugData(responderMasterMap[name], `マスターマッピングサンプル: ${name}`);
        }
      }
      
      return {
        tagDifficultyMap: tagDifficultyMap,
        tagDistributionMap: tagDistributionMap,
        responderMasterMap: responderMasterMap
      };
    }
  };