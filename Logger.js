// ロギングとエラー処理モジュール
const LogModule = {
    /**
     * 処理ログを記録する
     */
    logOperation: function(operation, result, details = '') {
      try {
        const logSheet = UtilsModule.getOrCreateSheet(SHEETS.LOG);
        
        // ヘッダーがなければ追加
        if (logSheet.getLastRow() === 0) {
          logSheet.getRange("A1:D1").setValues([["タイムスタンプ", "操作", "結果", "詳細"]])
            .setFontWeight('bold').setBackground('#f3f3f3');
          
          // 列幅調整
          logSheet.setColumnWidth(1, 180); // タイムスタンプ
          logSheet.setColumnWidth(2, 150); // 操作
          logSheet.setColumnWidth(3, 150); // 結果
          logSheet.setColumnWidth(4, 350); // 詳細
        }
        
        // ログ追加
        logSheet.appendRow([new Date(), operation, result, details]);
        
        console.log(`[${operation}] ${result} ${details ? '- ' + details : ''}`);
        
      } catch (e) {
        console.log('ログ記録エラー: ' + e.toString());
      }
    },
    
    /**
     * エラーを処理し、ログに記録する
     */
    handleError: function(error, operation) {
      const errorMessage = error.toString();
      const stack = error.stack || 'スタックトレースなし';
      
      // エラーログ記録
      this.logOperation(operation, "エラー", errorMessage);
      console.error(`[${operation}] エラー: ${errorMessage}\n${stack}`);
      
      // ユーザーに通知
      UtilsModule.showAlert(`${operation}でエラーが発生しました: ${errorMessage}`);
      
      return errorMessage;
    },
    
    /**
     * デバッグ用にデータの状態をログに記録する
     */
    debugData: function(dataObject, label = 'データデバッグ') {
      try {
        // 対象データの主要統計情報を記録
        let stats = {
          label: label,
          timestamp: new Date(),
          summary: {}
        };
        
        if (dataObject && typeof dataObject === 'object') {
          // 配列の場合
          if (Array.isArray(dataObject)) {
            stats.summary.type = 'Array';
            stats.summary.length = dataObject.length;
            
            // サンプルデータ（最初の数行）
            if (dataObject.length > 0) {
              stats.summary.sampleData = dataObject.slice(0, Math.min(3, dataObject.length));
            }
          } 
          // オブジェクトの場合
          else {
            stats.summary.type = 'Object';
            stats.summary.keys = Object.keys(dataObject);
            stats.summary.keyCount = stats.summary.keys.length;
            
            // キータイプの分析
            let valueSamples = {};
            for (const key of stats.summary.keys.slice(0, 5)) {
              valueSamples[key] = dataObject[key];
            }
            stats.summary.valueSamples = valueSamples;
          }
        } else {
          stats.summary.type = typeof dataObject;
          stats.summary.value = dataObject;
        }
        
        // ログに記録
        this.logOperation(label, "データデバッグ", JSON.stringify(stats.summary));
        console.log(`[${label}] デバッグ情報:`, stats);
        
        return stats;
      } catch (e) {
        this.handleError(e, 'データデバッグログ');
        return null;
      }
    }
  };