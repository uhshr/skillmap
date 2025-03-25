// ドキュメント生成モジュール
const DocumentationModule = {
    /**
     * 仕様書シートを更新する
     */
    updateSpecificationSheet: function() {
      LogModule.logOperation("仕様書更新", "開始");
      
      try {
        const docSheet = UtilsModule.getOrCreateSheet(SHEETS.DOCUMENTATION);
        
        // シートをクリア
        docSheet.clear();
        
        // タイトルと更新日
        docSheet.getRange("A1:H1").merge().setValue("スキルマップシステム 詳細仕様書")
          .setFontWeight('bold').setBackground('#E8EAF6').setFontSize(14);
        
        docSheet.getRange("A2:H2").merge().setValue(`更新日: ${new Date().toLocaleDateString()}`)
          .setFontStyle("italic");
        
        // 目次の作成
        this.createTableOfContents(docSheet);
        
        // 各シートの詳細仕様を追加
        let currentRow = 10;
        currentRow = this.addTagDifficultySpec(docSheet, currentRow);
        currentRow = this.addTagDistributionSpec(docSheet, currentRow);
        currentRow = this.addIndividualSkillSpec(docSheet, currentRow);
        currentRow = this.addTeamPerformanceSpec(docSheet, currentRow);
        currentRow = this.addAreaCoverageSpec(docSheet, currentRow);
        currentRow = this.addPersonalDashboardSpec(docSheet, currentRow);
        currentRow = this.addTeamDashboardSpec(docSheet, currentRow);
        currentRow = this.addAreaDashboardSpec(docSheet, currentRow);
        currentRow = this.addAdjustmentMatrixSpec(docSheet, currentRow);
        currentRow = this.addSettingsSheetSpec(docSheet, currentRow);
        
        // 列幅の調整
        docSheet.setColumnWidth(1, 150);
        docSheet.setColumnWidth(2, 250);
        docSheet.setColumnWidth(3, 300);
        
        LogModule.logOperation("仕様書更新", "完了");
        UtilsModule.showToast("仕様書が更新されました", "仕様書更新");
        
        return true;
      } catch (e) {
        LogModule.handleError(e, "仕様書更新");
        return false;
      }
    },
    
    /**
     * 目次を作成する
     */
    createTableOfContents: function(sheet) {
      const tocItems = [
        "1. タグ難易度分析シート",
        "2. タグ内難易度分布シート",
        "3. 個人スキル評価シート",
        "4. チームパフォーマンスシート",
        "5. 機能領域カバレッジシート",
        "6. 個人スキルダッシュボード",
        "7. チームパフォーマンスダッシュボード",
        "8. 機能領域カバレッジダッシュボード",
        "9. 難易度調整マトリックス",
        "10. 設定シート"
      ];
      
      sheet.getRange("A4").setValue("目次");
      
      for (let i = 0; i < tocItems.length; i++) {
        sheet.getRange(5 + i, 1).setValue(tocItems[i]);
      }
    },
    
    /**
     * タグ難易度分析シートの仕様を追加
     */
    addTagDifficultySpec: function(sheet, startRow) {
      sheet.getRange(startRow, 1, 1, 3).merge()
        .setValue("1. タグ難易度分析シート")
        .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
      
      startRow++;
      sheet.getRange(startRow, 1).setValue("目的");
      sheet.getRange(startRow, 2, 1, 2).merge()
        .setValue("機能タグの難易度を多次元的に評価し、対応の複雑さを客観的に分析します。");
      
      startRow += 2;
      sheet.getRange(startRow, 1).setValue("カラム説明");
      startRow++;
      
      // カラム説明テーブルのヘッダー
      sheet.getRange(startRow, 1, 1, 3).setValues([["カラム名", "説明", "ロジック・計算方法"]])
        .setFontWeight('bold').setBackground('#f3f3f3');
      
      startRow++;
      
      // カラムの説明データ
      const columns = [
        ["タグ名", "機能タグの名称", "取込データから抽出"],
        ["OPチーム名", "担当OPチーム", "取込データから抽出"],
        ["対応件数", "タグの総対応件数", "タグごとの対応件数を集計"],
        ["Rally平均", "対応1件あたりの平均Rally数", "Rally数の合計 ÷ 対応件数"],
        ["対応者数", "タグに対応できる担当者数", "ユニークな担当者数をカウント"],
        ["対応者カバー率(%)", "全担当者中このタグに対応できる割合", "(対応者数 ÷ 全担当者数) × 100"],
        ["対応時間中央値", "タグの対応時間の中央値", "対応時間の中央値（50パーセンタイル）"],
        ["25パーセンタイル", "対応時間の下位四分位", "対応時間の25パーセンタイル"],
        ["75パーセンタイル", "対応時間の上位四分位", "対応時間の75パーセンタイル"],
        ["IQR", "四分位範囲", "75パーセンタイル - 25パーセンタイル"],
        ["マクロ利用率(%)", "マクロを使用した対応の割合", "(マクロ使用件数 ÷ 対応件数) × 100"],
        ["相談発生率(%)", "相談が発生した対応の割合", "(相談発生件数 ÷ 対応件数) × 100"],
        ["技術的複雑性(1-5)", "技術的な複雑さのスコア", "対応時間(40%) + カバレッジ(30%) + マクロ利用率(30%)から算出"],
        ["技術的複雑性レベル", "技術的複雑性のL1-L5評価", "技術的複雑性スコアをL1-L5に変換"],
        ["解決難易度(1-5)", "解決プロセスの難しさスコア", "Rally数(50%) + 相談率(30%) + 対応時間のばらつき(20%)から算出"],
        ["解決難易度レベル", "解決難易度のL1-L5評価", "解決難易度スコアをL1-L5に変換"],
        ["知識要求度(1-5)", "必要な知識の深さスコア", "カバレッジ(70%) + 対応時間のばらつき(30%)から算出"],
        ["知識要求度レベル", "知識要求度のL1-L5評価", "知識要求度スコアをL1-L5に変換"],
        ["総合難易度スコア(1-5)", "総合的な難易度スコア", "技術的複雑性(40%) + 解決難易度(40%) + 知識要求度(20%)の加重平均"],
        ["難易度レベル", "総合難易度のL1-L5評価", "総合難易度スコアをL1-L5に変換"],
        ["タグタイプ", "コア/標準/レア分類", "出現頻度による分類（コア: 20%以上、レア: 5%以下）"],
        ["初級者向け", "初級者が対応しやすいか", "マクロ利用率50%以上で「○」"],
        ["高相談タグ", "相談が多く発生するか", "相談発生率30%以上で「○」"],
        ["更新日時", "データ更新日時", "分析実行時の日時"]
      ];
      
      // カラム説明をシートに書き込み
      sheet.getRange(startRow, 1, columns.length, 3).setValues(columns);
      
      startRow += columns.length + 2;
      
      // 主要ロジックの説明
      sheet.getRange(startRow, 1).setValue("主要ロジック");
      startRow++;
      
      const logicDesc = [
        "1. 技術的複雑性評価:",
        "   * 対応時間中央値が長いほど複雑",
        "   * 対応者カバー率が低いほど複雑",
        "   * マクロ利用率が低いほど複雑",
        "2. 解決難易度評価:",
        "   * Rally数が多いほど難しい",
        "   * 相談率が高いほど難しい",
        "   * 対応時間のばらつき(IQR)が大きいほど難しい",
        "3. 知識要求度評価:",
        "   * 対応者カバー率が低いほど専門知識が必要",
        "   * 対応時間のばらつきが大きいほど知識要求が高い",
        "4. 総合難易度:",
        "   * 上記三次元を加重平均して算出",
        "   * 技術と解決の比重が高い (各40%)"
      ];
      
      for (let i = 0; i < logicDesc.length; i++) {
        sheet.getRange(startRow + i, 1, 1, 3).merge().setValue(logicDesc[i]);
      }
      
      return startRow + logicDesc.length + 2; // 次のセクションの開始行を返す
    },
    
    /**
     * タグ内難易度分布シートの仕様を追加
     */
    addTagDistributionSpec: function(sheet, startRow) {
      sheet.getRange(startRow, 1, 1, 3).merge()
        .setValue("2. タグ内難易度分布シート")
        .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
      
      startRow++;
      sheet.getRange(startRow, 1).setValue("目的");
      sheet.getRange(startRow, 2, 1, 2).merge()
        .setValue("同一タグ内での案件の複雑さの分布を分析し、簡易/標準/複雑の3種類に分類します。これにより、タグごとの詳細な難易度パターンを把握できます。");
      
      startRow += 2;
      sheet.getRange(startRow, 1).setValue("カラム説明");
      startRow++;
      
      // カラム説明テーブルのヘッダー
      sheet.getRange(startRow, 1, 1, 3).setValues([["カラム名", "説明", "ロジック・計算方法"]])
        .setFontWeight('bold').setBackground('#f3f3f3');
      
      startRow++;
      
      // カラムの説明データ
      const columns = [
        ["タグ名", "機能タグの名称", "取込データから抽出"],
        ["件数", "タグの総対応件数", "タグごとの対応件数を集計"],
        ["基本難易度", "タグの基本難易度レベル", "タグ難易度分析からのレベル(L1-L5)"],
        ["月間件数", "月間の平均対応件数", "総件数 ÷ 3ヶ月（サンプル期間に応じて調整）"],
        ["時間P25", "対応時間の25パーセンタイル", "下位四分位の対応時間"],
        ["時間中央値", "対応時間の中央値", "50パーセンタイルの対応時間"],
        ["時間P75", "対応時間の75パーセンタイル", "上位四分位の対応時間"],
        ["RallyP25", "Rally数の25パーセンタイル", "下位四分位のRally数"],
        ["Rally中央値", "Rally数の中央値", "50パーセンタイルのRally数"],
        ["RallyP75", "Rally数の75パーセンタイル", "上位四分位のRally数"],
        ["マクロ使用率", "マクロを使用した対応の割合", "(マクロ使用件数 ÷ 対応件数) × 100"],
        ["相談率", "相談が発生した対応の割合", "(相談発生件数 ÷ 対応件数) × 100"],
        ["簡易ケース件数", "簡易ケースの件数", "複雑性スコアが下位25%のケース数"],
        ["標準ケース件数", "標準ケースの件数", "複雑性スコアが中間50%のケース数"],
        ["複雑ケース件数", "複雑ケースの件数", "複雑性スコアが上位25%のケース数"],
        ["簡易ケース割合", "簡易ケースの割合", "(簡易ケース件数 ÷ 総件数) × 100"],
        ["標準ケース割合", "標準ケースの割合", "(標準ケース件数 ÷ 総件数) × 100"],
        ["複雑ケース割合", "複雑ケースの割合", "(複雑ケース件数 ÷ 総件数) × 100"],
        ["簡易ケース基準", "簡易ケースの判定基準", "時間P25以下、複雑性スコア下位25%"],
        ["標準ケース基準", "標準ケースの判定基準", "複雑性スコア中間50%"],
        ["複雑ケース基準", "複雑ケースの判定基準", "時間P75以上、複雑性スコア上位25%"],
        ["簡易ケースレベル", "簡易ケースのスキルレベル", "基本難易度から1-2レベル下"],
        ["標準ケースレベル", "標準ケースのスキルレベル", "基本難易度と同じ"],
        ["複雑ケースレベル", "複雑ケースのスキルレベル", "基本難易度から1-2レベル上"]
      ];
      
      // カラム説明をシートに書き込み
      sheet.getRange(startRow, 1, columns.length, 3).setValues(columns);
      
      startRow += columns.length + 2;
      
      // 主要ロジックの説明
      sheet.getRange(startRow, 1).setValue("主要ロジック");
      startRow++;
      
      const logicDesc = [
        "1. 複雑性スコア:",
        "   * 対応時間(40%)、Rally数(30%)、マクロ非使用(20%)、相談あり(10%)の加重平均",
        "   * 複雑性スコアの分布により、ケースを三分類",
        "2. ケースタイプ判定:",
        "   * 簡易: 複雑性スコアが下位25%",
        "   * 標準: 複雑性スコアが中間50%",
        "   * 複雑: 複雑性スコアが上位25%",
        "3. ケースレベル判定:",
        "   * 基本難易度を中心に、ケースタイプに応じてレベルを相対的に設定",
        "   * 例: L3タグの場合 → 簡易:L2、標準:L3、複雑:L4"
      ];
      
      for (let i = 0; i < logicDesc.length; i++) {
        sheet.getRange(startRow + i, 1, 1, 3).merge().setValue(logicDesc[i]);
      }
      
      return startRow + logicDesc.length + 2; // 次のセクションの開始行を返す
    },
    
    /**
     * 他のシート仕様も同様の形式で実装
     */
    addIndividualSkillSpec: function(sheet, startRow) {
      // 個人スキル評価シートの仕様を追加
      sheet.getRange(startRow, 1, 1, 3).merge()
        .setValue("3. 個人スキル評価シート")
        .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
      
      startRow++;
      sheet.getRange(startRow, 1).setValue("目的");
      sheet.getRange(startRow, 2, 1, 2).merge()
        .setValue("担当者のスキルを多角的に評価し、強み・弱みを可視化します。また、パフォーマータイプを判定し、個人の特性を明らかにします。");
      
      // 詳細は省略（実装時に追加）
      return startRow + 15;
    },
    
    addTeamPerformanceSpec: function(sheet, startRow) {
      sheet.getRange(startRow, 1, 1, 3).merge()
        .setValue("4. チームパフォーマンスシート")
        .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
      
      startRow++;
      sheet.getRange(startRow, 1).setValue("目的");
      sheet.getRange(startRow, 2, 1, 2).merge()
        .setValue("チームごとのパフォーマンスを分析し、チーム内完結率や必要HCなどを算出します。これにより、チームのリソース状況を把握します。");
      
      // 詳細は省略（実装時に追加）
      return startRow + 15;
    },
    
    addAreaCoverageSpec: function(sheet, startRow) {
      sheet.getRange(startRow, 1, 1, 3).merge()
        .setValue("5. 機能領域カバレッジシート")
        .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
      
      startRow++;
      sheet.getRange(startRow, 1).setValue("目的");
      sheet.getRange(startRow, 2, 1, 2).merge()
        .setValue("機能領域（OPチーム）ごとの対応状況を分析し、リソースの適切な配分や不足領域を把握します。");
      
      // 詳細は省略（実装時に追加）
      return startRow + 15;
    },
    
    addPersonalDashboardSpec: function(sheet, startRow) {
      sheet.getRange(startRow, 1, 1, 3).merge()
        .setValue("6. 個人スキルダッシュボード")
        .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
      
      startRow++;
      sheet.getRange(startRow, 1).setValue("目的");
      sheet.getRange(startRow, 2, 1, 2).merge()
        .setValue("特定の担当者のスキル状況を詳細に可視化し、強み・弱み・成長領域を把握します。");
      
      // 詳細は省略（実装時に追加）
      return startRow + 15;
    },
    
    addTeamDashboardSpec: function(sheet, startRow) {
      sheet.getRange(startRow, 1, 1, 3).merge()
        .setValue("7. チームパフォーマンスダッシュボード")
        .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
      
      startRow++;
      sheet.getRange(startRow, 1).setValue("目的");
      sheet.getRange(startRow, 2, 1, 2).merge()
        .setValue("チームごとのパフォーマンス状況を可視化し、リソース配分や人員計画の意思決定を支援します。");
      
      // 詳細は省略（実装時に追加）
      return startRow + 15;
    },
    
    addAreaDashboardSpec: function(sheet, startRow) {
      sheet.getRange(startRow, 1, 1, 3).merge()
        .setValue("8. 機能領域カバレッジダッシュボード")
        .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
      
      startRow++;
      sheet.getRange(startRow, 1).setValue("目的");
      sheet.getRange(startRow, 2, 1, 2).merge()
        .setValue("機能領域（OPチーム）ごとのカバレッジ状況や対応分布を可視化し、リスクタグの特定やリソース配分の最適化を支援します。");
      
      // 詳細は省略（実装時に追加）
      return startRow + 15;
    },
    
    addAdjustmentMatrixSpec: function(sheet, startRow) {
      sheet.getRange(startRow, 1, 1, 3).merge()
        .setValue("9. 難易度調整マトリックス")
        .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
      
      startRow++;
      sheet.getRange(startRow, 1).setValue("目的");
      sheet.getRange(startRow, 2, 1, 2).merge()
        .setValue("自動計算された機能タグの難易度に対して、定性的な調整を加えるためのシートです。システムが算出した難易度に対して、実際の業務知識に基づく微調整を可能にします。");
      
      // 詳細は省略（実装時に追加）
      return startRow + 15;
    },
    
    addSettingsSheetSpec: function(sheet, startRow) {
      sheet.getRange(startRow, 1, 1, 3).merge()
        .setValue("10. 設定シート")
        .setFontWeight('bold').setBackground('#E3F2FD').setFontSize(12);
      
      startRow++;
      sheet.getRange(startRow, 1).setValue("目的");
      sheet.getRange(startRow, 2, 1, 2).merge()
        .setValue("システム全体の設定値を管理し、難易度計算や分析ロジックのパラメータを調整可能にします。");
      
      // 詳細は省略（実装時に追加）
      return startRow + 15;
    }
  };