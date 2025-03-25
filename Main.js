// グローバル変数定義
const SS = SpreadsheetApp.getActiveSpreadsheet();

// シート名定義
const SHEETS = {
  DATA: "取込データ",
  MASTER: "担当者マスター",
  SETTINGS: "設定",
  TAG_DIFFICULTY: "タグ難易度分析",
  TAG_DISTRIBUTION: "タグ内難易度分布",
  ADJUSTMENT: "難易度調整マトリックス",
  INDIVIDUAL_SKILL: "個人スキル評価",
  TEAM_PERFORMANCE: "チームパフォーマンス",
  AREA_COVERAGE: "機能領域カバレッジ",
  PERSONAL_DASHBOARD: "個人スキルダッシュボード",
  TEAM_PERFORMANCE_DASHBOARD: "チームパフォーマンスダッシュボード",
  AREA_COVERAGE_DASHBOARD: "機能領域カバレッジダッシュボード",
  DOCUMENTATION: "仕様書",
  LOG: "処理ログ",
  CONSULTATION_FLOW: "相談フロー分析",
  CONSULTATION_FLOW_DASHBOARD: "相談フローダッシュボード",
  INTEGRATED_DASHBOARD: "統合パフォーマンスダッシュボード"
};

/**
 * メニューを作成する（onOpenトリガー用）
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('スキルマップ')
      .addItem('データを分析する', 'analyzeAllData')
      .addSeparator()
      .addItem('1. タグ難易度分析', 'AnalysisModule.analyzeTagDifficulty')
      .addItem('2. タグ内難易度分布分析', 'AnalysisModule.analyzeTagDistribution')
      .addItem('3. 個人スキル分析', 'AnalysisModule.analyzeIndividualSkills')
      .addItem('4. チームパフォーマンス分析', 'AnalysisModule.analyzeTeamPerformance')
      .addItem('5. 機能領域カバレッジ分析', 'AnalysisModule.analyzeAreaCoverage')
      .addSeparator()
      .addItem('担当者マスターデータ生成', 'DataModule.generateMasterData')
      .addSeparator()
      .addSubMenu(ui.createMenu('ダッシュボード')
        .addItem('個人スキルダッシュボード初期化', 'DashboardModule.initPersonalDashboard')
        .addItem('個人スキルダッシュボード更新', 'DashboardModule.updatePersonalDashboard')
        .addItem('チームパフォーマンスダッシュボード初期化', 'DashboardModule.initTeamPerformanceDashboard')
        .addItem('チームパフォーマンスダッシュボード更新', 'DashboardModule.updateTeamPerformance')
        .addItem('機能領域カバレッジダッシュボード初期化', 'DashboardModule.initAreaCoverageDashboard')
        .addItem('機能領域カバレッジダッシュボード更新', 'DashboardModule.updateAreaCoverage')
        .addItem('相談フローダッシュボード初期化', 'DashboardModule.initConsultationFlowDashboard')
        .addItem('相談フローダッシュボード更新', 'DashboardModule.updateConsultationFlow')
        .addSeparator()
        .addItem('統合パフォーマンスダッシュボード初期化', 'DashboardModule.initIntegratedDashboard')
        .addItem('統合パフォーマンスダッシュボード更新', 'DashboardModule.updateIntegratedDashboard')
      )
      .addSeparator()
      .addSubMenu(ui.createMenu('設定')
        .addItem('設定シート作成', 'SettingsModule.createSettingsSheet')
        .addItem('設定値更新', 'SettingsModule.updateSettings')
        .addItem('設定をデフォルトに戻す', 'SettingsModule.resetSettingsToDefault')
      )
      .addSeparator()
      .addItem('仕様書更新', 'DocumentationModule.updateSpecificationSheet')
      .addToUi();
    
    LogModule.logOperation("スクリプト初期化", "完了");
  } catch (e) {
    LogModule.handleError(e, 'メニュー作成');
  }
}

/**
 * 全ての分析を順番に実行し、ダッシュボードを初期化する
 */
function analyzeAllData() {
  try {
    LogModule.logOperation("全分析処理", "開始");
    
    // 全分析の実行
    AnalysisModule.analyzeAllData();
    
    LogModule.logOperation("全分析処理", "完了");
    UtilsModule.showToast('全ての分析が完了しました！', 'スキルマップ分析');
  } catch (e) {
    LogModule.handleError(e, '全分析処理');
  }
}