package jp.co.pm.ai.desktop;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

class OperatorActionLogTabControllerTest {

    @Test
    void actionAndResultLabels_useJapanese() {
        assertEquals("段階2完了", OperatorActionLogTabController.actionLabel("stage2_complete"));
        assertEquals("同一化チェック", OperatorActionLogTabController.actionLabel("identity_check"));
        assertEquals("Excel出力", OperatorActionLogTabController.actionLabel("excel_export"));
        assertEquals("終了警告", OperatorActionLogTabController.actionLabel("close_warning"));
        assertEquals("セッション開始", OperatorActionLogTabController.actionLabel("session_start"));
        assertEquals("タブ選択", OperatorActionLogTabController.actionLabel("tab_select"));
        assertEquals("東レCSVドロップ", OperatorActionLogTabController.actionLabel("kouchin_drop"));
        assertEquals("東レ後加工賃検証", OperatorActionLogTabController.actionLabel("kouchin_verify"));
        assertEquals(
                "工場共有バックアップ",
                OperatorActionLogTabController.actionLabel("factory_share_backup"));
        assertEquals("段階1完了", OperatorActionLogTabController.actionLabel("stage1_complete"));
        assertEquals(
                "手動配台セル編集",
                OperatorActionLogTabController.actionLabel("dispatch_cell_edit"));
        assertEquals("環境変数変更", OperatorActionLogTabController.actionLabel("env_change"));
        assertEquals(
                "加工トレンド実行",
                OperatorActionLogTabController.actionLabel("processing_trend_run"));
        assertEquals("開発", OperatorActionLogTabController.featureLabel("developer"));
        assertEquals("環境変数", OperatorActionLogTabController.featureLabel("env"));
        assertEquals(
                "配台計画_タスク入力",
                OperatorActionLogTabController.featureLabel("planInput"));
        assertEquals("加工トレンド", OperatorActionLogTabController.featureLabel("processingTrend"));
        assertEquals("成功", OperatorActionLogTabController.resultLabel("ok"));
        assertEquals("空", OperatorActionLogTabController.resultLabel("empty"));
        assertEquals("警告", OperatorActionLogTabController.resultLabel("warn"));
        assertEquals("差異", OperatorActionLogTabController.resultLabel("mismatch"));
        assertEquals("失敗", OperatorActionLogTabController.resultLabel("error"));
        assertEquals("表示", OperatorActionLogTabController.resultLabel("shown"));
    }

    @Test
    void formatTs_usesJapaneseLocalDateTime() {
        assertEquals(
                "2026-08-17 20:45",
                OperatorActionLogTabController.formatTs("2026-08-17T20:45:12+09:00"));
        assertEquals("", OperatorActionLogTabController.formatTs(""));
        assertEquals("not-a-date", OperatorActionLogTabController.formatTs("not-a-date"));
    }
}
