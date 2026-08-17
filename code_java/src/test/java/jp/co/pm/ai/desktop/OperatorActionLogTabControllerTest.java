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
        assertEquals("成功", OperatorActionLogTabController.resultLabel("ok"));
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
