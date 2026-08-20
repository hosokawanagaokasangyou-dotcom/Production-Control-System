package jp.co.pm.ai.desktop.ui;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.Set;

import org.junit.jupiter.api.Test;

class FactorySiteSwitchBusySupportTest {

    @Test
    void keepBusyDialogForPostSwitchTabLoad_onlyWhenFactorySwitchLoadStarted() {
        assertTrue(FactorySiteSwitchBusySupport.keepBusyDialogForPostSwitchTabLoad(false, true));
        assertFalse(FactorySiteSwitchBusySupport.keepBusyDialogForPostSwitchTabLoad(true, true));
        assertFalse(FactorySiteSwitchBusySupport.keepBusyDialogForPostSwitchTabLoad(false, false));
        assertFalse(FactorySiteSwitchBusySupport.keepBusyDialogForPostSwitchTabLoad(true, false));
    }

    @Test
    void keepBusyVisibleThroughFinish_whenNotStartupAndNoOperatorDialog() {
        assertTrue(FactorySiteSwitchBusySupport.keepBusyVisibleThroughFinish(false, false));
        assertFalse(
                FactorySiteSwitchBusySupport.keepBusyVisibleThroughFinish(false, true),
                "操作者ダイアログと重ねない");
        assertFalse(
                FactorySiteSwitchBusySupport.keepBusyVisibleThroughFinish(true, false),
                "起動シーケンス中は起動側ダイアログに任せる");
        assertFalse(FactorySiteSwitchBusySupport.keepBusyVisibleThroughFinish(true, true));
    }

    @Test
    void finishFactorySiteSwitch_doesNotCloseBusyBeforePostSwitchWork() throws Exception {
        Path java = Path.of("src/main/java/jp/co/pm/ai/desktop/MainShellController.java");
        String text = Files.readString(java, StandardCharsets.UTF_8);
        int start = text.indexOf("private void finishFactorySiteSwitch");
        assertTrue(start >= 0, "finishFactorySiteSwitch が見つからない");
        int next = text.indexOf("\n    @Override\n    public Map<String, String> snapshotUiEnv", start);
        String body = next > start ? text.substring(start, next) : text.substring(start, start + 2200);
        assertTrue(body.contains("keepBusyVisibleThroughFinish"), "同一進捗を維持する方針を使う");
        assertFalse(
                body.contains("endFactorySiteSwitchBusy();\n        factorySiteSwitchInProgress = false;"),
                "切替完了直後に進捗ダイアログを閉じない");
    }

    @Test
    void workUnits_haveDistinctStatusesForPaintedPulses() {
        assertTrue(FactorySiteSwitchBusySupport.WORK_UNIT_COUNT >= 9);
        Set<String> statuses = new HashSet<>();
        for (int i = 0; i < FactorySiteSwitchBusySupport.WORK_UNIT_COUNT; i++) {
            String status = FactorySiteSwitchBusySupport.statusForWorkUnit(i);
            assertTrue(status != null && !status.isBlank(), "work unit " + i);
            statuses.add(status);
        }
        assertTrue(statuses.size() >= 8, "パルスごとに状況文言を変えてスピナーを回す");
        assertEquals(
                FactorySiteSwitchBusyDialog.STATUS_CONNECT,
                FactorySiteSwitchBusySupport.statusForWorkUnit(FactorySiteSwitchBusySupport.WORK_CONNECT));
        assertEquals(
                FactorySiteSwitchBusyDialog.STATUS_REFRESH_REQUEST_FORM,
                FactorySiteSwitchBusySupport.statusForWorkUnit(
                        FactorySiteSwitchBusySupport.WORK_REFRESH_REQUEST_FORM));
        assertNotEquals(
                FactorySiteSwitchBusySupport.statusForWorkUnit(
                        FactorySiteSwitchBusySupport.WORK_REFRESH_REQUEST_FORM),
                FactorySiteSwitchBusySupport.statusForWorkUnit(
                        FactorySiteSwitchBusySupport.WORK_REFRESH_PIPELINE));
    }

    @Test
    void postSwitchWorkUnits_splitAttendanceReloads() {
        assertTrue(FactorySiteSwitchBusySupport.POST_WORK_COUNT >= 5);
        assertEquals(
                FactorySiteSwitchBusyDialog.STATUS_ATTENDANCE_COMPANY,
                FactorySiteSwitchBusySupport.statusForPostSwitchWork(
                        FactorySiteSwitchBusySupport.POST_ATTENDANCE_COMPANY));
        assertEquals(
                FactorySiteSwitchBusyDialog.STATUS_ATTENDANCE_MEMBER,
                FactorySiteSwitchBusySupport.statusForPostSwitchWork(
                        FactorySiteSwitchBusySupport.POST_ATTENDANCE_MEMBER));
        assertEquals(
                FactorySiteSwitchBusyDialog.STATUS_ATTENDANCE_MACHINE,
                FactorySiteSwitchBusySupport.statusForPostSwitchWork(
                        FactorySiteSwitchBusySupport.POST_ATTENDANCE_MACHINE));
        assertEquals(
                FactorySiteSwitchBusyDialog.STATUS_ATTENDANCE_MASTER,
                FactorySiteSwitchBusySupport.statusForPostSwitchWork(
                        FactorySiteSwitchBusySupport.POST_ATTENDANCE_MASTER));
        assertEquals(
                FactorySiteSwitchBusyDialog.STATUS_BACKGROUND_LOAD,
                FactorySiteSwitchBusySupport.statusForPostSwitchWork(
                        FactorySiteSwitchBusySupport.POST_BACKGROUND_LOAD));
    }

    @Test
    void mainShell_pulsesStatusBeforeEachFactorySwitchWorkUnit() throws Exception {
        Path java = Path.of("src/main/java/jp/co/pm/ai/desktop/MainShellController.java");
        String text = Files.readString(java, StandardCharsets.UTF_8);
        assertTrue(text.contains("statusForWorkUnit"), "切替ステップの状況文言カタログを使う");
        assertTrue(text.contains("statusForPostSwitchWork"), "勤怠再読込も細かい進捗タスクにする");
        int finish = text.indexOf("private void finishFactorySiteSwitch");
        assertTrue(finish >= 0);
        int next = text.indexOf("\n    @Override\n    public Map<String, String> snapshotUiEnv", finish);
        String body = next > finish ? text.substring(finish, next) : text.substring(finish, finish + 2500);
        assertFalse(
                body.contains("reloadAttendanceTabsFromJson(true)"),
                "勤怠再読込を一括で走らせてダイアログを止めない");
    }

    @Test
    void resolveTabLoadStatus_usesMessageOrDefault() {
        assertEquals(
                FactorySiteSwitchBusyDialog.STATUS_BACKGROUND_LOAD,
                FactorySiteSwitchBusySupport.resolveTabLoadStatus(""));
        assertEquals(
                FactorySiteSwitchBusyDialog.STATUS_BACKGROUND_LOAD,
                FactorySiteSwitchBusySupport.resolveTabLoadStatus(null));
        assertEquals(
                "起動後読込 (1/6): リモートデスクトップ…",
                FactorySiteSwitchBusySupport.resolveTabLoadStatus("起動後読込 (1/6): リモートデスクトップ…"));
    }

    @Test
    void centerX_centersChildOverOwner() {
        assertEquals(400.0, FactorySiteSwitchBusySupport.centerX(100.0, 800.0, 200.0), 0.001);
        assertEquals(240.0, FactorySiteSwitchBusySupport.centerY(40.0, 600.0, 200.0), 0.001);
    }

    @Test
    void realizeStageForImmediateShow_noOpWhenStageOrSceneMissing() {
        FactorySiteSwitchBusySupport.realizeStageForImmediateShow(null);
    }
}
