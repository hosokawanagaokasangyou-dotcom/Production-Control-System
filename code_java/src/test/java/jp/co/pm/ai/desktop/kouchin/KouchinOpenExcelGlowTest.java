package jp.co.pm.ai.desktop.kouchin;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import jp.co.pm.ai.desktop.config.FactorySite;
import jp.co.pm.ai.kouchin.verify.DualWriteFiles;
import jp.co.pm.ai.kouchin.verify.VerifyRunSupport;

class KouchinOpenExcelGlowTest {

    @Test
    @DisplayName("検証Excelが書けていれば開くボタンを光らせる")
    void glowsWhenXlsxSucceeded() {
        VerifyRunSupport.Written written = new VerifyRunSupport.Written(
                new DualWriteFiles.WriteOutcome(
                        List.of(Path.of("検証結果_国分工場_20260916.xlsx")), List.of()),
                new DualWriteFiles.WriteOutcome(List.of(), List.of()),
                new DualWriteFiles.WriteOutcome(List.of(), List.of()),
                "20260916_000000",
                "報告メール_統合_20260916_000000");
        assertTrue(KouchinVerifyTabController.shouldGlowOpenExcel(written, FactorySite.KOKUBU));
    }

    @Test
    @DisplayName("Excel未書込・失敗・未実行では光らせない")
    void doesNotGlowWithoutOpenableExcel() {
        assertFalse(KouchinVerifyTabController.shouldGlowOpenExcel(null, FactorySite.KOKUBU));
        VerifyRunSupport.Written empty = new VerifyRunSupport.Written(
                new DualWriteFiles.WriteOutcome(List.of(), List.of("write failed")),
                new DualWriteFiles.WriteOutcome(List.of(), List.of()),
                new DualWriteFiles.WriteOutcome(List.of(), List.of()),
                "20260916_000000",
                "報告メール_統合_20260916_000000");
        assertFalse(KouchinVerifyTabController.shouldGlowOpenExcel(empty, FactorySite.KOKUBU));
        VerifyRunSupport.Written noXlsx = new VerifyRunSupport.Written(
                null,
                new DualWriteFiles.WriteOutcome(List.of(), List.of()),
                new DualWriteFiles.WriteOutcome(List.of(), List.of()),
                "20260916_000000",
                "報告メール_統合_20260916_000000");
        assertFalse(KouchinVerifyTabController.shouldGlowOpenExcel(noXlsx, FactorySite.KONAN));
        VerifyRunSupport.Written konanOnly = new VerifyRunSupport.Written(
                new DualWriteFiles.WriteOutcome(
                        List.of(Path.of("検証結果_湖南工場_20260916.xlsx")), List.of()),
                new DualWriteFiles.WriteOutcome(List.of(), List.of()),
                new DualWriteFiles.WriteOutcome(List.of(), List.of()),
                "20260916_000000",
                "報告メール_統合_20260916_000000");
        assertTrue(KouchinVerifyTabController.shouldGlowOpenExcel(konanOnly, FactorySite.KONAN));
        assertFalse(KouchinVerifyTabController.shouldGlowOpenExcel(konanOnly, FactorySite.KOKUBU));
    }

    @Test
    @DisplayName("国分・湖南のExcelを開くは読み取り専用")
    void openExcelForUsesReadOnly() throws Exception {
        String src = Files.readString(Path.of(
                "src/main/java/jp/co/pm/ai/desktop/kouchin/KouchinVerifyTabController.java"));
        int start = src.indexOf("private void openExcelFor(");
        assertTrue(start >= 0, "openExcelFor が無い");
        int end = src.indexOf("private void onOpenFolder(", start);
        assertTrue(end > start, "onOpenFolder が無い");
        String method = src.substring(start, end);
        assertTrue(method.contains("openFileReadOnly"), method);
        assertFalse(method.contains("DesktopFileOpener.openFile("), method);
    }

    @Test
    @DisplayName("検証未了かつ実行可なら検証ボタンを光らせる")
    void glowsRunWhenUnverifiedAndEnabled() {
        assertTrue(KouchinVerifyTabController.shouldGlowRunButton(true, false));
        assertFalse(KouchinVerifyTabController.shouldGlowRunButton(false, false));
        assertFalse(KouchinVerifyTabController.shouldGlowRunButton(true, true));
        assertTrue(KouchinVerifyTabController.shouldGlowRunBoth(true, false, false));
        assertTrue(KouchinVerifyTabController.shouldGlowRunBoth(true, true, false));
        assertTrue(KouchinVerifyTabController.shouldGlowRunBoth(true, false, true));
        assertFalse(KouchinVerifyTabController.shouldGlowRunBoth(true, true, true));
        assertFalse(KouchinVerifyTabController.shouldGlowRunBoth(false, false, false));
    }
}
