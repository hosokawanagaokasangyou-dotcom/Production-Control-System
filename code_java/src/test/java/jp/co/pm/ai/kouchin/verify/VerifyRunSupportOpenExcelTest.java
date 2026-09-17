package jp.co.pm.ai.kouchin.verify;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Path;
import java.util.List;

import jp.co.pm.ai.desktop.config.FactorySite;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

class VerifyRunSupportOpenExcelTest {

    @Test
    @DisplayName("まとめて検証は国分と湖南のExcelを重複なく両方返す")
    void excelFilesToOpenIncludesBothFactoriesOnce() {
        Path kokubuDir = Path.of("C:/out/kokubu");
        Path konanDir = Path.of("C:/out/konan");
        DualWriteFiles.WriteOutcome xlsx = new DualWriteFiles.WriteOutcome(
                List.of(
                        kokubuDir.resolve("検証結果_国分工場_20260916_174000.xlsx"),
                        konanDir.resolve("検証結果_国分工場_20260916_174000.xlsx"),
                        kokubuDir.resolve("検証結果_湖南工場_20260916_174000.xlsx"),
                        konanDir.resolve("検証結果_湖南工場_20260916_174000.xlsx")),
                List.of());
        VerifyRunSupport.Written written = written(xlsx);
        List<Path> open = VerifyRunSupport.excelFilesToOpen(written);
        assertEquals(2, open.size(), open.toString());
        assertTrue(open.stream().anyMatch(p -> p.getFileName().toString().contains("国分工場")), open.toString());
        assertTrue(open.stream().anyMatch(p -> p.getFileName().toString().contains("湖南工場")), open.toString());
    }

    @Test
    @DisplayName("片工場だけの検証は1ファイルだけ")
    void excelFilesToOpenSingleFactory() {
        DualWriteFiles.WriteOutcome xlsx = new DualWriteFiles.WriteOutcome(
                List.of(
                        Path.of("C:/out/kokubu/検証結果_国分工場_1.xlsx"),
                        Path.of("C:/out/konan/検証結果_国分工場_1.xlsx")),
                List.of());
        List<Path> open = VerifyRunSupport.excelFilesToOpen(written(xlsx));
        assertEquals(1, open.size());
        assertEquals("検証結果_国分工場_1.xlsx", open.get(0).getFileName().toString());
    }

    @Test
    @DisplayName("工場指定は自工場のExcelだけを返す")
    void excelFileToOpenIsFactorySpecific() {
        DualWriteFiles.WriteOutcome xlsx = new DualWriteFiles.WriteOutcome(
                List.of(
                        Path.of("C:/out/kokubu/検証結果_国分工場_1.xlsx"),
                        Path.of("C:/out/konan/検証結果_国分工場_1.xlsx"),
                        Path.of("C:/out/kokubu/検証結果_湖南工場_1.xlsx"),
                        Path.of("C:/out/konan/検証結果_湖南工場_1.xlsx")),
                List.of());
        VerifyRunSupport.Written written = written(xlsx);
        Path kokubu = VerifyRunSupport.excelFileToOpen(written, FactorySite.KOKUBU);
        Path konan = VerifyRunSupport.excelFileToOpen(written, FactorySite.KONAN);
        assertEquals("検証結果_国分工場_1.xlsx", kokubu.getFileName().toString());
        assertEquals("検証結果_湖南工場_1.xlsx", konan.getFileName().toString());
    }

    @Test
    @DisplayName("他工場のExcelしか無いときは自工場指定は空")
    void excelFileToOpenDoesNotFallBackToOtherFactory() {
        DualWriteFiles.WriteOutcome xlsx = new DualWriteFiles.WriteOutcome(
                List.of(Path.of("C:/out/kokubu/検証結果_湖南工場_1.xlsx")),
                List.of());
        VerifyRunSupport.Written written = written(xlsx);
        assertTrue(VerifyRunSupport.excelFileToOpen(written, FactorySite.KOKUBU) == null);
        assertEquals("検証結果_湖南工場_1.xlsx",
                VerifyRunSupport.excelFileToOpen(written, FactorySite.KONAN).getFileName().toString());
    }

    @Test
    @DisplayName("Excel未書込は空")
    void excelFilesToOpenEmpty() {
        assertTrue(VerifyRunSupport.excelFilesToOpen(null).isEmpty());
        assertTrue(VerifyRunSupport.excelFilesToOpen(written(null)).isEmpty());
        assertTrue(VerifyRunSupport.excelFilesToOpen(
                written(new DualWriteFiles.WriteOutcome(List.of(), List.of("fail")))).isEmpty());
    }

    @Test
    @DisplayName("片工場のBothResultでunifiedMailがnullでもメール本文を組み立てる")
    void mailTextToWriteFallsBackWhenUnifiedMailNull() {
        String text = VerifyRunSupport.mailTextToWrite(
                new BothResult(null, null, null, null, null), null, null);
        assertTrue(text != null && !text.isBlank(), text);
        assertTrue(text.contains("未検証"), text);
        assertEquals("固定本文", VerifyRunSupport.mailTextToWrite(
                new BothResult(null, null, null, null, "固定本文"), null, null));
        String fromNullBoth = VerifyRunSupport.mailTextToWrite(null, null, null);
        assertTrue(fromNullBoth.contains("未検証"), fromNullBoth);
    }

    private static VerifyRunSupport.Written written(DualWriteFiles.WriteOutcome xlsx) {
        return new VerifyRunSupport.Written(
                xlsx,
                new DualWriteFiles.WriteOutcome(List.of(), List.of()),
                new DualWriteFiles.WriteOutcome(List.of(), List.of()),
                "20260916_174000",
                "報告メール_統合_20260916_174000");
    }
}
