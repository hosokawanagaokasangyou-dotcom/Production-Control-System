package jp.co.pm.ai.desktop.print;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.util.List;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.desktop.config.FactorySite;

class FactoryOperatorUserPdfExporterTest {

    @Test
    void export_writesPdfHeader(@TempDir Path tmp) throws Exception {
        assumeTrue(
                Files.isRegularFile(Path.of("/mnt/c/Windows/Fonts/msgothic.ttc"))
                        || Files.isRegularFile(Path.of("/usr/share/fonts/truetype/fonts-japanese-gothic.ttf")),
                "日本語 PDF フォントが無いためスキップ");
        Path out = tmp.resolve("users.pdf");
        FactoryOperatorUserPdfExporter.export(
                out,
                FactorySite.KONAN,
                List.of(
                        new FactoryOperatorUserPdfExporter.Row("砂田", "設定済", "1234"),
                        new FactoryOperatorUserPdfExporter.Row("古家", "初回変更待", "5678")),
                "管理者",
                Instant.parse("2026-05-28T12:00:00Z"),
                tmp.resolve("factory-operator-users.bin").toString());
        assertTrue(Files.isRegularFile(out));
        assertTrue(Files.size(out) > 128);
        String head = Files.readString(out, StandardCharsets.ISO_8859_1);
        assertTrue(head.startsWith("%PDF"));
    }

    @Test
    void resolveOutputPath_usesFactorySuffix(@TempDir Path tmp) {
        Path expected = tmp.resolve("factory-operator-users-KOKUBU.pdf").toAbsolutePath().normalize();
        assertEquals(
                expected,
                FactoryOperatorUserPdfExporter.resolveOutputPath(
                        java.util.Map.of(
                                jp.co.pm.ai.desktop.config.AppPaths.KEY_PM_AI_SUMMARY_AI_DISPATCH_WORKBOOK,
                                tmp.resolve("summary.xlsx").toString()),
                        FactorySite.KOKUBU));
    }
}
