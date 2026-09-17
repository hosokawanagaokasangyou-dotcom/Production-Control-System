package jp.co.pm.ai.desktop.kouchin;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.file.Files;
import java.nio.file.Path;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import jp.co.pm.ai.kouchin.verify.KouchinDiscovery;
import jp.co.pm.ai.kouchin.verify.Source3TargetMonthCheck;

class KouchinSource3MonthNoteUiTest {

    @TempDir
    Path tmp;

    @Test
    @DisplayName("対象月なしの③備考は警告CSS")
    void noteCssIsWarnWhenSource3LacksTargetMonth() throws Exception {
        Path file = tmp.resolve("依頼NO別問合せ.xlsx");
        Files.write(file, new byte[] {1});
        KouchinDiscovery.Row warn = new KouchinDiscovery.Row(
                KouchinDiscovery.ROLE_3,
                file.getFileName().toString(),
                file.toString(),
                "2026年7月度",
                false,
                Source3TargetMonthCheck.warningNote(null, "加工金額行が0件"));
        KouchinVerifyTabController.DiscoveryLine line =
                KouchinVerifyTabController.DiscoveryLine.of("国分/湖南共通", warn);
        assertEquals(Source3TargetMonthCheck.NOTE_CSS, line.getNoteCss());
        assertTrue(line.getNote().contains("データがありません"), line.getNote());

        KouchinDiscovery.Row ok = new KouchinDiscovery.Row(
                KouchinDiscovery.ROLE_3, file.getFileName().toString(), file.toString(),
                "2026年7月度", false, "");
        assertEquals("", KouchinVerifyTabController.DiscoveryLine.of("国分/湖南共通", ok).getNoteCss());
    }

    @Test
    @DisplayName("備考列は警告時に赤背景白太字")
    void noteColumnWiresWarnStyle() throws Exception {
        String src = Files.readString(Path.of(
                "src/main/java/jp/co/pm/ai/desktop/kouchin/KouchinVerifyTabController.java"));
        int start = src.indexOf("private void setupDiscoveryTable(");
        assertTrue(start >= 0);
        int end = src.indexOf("private void addAccessColumn(", start);
        String method = src.substring(start, end);
        assertTrue(method.contains("getNoteCss"), method);
        assertTrue(method.contains("pm-kouchin-note-warn") || method.contains("getNoteCss()"), method);
        String css = Files.readString(Path.of("src/main/resources/jp/co/pm/ai/desktop/css/pm-ai-desktop.css"));
        assertTrue(css.contains(".table-cell.pm-kouchin-note-warn"), css);
        assertTrue(css.contains("#ffffff") || css.contains("#FFFFFF") || css.contains("white"), css);
        int cssAt = css.indexOf(".table-cell.pm-kouchin-note-warn");
        String block = css.substring(cssAt, css.indexOf('}', cssAt) + 1);
        assertTrue(block.contains("-fx-font-weight: bold"), block);
        assertTrue(block.contains("-fx-text-fill"), block);
        assertTrue(block.contains("-fx-background-color"), block);
    }
}
