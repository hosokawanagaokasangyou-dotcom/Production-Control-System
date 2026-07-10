package jp.co.pm.ai.desktop.print;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.time.LocalDate;
import java.util.List;

import javafx.application.Platform;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;

import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

/**
 * 印刷（PDF出力）時のレイアウト・ページ分割の回帰防止テスト。
 */
class OperatorCardPreviewFactoryTest {

    @BeforeAll
    static void initJavaFx() {
        try {
            Platform.startup(() -> {});
        } catch (IllegalStateException ignored) {
            // already started
        }
    }

    private static OperatorCardPage samplePage() {
        OperatorCardTaskRow row =
                new OperatorCardTaskRow("08:50-09:00", "巻返し", "スリット機1 湖南", "TPI06-01", "194", "194", "-");
        OperatorCardDaySection day =
                new OperatorCardDaySection(LocalDate.of(2026, 7, 10), "07/10", List.of(row));
        return new OperatorCardPage("図司 智子", List.of(day));
    }

    @Test
    void buildRootWithoutWidthArgUsesA4PrefWidthForScreenPreview() {
        Parent root = OperatorCardPreviewFactory.buildRoot(samplePage(), "SansSerif");
        assertTrue(root instanceof Region);
        assertEquals(
                OperatorCardPreviewFactory.A4_PREF_WIDTH, ((Region) root).getPrefWidth(), 0.001);
    }

    @Test
    void printScaleShrinksLayoutWidthToPrintableArea() {
        double printableWidth = 555.0;
        double scale = OperatorCardPrintCompositor.printScaleForWidth(printableWidth);
        assertTrue(scale < 1.0);
        assertEquals(printableWidth, OperatorCardPreviewFactory.A4_PREF_WIDTH * scale, 0.5);
    }

    @Test
    void scaledPrintPageKeepsFullColumnHeadersAtPrintableWidth() {
        double printableWidth = 555.0;
        VBox layoutRoot =
                (VBox) OperatorCardPreviewFactory.buildRoot(samplePage(), "SansSerif");
        Parent printRoot =
                OperatorCardPrintCompositor.wrapScaledPrintPage(layoutRoot, printableWidth);
        OperatorCardPrintCompositor.createPrintScene(printRoot, printableWidth, 800.0);

        GridPane grid = findFirstGrid(layoutRoot);
        assertTrue(grid != null, "day grid should be present");

        List<String> headerTexts =
                grid.getChildren().stream()
                        .filter(n -> GridPane.getRowIndex(n) == null || GridPane.getRowIndex(n) == 0)
                        .filter(n -> n instanceof Label)
                        .map(n -> ((Label) n).getText())
                        .toList();
        assertEquals(
                List.of("時間帯", "工程", "機械", "依頼NO", "当日配台", "換算", "メンバー"), headerTexts);
        for (String hdr : headerTexts) {
            assertTrue(!hdr.endsWith("…") && !hdr.endsWith("..."), "ヘッダーが途中で切れている: " + hdr);
        }
    }

    private static OperatorCardDaySection dayWithRows(LocalDate date, int rowCount) {
        List<OperatorCardTaskRow> rows = new java.util.ArrayList<>();
        for (int i = 0; i < rowCount; i++) {
            rows.add(
                    new OperatorCardTaskRow(
                            "08:%02d-09:%02d".formatted(i, i),
                            "工程" + i,
                            "機械" + i,
                            "NO-" + i,
                            "100",
                            "100",
                            "メンバー" + i));
        }
        return new OperatorCardDaySection(date, date.toString(), rows);
    }

    private static OperatorCardDaySection emptyDay(LocalDate date) {
        return new OperatorCardDaySection(date, date.toString(), List.of());
    }

    @Test
    void buildPrintPagesReturnsSinglePageWhenEverythingFits() {
        OperatorCardPage page =
                new OperatorCardPage(
                        "図司 智子",
                        List.of(
                                dayWithRows(LocalDate.of(2026, 7, 10), 2),
                                dayWithRows(LocalDate.of(2026, 7, 11), 1)));
        List<Parent> pages =
                OperatorCardPreviewFactory.buildPrintPages(page, "SansSerif", 555.0, 5000.0);
        assertEquals(1, pages.size());
        assertEquals(2, totalDayBoxCount(pages));
    }

    @Test
    void buildPrintPagesSplitsAcrossMultiplePagesWhenDaysOverflowPrintableHeight() {
        OperatorCardPage page =
                new OperatorCardPage(
                        "図司 智子",
                        List.of(
                                dayWithRows(LocalDate.of(2026, 7, 10), 6),
                                dayWithRows(LocalDate.of(2026, 7, 11), 6),
                                dayWithRows(LocalDate.of(2026, 7, 12), 6),
                                dayWithRows(LocalDate.of(2026, 7, 13), 6),
                                dayWithRows(LocalDate.of(2026, 7, 14), 6),
                                dayWithRows(LocalDate.of(2026, 7, 15), 6)));
        double printableWidth = 555.0;
        double printableHeight = 500.0;

        List<Parent> pages =
                OperatorCardPreviewFactory.buildPrintPages(
                        page, "SansSerif", printableWidth, printableHeight);

        assertTrue(pages.size() > 1, "6日分・各6行は1ページに収まらず複数ページへ分割されるはず");
        assertEquals(6, totalDayBoxCount(pages), "分割後もページ全体で日数の合計は変わらない（欠落しない）");
    }

    @Test
    void buildPrintPagesPreservesAllDaysInRealisticSixDayPattern() {
        OperatorCardPage page =
                new OperatorCardPage(
                        "図司 智子",
                        List.of(
                                dayWithRows(LocalDate.of(2026, 7, 10), 4),
                                emptyDay(LocalDate.of(2026, 7, 11)),
                                emptyDay(LocalDate.of(2026, 7, 12)),
                                dayWithRows(LocalDate.of(2026, 7, 13), 6),
                                emptyDay(LocalDate.of(2026, 7, 14)),
                                dayWithRows(LocalDate.of(2026, 7, 15), 4)));
        double printableWidth = 523.0;
        double printableHeight = 750.0;

        List<Parent> pages =
                OperatorCardPreviewFactory.buildPrintPages(
                        page, "SansSerif", printableWidth, printableHeight);

        assertEquals(6, totalDayBoxCount(pages), "6日分すべてが印刷ページに含まれる");
    }

    @Test
    void buildPrintPagesFallsBackToSinglePageWhenPrintableHeightInvalid() {
        OperatorCardPage page =
                new OperatorCardPage("図司 智子", List.of(dayWithRows(LocalDate.of(2026, 7, 10), 3)));
        List<Parent> pages =
                OperatorCardPreviewFactory.buildPrintPages(page, "SansSerif", 555.0, 0.0);
        assertEquals(1, pages.size());
    }

    private static int totalDayBoxCount(List<Parent> pages) {
        int total = 0;
        for (Parent p : pages) {
            total += dayBoxCount((VBox) p);
        }
        return total;
    }

    private static int dayBoxCount(VBox root) {
        return Math.max(0, root.getChildren().size() - 4);
    }

    private static GridPane findFirstGrid(javafx.scene.Parent parent) {
        for (javafx.scene.Node child : parent.getChildrenUnmodifiable()) {
            if (child instanceof GridPane gp) {
                return gp;
            }
            if (child instanceof javafx.scene.Parent p) {
                GridPane found = findFirstGrid(p);
                if (found != null) {
                    return found;
                }
            }
        }
        return null;
    }
}
