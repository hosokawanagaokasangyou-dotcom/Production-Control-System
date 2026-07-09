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
 * 印刷（PDF出力）時に、A4 の可印刷幅（{@code javafx.print.PageLayout#getPrintableWidth()}
 * 相当）へルート幅を合わせないと右側の列（換算・メンバー）がクリップされる不具合の再発防止テスト。
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
    void buildRootHonorsExplicitPrintableWidthForPdfPrinting() {
        double printableWidth = 555.0; // A4 可印刷幅相当（DEFAULT マージン後）
        Parent root = OperatorCardPreviewFactory.buildRoot(samplePage(), "SansSerif", printableWidth);
        assertEquals(printableWidth, ((Region) root).getPrefWidth(), 0.001);
        assertTrue(
                printableWidth < OperatorCardPreviewFactory.A4_PREF_WIDTH,
                "可印刷幅は画面プレビュー固定幅より狭いことを前提にした回帰テスト");
    }

    @Test
    void allSevenColumnsRemainVisibleAtPrintableWidthNoRightSideClipping() {
        double printableWidth = 555.0;
        VBox root =
                (VBox)
                        OperatorCardPreviewFactory.buildRoot(
                                samplePage(), "SansSerif", printableWidth);
        new Scene(root, printableWidth, root.prefHeight(printableWidth), Color.WHITE);
        root.applyCss();
        root.layout();

        GridPane grid = findFirstGrid(root);
        assertTrue(grid != null, "day grid should be present");

        // ヘッダー行（row 0）に「換算」「メンバー」を含む全7列が存在し、
        // グリッド全体の実測幅がルート幅（＝可印刷幅）を超えないこと。
        List<String> headerTexts =
                grid.getChildren().stream()
                        .filter(n -> GridPane.getRowIndex(n) == null || GridPane.getRowIndex(n) == 0)
                        .filter(n -> n instanceof Label)
                        .map(n -> ((Label) n).getText())
                        .toList();
        assertEquals(
                List.of("時間帯", "工程", "機械", "依頼NO", "当日配台", "換算", "メンバー"), headerTexts);

        grid.applyCss();
        grid.layout();
        assertTrue(
                grid.getWidth() <= printableWidth + 0.5,
                "グリッド幅（" + grid.getWidth() + "）が可印刷幅（" + printableWidth + "）を超えている");
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
