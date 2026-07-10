package jp.co.pm.ai.desktop.print;

import java.net.URL;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

import javafx.geometry.HPos;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.Separator;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.Region;
import javafx.scene.layout.RowConstraints;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;

/** Builds the printable / preview {@link Parent} for one {@link OperatorCardPage}. */
public final class OperatorCardPreviewFactory {

    private static final String[] JP_WEEKDAY_SHORT =
            new String[] {
                "月",
                "火",
                "水",
                "木",
                "金",
                "土",
                "日"
            };

    /** Approximate A4 width at 96 dpi for layout pref widths (210 mm). */
    public static final double A4_PREF_WIDTH = 794;

    /** Approximate A4 height at 96 dpi (297 mm). */
    public static final double A4_PREF_HEIGHT = 1123;

    private OperatorCardPreviewFactory() {}

    public static Parent buildRoot(OperatorCardPage page, String fontFamily) {
        return buildRoot(page, fontFamily, A4_PREF_WIDTH);
    }

    /**
     * 印刷用に、用紙の可印刷幅（{@code javafx.print.PageLayout#getPrintableWidth()} 等）へ
     * ルート幅を合わせて組み立てる。画面プレビュー用。印刷は {@link #buildPrintPages} が
     * {@link #A4_PREF_WIDTH} で組み立て、{@link OperatorCardPrintCompositor} がスケールする。
     */
    public static Parent buildRoot(OperatorCardPage page, String fontFamily, double rootWidth) {
        return assembleRoot(page.operatorName(), page.days(), fontFamily, rootWidth);
    }

    /**
     * 印刷用に、1 オペレーター分のカードを用紙の可印刷高さへ収まる単位で複数ページに分割する。
     *
     * <p>レイアウトは画面プレビューと同じ {@link #A4_PREF_WIDTH} で組み立て、高さ判定は
     * {@code printableHeight / printScale}（スケール前のレイアウト高さ）で行う。
     * 1 日の行数が多い場合は行単位でも分割する。
     */
    public static List<Parent> buildPrintPages(
            OperatorCardPage page, String fontFamily, double printableWidth, double printableHeight) {
        List<OperatorCardDaySection> days = page.days();
        if (days.isEmpty()) {
            return List.of(assembleRoot(page.operatorName(), days, fontFamily, A4_PREF_WIDTH));
        }
        if (!Double.isFinite(printableWidth)
                || printableWidth <= 0
                || !Double.isFinite(printableHeight)
                || printableHeight <= 0) {
            return List.of(assembleRoot(page.operatorName(), days, fontFamily, A4_PREF_WIDTH));
        }

        double scale = OperatorCardPrintCompositor.printScaleForWidth(printableWidth);
        double layoutHeightBudget = printableHeight / scale;
        double pageHeaderHeight =
                measuredLayoutHeight(
                        assembleRoot(page.operatorName(), List.of(), fontFamily, A4_PREF_WIDTH),
                        A4_PREF_WIDTH);
        String ff = cssFontFamily(fontFamily);

        List<Parent> pages = new ArrayList<>();
        VBox current = assembleRoot(page.operatorName(), List.of(), fontFamily, A4_PREF_WIDTH);

        for (OperatorCardDaySection day : days) {
            for (VBox dayBox :
                    splitDayBoxesForPrint(day, ff, layoutHeightBudget, pageHeaderHeight)) {
                current.getChildren().add(dayBox);
                double height = measuredLayoutHeight(current, A4_PREF_WIDTH);
                boolean overflows = height > layoutHeightBudget + 0.5 && dayCountOf(current) > 1;
                if (overflows) {
                    current.getChildren().remove(dayBox);
                    pages.add(current);
                    current = assembleRoot(page.operatorName(), List.of(), fontFamily, A4_PREF_WIDTH);
                    current.getChildren().add(dayBox);
                }
            }
        }
        pages.add(current);
        return pages;
    }

    /**
     * 1 日分を 1 つ以上の {@link VBox}（日見出し＋表）へ分割する。1 ページに収まらない行数の日は
     * 行チャンクへ切り出す（各チャンクは日見出しを繰り返す）。
     */
    private static List<VBox> splitDayBoxesForPrint(
            OperatorCardDaySection day,
            String ff,
            double layoutHeightBudget,
            double pageHeaderHeight) {
        VBox full = buildDayBox(day, ff);
        double maxDayHeight = Math.max(80.0, layoutHeightBudget - pageHeaderHeight);
        if (measuredLayoutHeight(full, A4_PREF_WIDTH) <= maxDayHeight + 0.5) {
            return List.of(full);
        }

        List<OperatorCardTaskRow> rows = day.rows();
        if (rows.isEmpty()) {
            return List.of(full);
        }

        List<VBox> chunks = new ArrayList<>();
        int start = 0;
        while (start < rows.size()) {
            int end = start + 1;
            while (end <= rows.size()) {
                OperatorCardDaySection slice =
                        new OperatorCardDaySection(
                                day.date(),
                                day.dateColumnHeader(),
                                rows.subList(start, end));
                VBox box = buildDayBox(slice, ff);
                double h = measuredLayoutHeight(box, A4_PREF_WIDTH);
                if (h > maxDayHeight + 0.5 && end - start > 1) {
                    end--;
                    slice =
                            new OperatorCardDaySection(
                                    day.date(),
                                    day.dateColumnHeader(),
                                    rows.subList(start, end));
                    chunks.add(buildDayBox(slice, ff));
                    start = end;
                    break;
                }
                if (end == rows.size()) {
                    chunks.add(box);
                    start = end;
                    break;
                }
                end++;
            }
        }
        return chunks.isEmpty() ? List.of(full) : chunks;
    }

    /** CSS 適用後に {@code layout()} まで行い、内容に応じた高さを返す（Scene 高さに引き伸ばされない）。 */
    public static void prepareForLayoutMeasure(VBox root, double layoutWidth) {
        root.setMaxHeight(Region.USE_PREF_SIZE);
        Scene scene = root.getScene();
        if (scene == null) {
            scene = new Scene(root, layoutWidth, 1, Color.WHITE);
            attachDesktopStylesheet(scene);
        } else {
            attachDesktopStylesheet(scene);
        }
        root.applyCss();
        root.layout();
    }

    static double measuredLayoutHeight(VBox root, double layoutWidth) {
        root.setMaxHeight(Region.USE_PREF_SIZE);
        Scene scene = root.getScene();
        boolean tempScene = scene == null;
        if (tempScene) {
            scene = new Scene(root, layoutWidth, 1, Color.WHITE);
            attachDesktopStylesheet(scene);
        } else {
            attachDesktopStylesheet(scene);
        }
        root.applyCss();
        root.layout();
        double h = root.prefHeight(layoutWidth);
        if (h <= 0) {
            h = root.getBoundsInLocal().getHeight();
        }
        if (tempScene) {
            scene.setRoot(new javafx.scene.Group());
        }
        return h;
    }

    private static int dayCountOf(VBox root) {
        // root children: headingRow, Separator, title, Separator, dayBox...
        return Math.max(0, root.getChildren().size() - 4);
    }

    private static VBox assembleRoot(
            String operatorName, List<OperatorCardDaySection> days, String fontFamily, double rootWidth) {
        String ff = cssFontFamily(fontFamily);
        double width = rootWidth > 0 ? rootWidth : A4_PREF_WIDTH;

        VBox root = new VBox(10);
        root.setPadding(new Insets(16, 20, 16, 20));
        root.setPrefWidth(width);
        root.setMinWidth(width);
        root.setMaxWidth(width);
        root.setMaxHeight(Region.USE_PREF_SIZE);
        root.setStyle("-fx-font-family: " + ff + ";");
        root.getStyleClass().add("pm-operator-card-root");

        Label docHeading =
                new Label(
                        "オペレーション"
                                + "カード");
        docHeading.getStyleClass().add("pm-operator-card-doc-title");
        docHeading.setMaxWidth(Double.MAX_VALUE);
        HBox.setHgrow(docHeading, Priority.ALWAYS);

        Label issuedAt = new Label(formatIssuedAt());
        issuedAt.getStyleClass().add("pm-operator-card-issued-at");
        issuedAt.setAlignment(Pos.CENTER_RIGHT);

        HBox headingRow = new HBox();
        headingRow.setAlignment(Pos.CENTER_LEFT);
        headingRow.setSpacing(12);
        headingRow.getChildren().addAll(docHeading, issuedAt);

        Label title = new Label(operatorName);
        title.getStyleClass().add("pm-operator-card-title");
        title.setMaxWidth(Double.MAX_VALUE);
        title.setAlignment(Pos.CENTER_LEFT);

        root.getChildren().add(headingRow);
        root.getChildren().add(new Separator());
        root.getChildren().add(title);
        root.getChildren().add(new Separator());

        for (OperatorCardDaySection day : days) {
            root.getChildren().add(buildDayBox(day, ff));
        }

        return root;
    }

    private static VBox buildDayBox(OperatorCardDaySection day, String ff) {
        Label dayTitle = new Label(formatDaySectionTitle(day.date()));
        dayTitle.getStyleClass().add("pm-operator-card-day-title");
        dayTitle.setMaxWidth(Double.MAX_VALUE);

        GridPane grid = buildDayGrid(day);
        grid.setStyle("-fx-font-family: " + ff + ";");
        grid.getStyleClass().add("pm-operator-card-grid");

        return new VBox(6, dayTitle, grid);
    }

    /**
     * {@code scene} へアプリ本体の {@code pm-ai-desktop.css} を適用する。印刷は画面プレビューと違い
     * アプリ本体の {@link Scene}（スタイルシート適用済み）を経由しないため、明示的に当てないと
     * 枠線・見出し背景色・行の詰め幅がプレビューと異なって見える。
     */
    public static void attachDesktopStylesheet(Scene scene) {
        URL url =
                OperatorCardPreviewFactory.class.getResource(
                        "/jp/co/pm/ai/desktop/css/pm-ai-desktop.css");
        if (url == null) {
            return;
        }
        String external = url.toExternalForm();
        if (!scene.getStylesheets().contains(external)) {
            scene.getStylesheets().add(external);
        }
    }

    static GridPane buildDayGrid(OperatorCardDaySection day) {
        GridPane grid = new GridPane();
        grid.setHgap(0);
        grid.setVgap(0);
        grid.setPadding(new Insets(4, 0, 12, 0));

        ColumnConstraints c0 = new ColumnConstraints(88, 88, 120);
        ColumnConstraints c1 = new ColumnConstraints(80, 100, 180);
        ColumnConstraints c2 = new ColumnConstraints(80, 120, 220);
        ColumnConstraints c3 = new ColumnConstraints(56, 72, 100);
        ColumnConstraints c4 = new ColumnConstraints(56, 72, 88);
        ColumnConstraints c5 = new ColumnConstraints(56, 72, 88);
        ColumnConstraints c6 = new ColumnConstraints(80, 120, 280);
        c1.setHgrow(Priority.ALWAYS);
        c2.setHgrow(Priority.ALWAYS);
        c6.setHgrow(Priority.ALWAYS);
        grid.getColumnConstraints().addAll(c0, c1, c2, c3, c4, c5, c6);

        String[] hdr =
                new String[] {
                    "時間帯",
                    "工程",
                    "機械",
                    "依頼NO",
                    "当日配台",
                    "換算",
                    "メンバー"
                };
        for (int c = 0; c < hdr.length; c++) {
            Label h = new Label(hdr[c]);
            h.getStyleClass().add("pm-operator-card-th");
            h.setMaxWidth(Double.MAX_VALUE);
            GridPane.setHgrow(h, Priority.ALWAYS);
            grid.add(h, c, 0);
        }

        int row = 1;
        for (OperatorCardTaskRow tr : day.rows()) {
            grid.getRowConstraints().add(new RowConstraints(22));
            addCell(grid, tr.timeRange(), row, 0, "pm-operator-card-td-time");
            addCell(grid, tr.processName(), row, 1, "pm-operator-card-td");
            addCell(grid, tr.machineName(), row, 2, "pm-operator-card-td");
            addCell(grid, tr.requestNo(), row, 3, "pm-operator-card-td");
            addCell(grid, tr.qtyDispatchDay(), row, 4, "pm-operator-card-td-num");
            addCell(grid, tr.qtyConverted(), row, 5, "pm-operator-card-td-num");
            addCell(grid, tr.memberNames(), row, 6, "pm-operator-card-td");
            row++;
        }

        if (day.rows().isEmpty()) {
            Label empty = new Label("この日の予定はありません");
            empty.getStyleClass().add("pm-operator-card-empty");
            grid.add(empty, 0, 1, 7, 1);
            GridPane.setHalignment(empty, HPos.CENTER);
        }

        return grid;
    }

    private static void addCell(GridPane grid, String text, int row, int col, String styleClass) {
        Label l = new Label(text != null ? text : "");
        l.setWrapText(true);
        l.setMaxWidth(Double.MAX_VALUE);
        l.getStyleClass().add(styleClass);
        GridPane.setHgrow(l, Priority.ALWAYS);
        grid.add(l, col, row);
    }

    static String cssFontFamily(String fontFamily) {
        String f = fontFamily != null ? fontFamily.trim() : "SansSerif";
        if (f.contains("'")) {
            return "\"" + f.replace("\\", "\\\\").replace("\"", "\\\"") + "\"";
        }
        return "'" + f + "'";
    }

    /** {@code uuuu-MM-dd  MM/dd（月火...）} */
    static String formatDaySectionTitle(LocalDate date) {
        String iso = date.toString();
        String md = date.format(DateTimeFormatter.ofPattern("MM/dd"));
        String wd = japaneseWeekdayShort(date.getDayOfWeek());
        return iso + "  " + md + "（" + wd + "）";
    }

    static String japaneseWeekdayShort(DayOfWeek dow) {
        return JP_WEEKDAY_SHORT[dow.getValue() - 1];
    }

    static String formatIssuedAt() {
        LocalDateTime now = LocalDateTime.now();
        DateTimeFormatter f =
                DateTimeFormatter.ofPattern(
                        "発行日時： uuuu年M月d日 HH:mm",
                        Locale.JAPAN);
        return now.format(f);
    }
}
