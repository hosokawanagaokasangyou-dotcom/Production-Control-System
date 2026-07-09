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
     * ルート幅を合わせて組み立てる。画面プレビューは {@link ScrollPane}（fitToWidth）で縮小表示するため
     * 見た目上は全列表示されるが、{@code PrinterJob#printPage} は Node をそのままの寸法で用紙に描画する
     * ため、固定幅（{@link #A4_PREF_WIDTH}）のままだと A4 の可印刷幅を超えて右側の列がクリップされる。
     */
    public static Parent buildRoot(OperatorCardPage page, String fontFamily, double rootWidth) {
        return assembleRoot(page.operatorName(), page.days(), fontFamily, rootWidth);
    }

    /**
     * 印刷用に、1 オペレーター分のカードを用紙の可印刷高さ（{@code PageLayout#getPrintableHeight()}）
     * へ収まる単位で複数ページ（{@link Parent} のリスト）に分割する。
     *
     * <p>{@link #buildRoot} は全日分を 1 枚のルートへ詰め込むため、選択日数が多い・当日配台の行数が
     * 多い場合に、可印刷高さを超えた分（末尾の日）が用紙からクリップされて印刷結果に現れない不具合が
     * あった。本メソッドは実際の CSS 適用後の必要高さを 1 日ずつ計測しながら、収まらなくなった時点で
     * 新しいページ（見出し・オペレーター名を再掲）へ切り出す。
     *
     * @param printableHeight 用紙 1 枚の可印刷高さ（pt）。非正・非有限値のときは分割せず 1 ページのみ返す。
     */
    public static List<Parent> buildPrintPages(
            OperatorCardPage page, String fontFamily, double rootWidth, double printableHeight) {
        List<OperatorCardDaySection> days = page.days();
        if (days.isEmpty() || !Double.isFinite(printableHeight) || printableHeight <= 0) {
            return List.of(assembleRoot(page.operatorName(), days, fontFamily, rootWidth));
        }

        String ff = cssFontFamily(fontFamily);
        double width = rootWidth > 0 ? rootWidth : A4_PREF_WIDTH;

        List<Parent> pages = new ArrayList<>();
        VBox current = assembleRoot(page.operatorName(), List.of(), fontFamily, rootWidth);
        for (OperatorCardDaySection day : days) {
            VBox dayBox = buildDayBox(day, ff);
            current.getChildren().add(dayBox);
            double height = measuredPrefHeight(current, width);
            boolean overflowsWithMoreThanOneDay = height > printableHeight + 0.5 && dayCountOf(current) > 1;
            if (overflowsWithMoreThanOneDay) {
                current.getChildren().remove(dayBox);
                pages.add(current);
                current = assembleRoot(page.operatorName(), List.of(), fontFamily, rootWidth);
                current.getChildren().add(dayBox);
            }
        }
        pages.add(current);
        return pages;
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
     * {@code root} を（未接続なら）一時 {@link Scene} へ接続して {@code pm-ai-desktop.css} を適用し、
     * CSS 反映後の必要高さ（ボーダー・パディング・フォントサイズを含む）を返す。
     *
     * <p>印刷は画面プレビューと違いアプリ本体の {@link Scene}（スタイルシート適用済み）を経由しないため、
     * 明示的にスタイルシートを当てないと枠線・見出し背景色・行の詰め幅がプレビューと異なって見える。
     */
    private static double measuredPrefHeight(VBox root, double width) {
        Scene scene = root.getScene();
        boolean tempScene = scene == null;
        if (tempScene) {
            scene = new Scene(root, Math.max(1, width), 1, Color.WHITE);
        }
        attachDesktopStylesheet(scene);
        root.applyCss();
        double height = root.prefHeight(width);
        if (tempScene) {
            // 計測専用の Scene から root を解放し、呼び出し側が後で（可印刷高さぴったりの）
            // 本番用 Scene を新規に割り当てられるようにする。
            scene.setRoot(new javafx.scene.Group());
        }
        return height;
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
