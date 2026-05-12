package jp.co.pm.ai.desktop.print;

import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Locale;

import javafx.geometry.HPos;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Parent;
import javafx.scene.control.Label;
import javafx.scene.control.Separator;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.RowConstraints;
import javafx.scene.layout.VBox;

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
        String ff = cssFontFamily(fontFamily);

        VBox root = new VBox(10);
        root.setPadding(new Insets(16, 20, 16, 20));
        root.setPrefWidth(A4_PREF_WIDTH);
        root.setMinWidth(A4_PREF_WIDTH);
        root.setMaxWidth(A4_PREF_WIDTH);
        root.setPrefHeight(A4_PREF_HEIGHT);
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

        Label title = new Label(page.operatorName());
        title.getStyleClass().add("pm-operator-card-title");
        title.setMaxWidth(Double.MAX_VALUE);
        title.setAlignment(Pos.CENTER_LEFT);

        root.getChildren().add(headingRow);
        root.getChildren().add(new Separator());
        root.getChildren().add(title);
        root.getChildren().add(new Separator());

        for (OperatorCardDaySection day : page.days()) {
            Label dayTitle = new Label(formatDaySectionTitle(day.date()));
            dayTitle.getStyleClass().add("pm-operator-card-day-title");
            dayTitle.setMaxWidth(Double.MAX_VALUE);

            GridPane grid = buildDayGrid(day);
            grid.setStyle("-fx-font-family: " + ff + ";");
            grid.getStyleClass().add("pm-operator-card-grid");

            VBox dayBox = new VBox(6, dayTitle, grid);
            root.getChildren().add(dayBox);
        }

        return root;
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
