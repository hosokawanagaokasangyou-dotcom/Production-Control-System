package jp.co.pm.ai.desktop.ui;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.canvas.Canvas;
import javafx.scene.canvas.GraphicsContext;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.Tooltip;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Region;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;

/**
 * 「結果_設備ガント」Excel と同一データ（JSON の columns / rows）から、横軸が時刻スロットの
 * タイムラインをグラフィカルに描画するビュー。計画結果ビューアの表／セル着色ガントより視認性を優先する。
 */
public final class EquipmentGraphicGanttPane extends BorderPane {

    private static final Pattern TIME_SLOT_HEADER =
            Pattern.compile("^\\s*(\\d{1,2}):(\\d{2})\\s*$");

    private static final double LABEL_MIN_WIDTH = 220;
    private static final double LABEL_MAX_WIDTH = 320;
    private static final double ROW_HEIGHT = 26;
    private static final double SECTION_ROW_HEIGHT = 30;
    private static final double HEADER_HEIGHT = 40;
    /** 1 スロットあたりの幅（px）。Excel の 10 分スロットを想定 */
    private static final double SLOT_WIDTH = 9;

    private static final Color EMPTY_LIGHT = Color.web("#ffffff");
    private static final Color EMPTY_BAND = Color.web("#f2f2f2");
    private static final Color BORDER_GRID = Color.web("#d9d9d9");
    private static final Color BAR_DEFAULT = Color.web("#5b9bd5");
    private static final Color BAR_BREAK = Color.web("#90CAF9");
    private static final Color BAR_STARTUP = Color.web("#fed7aa");
    private static final Color HEADER_AXIS = Color.web("#d9d9d9");

    private EquipmentGraphicGanttPane() {}

    /**
     * @param columns シート列見出し
     * @param rows データ行（フィルタ行を含まない素データ）
     * @return 時刻列が検出できない場合は説明ラベルのみのペイン
     */
    public static BorderPane build(
            List<String> columns, ObservableList<ObservableList<String>> rows) {
        BorderPane root = new BorderPane();
        List<String> effCols = columns;
        ObservableList<ObservableList<String>> effRows = rows;
        RepairResult repaired = tryRepairPandasUnnamedEquipmentTimeline(effCols, effRows);
        if (repaired != null) {
            effCols = repaired.columns();
            effRows = repaired.rows();
        }
        ParseResult parsed = parse(effCols, effRows);
        // #region agent log
        {
            int nonEmptySlots = 0;
            int dataDisplayRows = 0;
            int sectionRows = 0;
            for (DisplayRow dr : parsed.displayRows) {
                if (dr.sectionBanner != null) {
                    sectionRows++;
                    continue;
                }
                dataDisplayRows++;
                if (dr.cellsInSlots != null) {
                    for (String c : dr.cellsInSlots) {
                        if (c != null && !c.strip().isEmpty()) {
                            nonEmptySlots++;
                        }
                    }
                }
            }
            String col0preview = "";
            if (effRows != null && !effRows.isEmpty()) {
                ObservableList<String> r0 = effRows.get(0);
                if (r0 != null && !r0.isEmpty()) {
                    String s = r0.get(0);
                    if (s != null) {
                        col0preview = s.length() > 72 ? s.substring(0, 72) : s;
                    }
                }
            }
            StringBuilder hb = new StringBuilder();
            int hlim = Math.min(10, effCols.size());
            for (int hi = 0; hi < hlim; hi++) {
                if (hi > 0) {
                    hb.append('|');
                }
                String hx = effCols.get(hi);
                hb.append(hx != null ? hx : "");
            }
            String headersPreview = hb.length() > 240 ? hb.substring(0, 240) : hb.toString();
            agentDebugLog(
                    "T1",
                    "EquipmentGraphicGanttPane.build:afterParse",
                    "equipment graphic parse summary",
                    String.format(
                            "{\"hypothesisMap\":{\"H1_jsonCellsEmpty\":%s,\"H2_allSectionRows\":%s},"
                                    + "\"nonEmptySlotCells\":%d,\"timelineCanvasRows\":%d,\"sectionBannerRows\":%d,"
                                    + "\"inputRows\":%d,\"slotColumnCount\":%d,\"slotMinutes\":%d,"
                                    + "\"repairedUnnamed\":%b,\"firstCol0Preview\":\"%s\","
                                    + "\"headersPreview\":\"%s\"}",
                            nonEmptySlots == 0,
                            sectionRows > 0 && dataDisplayRows == 0,
                            nonEmptySlots,
                            dataDisplayRows,
                            sectionRows,
                            effRows != null ? effRows.size() : 0,
                            parsed.slotColumnIndices.size(),
                            parsed.slotMinutes,
                            repaired != null,
                            jsonEscape(col0preview),
                            jsonEscape(headersPreview)));
        }
        // #endregion
        if (parsed.slotColumnIndices.isEmpty()) {
            Label msg =
                    new Label(
                            "このシートから時刻列（列見出しが HH:MM 形式）を検出できませんでした。\n"
                                    + "「結果_設備ガント」形式の JSON を開いているか確認してください。");
            msg.setWrapText(true);
            msg.setPadding(new Insets(16));
            root.setCenter(msg);
            return root;
        }

        VBox body = new VBox(0);
        double timelineWidth = parsed.slotColumnIndices.size() * SLOT_WIDTH;
        double labelWidth = Math.min(LABEL_MAX_WIDTH, Math.max(LABEL_MIN_WIDTH, 200));

        Canvas headerCanvas = new Canvas(timelineWidth, HEADER_HEIGHT);
        drawTimeAxis(headerCanvas.getGraphicsContext2D(), parsed, timelineWidth);

        HBox headerRow = new HBox(0);
        Region spacer = new Region();
        spacer.setMinWidth(labelWidth);
        spacer.setPrefWidth(labelWidth);
        headerRow.getChildren().addAll(spacer, headerCanvas);
        body.getChildren().add(headerRow);

        int progressTotal =
                parsed.progressColumnIndices.size() * 56 + Math.max(0, parsed.progressColumnIndices.size() - 1) * 4;
        int dataStripe = 0;
        for (int ri = 0; ri < parsed.displayRows.size(); ri++) {
            DisplayRow dr = parsed.displayRows.get(ri);
            if (dr.sectionBanner != null) {
                Label ban = new Label(dr.sectionBanner);
                ban.setPrefHeight(SECTION_ROW_HEIGHT);
                ban.setMinHeight(SECTION_ROW_HEIGHT);
                ban.setMaxWidth(Double.MAX_VALUE);
                ban.setAlignment(Pos.CENTER_LEFT);
                ban.setPadding(new Insets(2, 8, 2, 8));
                ban.setTextFill(Color.WHITE);
                ban.setStyle(
                        "-fx-background-color: #1f4e79; -fx-font-weight: bold; -fx-font-size: 12px;");
                ban.setMinWidth(labelWidth + timelineWidth + progressTotal);
                body.getChildren().add(ban);
                continue;
            }

            Label rowLabel = new Label(dr.leftLabel);
            rowLabel.setMinWidth(labelWidth);
            rowLabel.setPrefWidth(labelWidth);
            rowLabel.setMaxWidth(labelWidth);
            rowLabel.setWrapText(true);
            rowLabel.setAlignment(Pos.CENTER_LEFT);
            rowLabel.setPadding(new Insets(2, 6, 2, 6));
            rowLabel.setFont(Font.font(11));
            rowLabel.setStyle(
                    "-fx-background-color: #f1f5f9; -fx-border-color: #e2e8f0; -fx-border-width: 0 1 0 0;");

            Canvas rowCanvas = new Canvas(timelineWidth, ROW_HEIGHT);
            drawTimelineRow(
                    rowCanvas.getGraphicsContext2D(), dr.cellsInSlots, dataStripe++);

            String tip =
                    dr.leftLabel.replace("\n", " ")
                            + "\n（スロット "
                            + parsed.slotMinutes
                            + " 分・設備ガント JSON と同一データ）";
            Tooltip.install(rowCanvas, new Tooltip(tip));

            HBox progBox = new HBox(4);
            progBox.setAlignment(Pos.CENTER_LEFT);
            for (int pc : parsed.progressColumnIndices) {
                String pv =
                        dr.rawRow.size() > pc && dr.rawRow.get(pc) != null
                                ? dr.rawRow.get(pc).strip()
                                : "";
                Label pl = new Label(pv);
                pl.setMinWidth(52);
                pl.setPrefWidth(52);
                pl.setMaxWidth(52);
                pl.setWrapText(true);
                pl.setAlignment(Pos.CENTER);
                pl.setFont(Font.font(9));
                pl.setStyle(
                        pv.isEmpty()
                                ? "-fx-background-color: #fffbf0; -fx-border-color: #f0e1b7;"
                                : "-fx-background-color: #fff2cc; -fx-border-color: #d6b656;");
                progBox.getChildren().add(pl);
            }

            HBox line = new HBox(0);
            line.getChildren().addAll(rowLabel, rowCanvas);
            if (!progBox.getChildren().isEmpty()) {
                line.getChildren().add(progBox);
            }
            body.getChildren().add(line);
        }

        ScrollPane scroll = new ScrollPane(body);
        scroll.setFitToHeight(true);
        scroll.setPannable(true);
        scroll.setPadding(new Insets(4));
        root.setCenter(scroll);

        Label hint =
                new Label(
                        "ヒント: 横スクロールで時刻軸を追えます。ブロックは Excel と同じセル内容を連続結合した帯です。"
                                + "（10 分スロット幅は生成側の設定と一致している必要があります）");
        hint.setWrapText(true);
        hint.setStyle("-fx-text-fill: #64748b; -fx-font-size: 11px;");
        hint.setPadding(new Insets(0, 8, 8, 8));
        root.setBottom(hint);
        return root;
    }

    private static void drawTimeAxis(
            GraphicsContext gc, ParseResult parsed, double timelineWidth) {
        gc.setFill(HEADER_AXIS);
        gc.fillRect(0, 0, timelineWidth, HEADER_HEIGHT);
        gc.setStroke(BORDER_GRID);
        gc.setLineWidth(0.5);
        gc.strokeRect(0, 0, timelineWidth, HEADER_HEIGHT);

        List<Integer> slotCols = parsed.slotColumnIndices;
        int n = slotCols.size();
        int step = Math.max(1, 60 / Math.max(1, parsed.slotMinutes));

        gc.setFill(Color.BLACK);
        gc.setFont(Font.font(10));
        LocalTime t0 = parsed.slotBaseTime;
        for (int i = 0; i < n; i += step) {
            double x = i * SLOT_WIDTH;
            LocalTime tt = t0.plusMinutes((long) i * parsed.slotMinutes);
            String txt = tt.format(DateTimeFormatter.ofPattern("H:mm"));
            gc.fillText(txt, x + 2, HEADER_HEIGHT - 8);
            gc.strokeLine(x, 0, x, HEADER_HEIGHT);
        }
    }

    private static void drawTimelineRow(
            GraphicsContext gc, List<String> slotTexts, int stripeIndex) {
        int n = slotTexts.size();
        boolean stripe = (stripeIndex & 1) == 0;
        for (int i = 0; i < n; i++) {
            double x = i * SLOT_WIDTH;
            gc.setFill(stripe ? EMPTY_LIGHT : EMPTY_BAND);
            gc.fillRect(x, 0, SLOT_WIDTH, ROW_HEIGHT);
        }
        gc.setStroke(BORDER_GRID);
        gc.setLineWidth(0.3);
        for (int i = 0; i <= n; i++) {
            gc.strokeLine(i * SLOT_WIDTH, 0, i * SLOT_WIDTH, ROW_HEIGHT);
        }

        int runStart = -1;
        String runText = "";
        for (int i = 0; i < n; i++) {
            String t = slotTexts.get(i) != null ? slotTexts.get(i).strip() : "";
            boolean empty = t.isEmpty();
            if (empty) {
                if (runStart >= 0) {
                    paintBar(gc, runStart, i - 1, runText);
                    runStart = -1;
                    runText = "";
                }
                continue;
            }
            if (runStart < 0) {
                runStart = i;
                runText = t;
            } else if (!t.equals(runText)) {
                paintBar(gc, runStart, i - 1, runText);
                runStart = i;
                runText = t;
            }
        }
        if (runStart >= 0) {
            paintBar(gc, runStart, n - 1, runText);
        }
    }

    private static void paintBar(GraphicsContext gc, int fromSlot, int toSlot, String text) {
        double x = fromSlot * SLOT_WIDTH;
        double w = (toSlot - fromSlot + 1) * SLOT_WIDTH;
        Color fill = barColorFor(text);
        gc.setFill(fill);
        double arc = 3;
        gc.fillRoundRect(x + 0.5, 3, w - 1, ROW_HEIGHT - 6, arc, arc);
        gc.setStroke(Color.web("#2e5597"));
        gc.setLineWidth(0.5);
        gc.strokeRoundRect(x + 0.5, 3, w - 1, ROW_HEIGHT - 6, arc, arc);

        gc.setFill(Color.WHITE);
        gc.setFont(Font.font(9));
        String shortTxt = text.replace('\n', ' ');
        if (shortTxt.length() > 80) {
            shortTxt = shortTxt.substring(0, 77) + "...";
        }
        double maxChars = Math.max(4, (w - 6) / 5);
        if (shortTxt.length() > maxChars) {
            shortTxt = shortTxt.substring(0, (int) maxChars - 2) + "..";
        }
        if (w > 28 && !shortTxt.isEmpty()) {
            gc.fillText(shortTxt, x + 4, ROW_HEIGHT / 2 + 3);
        }
    }

    private static Color barColorFor(String t) {
        if (t.contains("休憩") || t.contains("（休憩）")) {
            return BAR_BREAK;
        }
        if (t.contains("日次始業準備")) {
            return BAR_STARTUP;
        }
        return BAR_DEFAULT;
    }

    private static LocalTime parseTimeHeader(String col) {
        if (col == null) {
            return null;
        }
        var m = TIME_SLOT_HEADER.matcher(col.strip());
        if (!m.matches()) {
            return null;
        }
        int hh = Integer.parseInt(m.group(1));
        int mm = Integer.parseInt(m.group(2));
        try {
            return LocalTime.of(hh, mm);
        } catch (Exception e) {
            return null;
        }
    }

    /**
     * production_plan JSON が pandas 由来で列名が Unnamed:0 のみのとき、Excel 上の
     * 「日付 / 機械名 / … / HH:MM …」行を列見出し行として採用する（read_excel header=0 ミスアラインの救済）。
     */
    private static RepairResult tryRepairPandasUnnamedEquipmentTimeline(
            List<String> columns, ObservableList<ObservableList<String>> rows) {
        if (columns == null
                || columns.isEmpty()
                || rows == null
                || rows.isEmpty()
                || !looksLikePandasUnnamedHeaderColumns(columns)) {
            return null;
        }
        int width = columns.size();
        for (int ri = 0; ri < Math.min(50, rows.size()); ri++) {
            ObservableList<String> row = rows.get(ri);
            if (row == null || row.isEmpty()) {
                continue;
            }
            if (!"日付".equals(strAt(row, 0))) {
                continue;
            }
            int firstTimeCol = -1;
            for (int c = 1; c < row.size(); c++) {
                if (parseTimeHeader(strAt(row, c)) != null) {
                    firstTimeCol = c;
                    break;
                }
            }
            if (firstTimeCol < 1) {
                continue;
            }
            int timeHits = 0;
            for (int c = firstTimeCol; c < row.size(); c++) {
                if (parseTimeHeader(strAt(row, c)) != null) {
                    timeHits++;
                }
            }
            if (timeHits < 2) {
                continue;
            }
            int w = Math.max(width, row.size());
            List<String> newCols = new ArrayList<>(w);
            for (int c = 0; c < w; c++) {
                String v = strAt(row, c);
                if (v.isEmpty()) {
                    newCols.add(c < columns.size() ? columns.get(c) : ("列" + c));
                } else {
                    newCols.add(v);
                }
            }
            ObservableList<ObservableList<String>> newRows = FXCollections.observableArrayList();
            for (int rj = ri + 1; rj < rows.size(); rj++) {
                newRows.add(rows.get(rj));
            }
            return new RepairResult(newCols, newRows);
        }
        return null;
    }

    private static boolean looksLikePandasUnnamedHeaderColumns(List<String> columns) {
        int n = Math.min(3, columns.size());
        for (int i = 0; i < n; i++) {
            String h = columns.get(i);
            if (h == null || !h.startsWith("Unnamed")) {
                return false;
            }
        }
        return true;
    }

    private static String strAt(ObservableList<String> row, int c) {
        if (row == null || c < 0 || c >= row.size()) {
            return "";
        }
        String v = row.get(c);
        return v != null ? v.strip() : "";
    }

    // #region agent log
    /**
     * デバッグ NDJSON（セッション f0dedd）。Windows / WSL のいずれかに書ければよい。
     */
    private static void agentDebugLog(
            String hypothesisId, String location, String message, String dataJsonObject) {
        long ts = System.currentTimeMillis();
        String line =
                "{\"sessionId\":\"f0dedd\",\"timestamp\":"
                        + ts
                        + ",\"hypothesisId\":\""
                        + hypothesisId
                        + "\",\"location\":\""
                        + jsonEscape(location)
                        + "\",\"message\":\""
                        + jsonEscape(message)
                        + "\",\"data\":"
                        + dataJsonObject
                        + "}\n";
        String env = System.getenv("PM_AI_DEBUG_LOG");
        String[] candidates =
                new String[] {
                    env != null && !env.isBlank() ? env : "",
                    "/mnt/c/工程管理AIプロジェクト_JAVA/.cursor/debug-f0dedd.log",
                    "C:\\工程管理AIプロジェクト_JAVA\\.cursor\\debug-f0dedd.log",
                };
        for (String p : candidates) {
            if (p == null || p.isBlank()) {
                continue;
            }
            try {
                Path path = Path.of(p);
                Path parent = path.getParent();
                if (parent != null) {
                    Files.createDirectories(parent);
                }
                Files.writeString(
                        path,
                        line,
                        StandardCharsets.UTF_8,
                        StandardOpenOption.CREATE,
                        StandardOpenOption.APPEND);
                break;
            } catch (Exception ignored) {
                // try next path
            }
        }
    }

    private static String jsonEscape(String s) {
        if (s == null) {
            return "";
        }
        return s.replace("\\", "\\\\").replace("\"", "\\\"");
    }

    /** デバッグ: グラフィックタブで選択シートを読み込んだとき */
    public static void agentLogSheetLoad(String sheetName, int columnCount) {
        agentDebugLog(
                "H2",
                "EquipmentGanttGraphicTabController.applySelectedSheetFromMap",
                "selected sheet for graphic gantt",
                String.format(
                        "{\"sheetName\":\"%s\",\"columnCount\":%d}",
                        jsonEscape(sheetName != null ? sheetName : ""),
                        columnCount));
    }
    // #endregion

    private record RepairResult(
            List<String> columns, ObservableList<ObservableList<String>> rows) {}

    private static ParseResult parse(List<String> columns, ObservableList<ObservableList<String>> rows) {
        ParseResult pr = new ParseResult();
        List<Integer> slots = new ArrayList<>();
        for (int c = 0; c < columns.size(); c++) {
            String h = columns.get(c);
            if (parseTimeHeader(h) != null) {
                slots.add(c);
            }
        }
        pr.slotColumnIndices = slots;
        if (!slots.isEmpty()) {
            LocalTime t0 = parseTimeHeader(columns.get(slots.get(0)));
            if (t0 != null) {
                pr.slotBaseTime = t0;
            }
        }
        if (slots.size() >= 2) {
            LocalTime a = parseTimeHeader(columns.get(slots.get(0)));
            LocalTime b = parseTimeHeader(columns.get(slots.get(1)));
            if (a != null && b != null) {
                int delta =
                        (b.getHour() * 60 + b.getMinute()) - (a.getHour() * 60 + a.getMinute());
                if (delta > 0) {
                    pr.slotMinutes = delta;
                }
            }
        }
        for (int c = 0; c < columns.size(); c++) {
            String h = columns.get(c);
            if (h != null && h.endsWith("進度")) {
                pr.progressColumnIndices.add(c);
            }
        }

        String carriedDate = "";
        for (int r = 0; r < rows.size(); r++) {
            ObservableList<String> row = rows.get(r);
            if (row == null || row.isEmpty()) {
                continue;
            }
            String c0 = row.size() > 0 && row.get(0) != null ? row.get(0).strip() : "";
            if (isSectionRow(row)) {
                String banner = !c0.isEmpty() ? c0 : sectionTitleFromRow(row);
                pr.displayRows.add(new DisplayRow(banner, null, null, row));
                continue;
            }
            int dateCol = columnIndex(columns, "日付");
            if (dateCol >= 0 && row.size() > dateCol) {
                String dv = row.get(dateCol) != null ? row.get(dateCol).strip() : "";
                if (!dv.isEmpty()) {
                    carriedDate = dv;
                }
            }
            String left = buildLeftLabel(columns, row, carriedDate);
            List<String> slotCells = new ArrayList<>();
            for (int si : slots) {
                String v =
                        row.size() > si && row.get(si) != null ? row.get(si) : "";
                slotCells.add(v);
            }
            pr.displayRows.add(new DisplayRow(null, left, slotCells, row));
        }
        return pr;
    }

    private static boolean isSectionRow(ObservableList<String> row) {
        for (int i = 0; i < Math.min(4, row.size()); i++) {
            String s = row.get(i) != null ? row.get(i) : "";
            if (s.contains("■") || s.contains("▪") || s.contains("【")) {
                return true;
            }
        }
        return false;
    }

    private static String sectionTitleFromRow(ObservableList<String> row) {
        for (int i = 0; i < row.size(); i++) {
            String s = row.get(i) != null ? row.get(i).strip() : "";
            if (!s.isEmpty()) {
                return s;
            }
        }
        return "（区切り）";
    }

    private static int columnIndex(List<String> columns, String name) {
        for (int i = 0; i < columns.size(); i++) {
            if (name.equals(columns.get(i))) {
                return i;
            }
        }
        return -1;
    }

    private static String buildLeftLabel(
            List<String> columns, ObservableList<String> row, String carriedDate) {
        String mach = cellAt(columns, row, "機械名");
        String proc = cellAt(columns, row, "工程名");
        String task = cellAt(columns, row, "タスク概覝");
        String tb = cellAt(columns, row, "日時帯");
        StringBuilder sb = new StringBuilder();
        if (!carriedDate.isEmpty()) {
            sb.append(carriedDate);
        }
        if (!mach.isEmpty()) {
            if (sb.length() > 0) {
                sb.append("\n");
            }
            sb.append(mach);
        }
        if (!proc.isEmpty() && !proc.equals("—")) {
            sb.append("\n").append(proc);
        }
        if (!task.isEmpty() && !task.equals("—")) {
            String t = task.length() > 120 ? task.substring(0, 117) + "..." : task;
            sb.append("\n").append(t);
        }
        if (!tb.isEmpty() && sb.length() == 0) {
            sb.append(tb);
        }
        if (sb.length() == 0) {
            return "（行）";
        }
        return sb.toString();
    }

    private static String cellAt(List<String> columns, ObservableList<String> row, String colName) {
        int i = columnIndex(columns, colName);
        if (i < 0 || row.size() <= i) {
            return "";
        }
        String v = row.get(i);
        return v != null ? v.strip() : "";
    }

    private static final class ParseResult {
        List<Integer> slotColumnIndices = new ArrayList<>();
        List<Integer> progressColumnIndices = new ArrayList<>();
        int slotMinutes = 10;
        LocalTime slotBaseTime = LocalTime.of(8, 0);
        List<DisplayRow> displayRows = new ArrayList<>();
    }

    private static final class DisplayRow {
        /** 非 null のときセクション行 */
        final String sectionBanner;
        /** データ行の左ラベル */
        final String leftLabel;
        final List<String> cellsInSlots;
        final ObservableList<String> rawRow;

        DisplayRow(
                String sectionBanner,
                String leftLabel,
                List<String> cellsInSlots,
                ObservableList<String> rawRow) {
            this.sectionBanner = sectionBanner;
            this.leftLabel = leftLabel;
            this.cellsInSlots = cellsInSlots;
            this.rawRow = rawRow;
        }
    }
}
