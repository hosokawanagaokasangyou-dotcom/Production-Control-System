package jp.co.pm.ai.desktop.ui;

import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.function.BiConsumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Group;
import javafx.scene.Scene;
import javafx.scene.canvas.Canvas;
import javafx.scene.canvas.GraphicsContext;
import javafx.scene.control.Label;
import javafx.scene.control.ScrollPane;
import javafx.scene.control.SplitPane;
import javafx.scene.control.Tooltip;
import javafx.application.Platform;
import javafx.scene.Node;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;

import javafx.scene.layout.Region;

import jp.co.pm.ai.desktop.config.DesktopTheme;

/**
 * 「結果_設備ガント」Excel と同一データ（JSON の columns / rows）から、横軸が時刻スロットの
 * タイムラインをグラフィカルに描画するビュー。計画結果ビューアの表／セル着色ガントより視認性を優先する。
 */
public final class EquipmentGraphicGanttPane extends BorderPane {

    private static final Pattern TIME_SLOT_HEADER =
            Pattern.compile("^\\s*(\\d{1,2}):(\\d{2})\\s*$");

    /**
     * 「結果_設備ガント」日付列の「【2026/05/07】」「【2026-05-07】」等。先頭列の【による誤判定を避ける。
     */
    private static final Pattern BRACKETED_PLAIN_DATE_LABEL =
            Pattern.compile("^\\s*【\\s*\\d{4}[/\\-]\\d{1,2}[/\\-]\\d{1,2}\\s*】\\s*$");

    private static final Pattern LOOSE_YMD =
            Pattern.compile("(\\d{4})[/\\-.](\\d{1,2})[/\\-.](\\d{1,2})");

    private static final double BASE_LABEL_MIN_WIDTH = 220;
    private static final double BASE_LABEL_MAX_WIDTH = 360;
    private static final double BASE_ROW_HEIGHT = 26;
    private static final double BASE_SECTION_ROW_HEIGHT = 30;
    private static final double BASE_HEADER_HEIGHT = 40;
    /** 1 スロットあたりの幅（px、倍率1のとき）。Excel の 10 分スロットを想定 */
    private static final double BASE_SLOT_WIDTH = 9;

    public static final double DEFAULT_MACHINE_COLUMN_WIDTH = 140;
    public static final double DEFAULT_PROCESS_COLUMN_WIDTH = 220;
    private static final double MIN_SIDE_COL_WIDTH = 48;
    private static final double MAX_SIDE_COL_WIDTH = 800;

    private static final Group MEASURE_ROOT = new Group();
    private static final Scene MEASURE_SCENE = new Scene(MEASURE_ROOT, 4000, 4000);

    private EquipmentGraphicGanttPane() {}

    public static double clampMachineColumnWidth(double w) {
        return clampSideCol(w, DEFAULT_MACHINE_COLUMN_WIDTH);
    }

    public static double clampProcessColumnWidth(double w) {
        return clampSideCol(w, DEFAULT_PROCESS_COLUMN_WIDTH);
    }

    private static double clampSideCol(double w, double defaultW) {
        if (!Double.isFinite(w) || w < MIN_SIDE_COL_WIDTH) {
            return defaultW;
        }
        return Math.min(MAX_SIDE_COL_WIDTH, Math.max(MIN_SIDE_COL_WIDTH, w));
    }

    /**
     * @param columns シート列見出し
     * @param rows データ行（フィルタ行を含まない素データ）
     * @return 時刻列が検出できない場合は説明ラベルのみのペイン
     */
    public static BorderPane build(
            List<String> columns, ObservableList<ObservableList<String>> rows) {
        return build(columns, rows, DesktopTheme.LIGHT, 1.0);
    }

    public static BorderPane build(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            DesktopTheme theme,
            double zoom) {
        return build(
                columns,
                rows,
                theme,
                zoom,
                DEFAULT_MACHINE_COLUMN_WIDTH,
                DEFAULT_PROCESS_COLUMN_WIDTH,
                null);
    }

    /**
     * @param theme アプリの {@link DesktopTheme}（Canvas 帯の配色に反映）
     * @param zoom 表示倍率（0.5〜2.0 程度を想定。スロット幅・行高・フォントに連動）
     * @param machineColWidth 機械名列の幅（px）
     * @param processColWidth 工程名列の幅（px）
     * @param onLeftColumnWidthsChanged ヘッダの境界ドラッグ後に {@code (機械列幅, 工程列幅)} を通知（永続化用）
     */
    public static BorderPane build(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            DesktopTheme theme,
            double zoom,
            double machineColWidth,
            double processColWidth,
            BiConsumer<Double, Double> onLeftColumnWidthsChanged) {
        BorderPane root = new BorderPane();
        List<String> effCols = columns;
        ObservableList<ObservableList<String>> effRows = rows;
        RepairResult repaired = tryRepairPandasUnnamedEquipmentTimeline(effCols, effRows);
        if (repaired != null) {
            effCols = repaired.columns();
            effRows = repaired.rows();
        }
        ParseResult parsed = parse(effCols, effRows);
        if (parsed.slotColumnIndices().isEmpty()) {
            Label msg =
                    new Label(
                            """
                            このシートから時刻列（列見出しが HH:MM 形式）を検出できませんでした。
                            「結果_設備ガント」形式の JSON を開いているか確認してください。""");
            msg.setWrapText(true);
            msg.setPadding(new Insets(16));
            root.setCenter(msg);
            return root;
        }

        LayoutMetrics layout = LayoutMetrics.fromZoom(zoom);
        GanttPalette palette = GanttPalette.forTheme(theme);

        double machW = clampMachineColumnWidth(machineColWidth);
        double procW = clampProcessColumnWidth(processColWidth);
        double leftTotal = machW + procW;

        double timelineWidth = parsed.slotColumnIndices().size() * layout.slotWidth;

        Canvas headerCanvas = new Canvas(timelineWidth, layout.headerHeight);
        drawTimeAxis(headerCanvas.getGraphicsContext2D(), parsed, timelineWidth, layout, palette);

        Label hMach = new Label("機械名");
        Label hProc = new Label("工程名");
        applySideHeaderStyle(hMach, machW, layout, palette);
        applySideHeaderStyle(hProc, procW, layout, palette);

        VBox wrapMach = new VBox(hMach);
        VBox wrapProc = new VBox(hProc);
        SplitPane headSplit = new SplitPane(wrapMach, wrapProc);
        headSplit.setMinWidth(leftTotal);
        headSplit.setPrefWidth(leftTotal);
        headSplit.setMaxHeight(layout.headerHeight);
        wrapMach.setMinHeight(layout.headerHeight);
        wrapProc.setMinHeight(layout.headerHeight);
        styleEquipmentHeadSplit(headSplit, theme);
        double divPos = leftTotal > 1 ? machW / leftTotal : 0.5;
        AtomicBoolean suppressDividerCallback = new AtomicBoolean(true);
        headSplit.setDividerPositions(divPos);
        suppressDividerCallback.set(false);

        if (onLeftColumnWidthsChanged != null) {
            headSplit
                    .getDividers()
                    .get(0)
                    .positionProperty()
                    .addListener(
                            (obs, o, n) -> {
                                if (suppressDividerCallback.get()) {
                                    return;
                                }
                                double tw = headSplit.getWidth();
                                if (tw < 40) {
                                    return;
                                }
                                double pos = n.doubleValue();
                                double newM = tw * pos;
                                double newP = tw * (1.0 - pos);
                                if (newM >= MIN_SIDE_COL_WIDTH && newP >= MIN_SIDE_COL_WIDTH) {
                                    onLeftColumnWidthsChanged.accept(newM, newP);
                                }
                            });
        }

        int progCell = layout.progressCellWidth;
        int gap = layout.progressGap;
        int progressTotal =
                parsed.progressColumnIndices().size() * progCell
                        + Math.max(0, parsed.progressColumnIndices().size() - 1) * gap;

        double contentMinWidth = leftTotal + timelineWidth + progressTotal;

        HBox headRow = new HBox(0, headSplit, headerCanvas);
        headRow.setMinWidth(contentMinWidth);
        headSplit.setMinWidth(leftTotal);
        HBox.setHgrow(headSplit, Priority.NEVER);

        VBox scrollBody = new VBox(0);
        scrollBody.setMinWidth(contentMinWidth);

        int dataStripe = 0;
        for (int ri = 0; ri < parsed.displayRows().size(); ri++) {
            DisplayRow dr = parsed.displayRows().get(ri);
            if (dr.sectionBanner() != null) {
                Label ban = new Label(dr.sectionBanner());
                ban.setPrefHeight(layout.sectionRowHeight);
                ban.setMinHeight(layout.sectionRowHeight);
                ban.setMaxWidth(Double.MAX_VALUE);
                ban.setAlignment(Pos.CENTER_LEFT);
                ban.setPadding(new Insets(2 * layout.zoom, 8 * layout.zoom, 2 * layout.zoom, 8 * layout.zoom));
                ban.setStyle(palette.sectionBannerCss());
                ban.setMinWidth(contentMinWidth);
                scrollBody.getChildren().add(ban);
                continue;
            }

            String machTxt = dr.machineLine() != null ? dr.machineLine() : "";
            String procTxt = dr.processBlock() != null ? dr.processBlock() : "";

            Label ml = new Label(machTxt);
            Label pl = new Label(procTxt);
            applySideDataStyle(ml, machW, layout, palette);
            applySideDataStyle(pl, procW, layout, palette);
            ml.setWrapText(true);
            pl.setWrapText(true);

            fitFontIntoColumn(ml, machTxt, machW - 8, layout.rowHeight - 4, layout.rowLabelFontSize);
            fitFontIntoColumn(pl, procTxt, procW - 8, layout.rowHeight - 4, layout.rowLabelFontSize);

            Canvas rowCanvas = new Canvas(timelineWidth, layout.rowHeight);
            drawTimelineRow(
                    rowCanvas.getGraphicsContext2D(), dr.cellsInSlots(), dataStripe++, layout, palette);

            String tip =
                    (dr.rowSummary() != null ? dr.rowSummary() : "")
                            + "\n（スロット "
                            + parsed.slotMinutes()
                            + " 分・設備ガント JSON と同一データ）";
            Tooltip.install(rowCanvas, new Tooltip(tip));

            HBox progBox = new HBox(gap);
            progBox.setAlignment(Pos.CENTER_LEFT);
            for (int pc : parsed.progressColumnIndices()) {
                String pv =
                        dr.rawRow().size() > pc && dr.rawRow().get(pc) != null
                                ? dr.rawRow().get(pc).strip()
                                : "";
                Label pLab = new Label(pv);
                pLab.setMinWidth(progCell);
                pLab.setPrefWidth(progCell);
                pLab.setMaxWidth(progCell);
                pLab.setWrapText(true);
                pLab.setAlignment(Pos.CENTER);
                pLab.setFont(Font.font(layout.progressFontSize));
                pLab.setStyle(pv.isEmpty() ? palette.progressEmptyCss() : palette.progressFilledCss());
                progBox.getChildren().add(pLab);
            }

            HBox line = new HBox(0);
            line.setMinWidth(contentMinWidth);
            line.getChildren().addAll(ml, pl, rowCanvas);
            if (!progBox.getChildren().isEmpty()) {
                line.getChildren().add(progBox);
            }
            scrollBody.getChildren().add(line);
        }

        ScrollPane headerScroll = new ScrollPane(headRow);
        headerScroll.setHbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);
        headerScroll.setVbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);
        headerScroll.setPannable(false);
        headerScroll.setFitToHeight(false);
        headerScroll.setPrefViewportHeight(layout.headerHeight);
        headerScroll.setPrefHeight(layout.headerHeight);
        headerScroll.setMinHeight(layout.headerHeight);
        headerScroll.setMaxHeight(layout.headerHeight);
        headerScroll.setMinWidth(Region.USE_COMPUTED_SIZE);

        ScrollPane bodyScroll = new ScrollPane(scrollBody);
        bodyScroll.setFitToWidth(true);
        bodyScroll.setPannable(true);
        VBox.setVgrow(bodyScroll, Priority.ALWAYS);

        headerScroll.hvalueProperty().bindBidirectional(bodyScroll.hvalueProperty());

        VBox mainColumn = new VBox(0, headerScroll, bodyScroll);
        VBox.setVgrow(bodyScroll, Priority.ALWAYS);
        mainColumn.setPadding(new Insets(4));
        root.setCenter(mainColumn);

        attachSplitDividerGrip(headSplit, theme);

        Label hint =
                new Label(
                        """
                        ヒント: 横スクロールで時刻軸を追えます。Ctrl+ホイールで表示倍率。 \
                        機械名／工程名の間の縦線をドラッグして列幅変更（自動保存）。 \
                        ブロックは Excel と同じセル内容を連続結合した帯です。""");
        hint.setWrapText(true);
        hint.setStyle(palette.hintCss());
        hint.setPadding(new Insets(0, 8, 8, 8));
        root.setBottom(hint);
        return root;
    }

    private static void styleEquipmentHeadSplit(SplitPane sp, DesktopTheme theme) {
        boolean dark = theme.isDarkUi();
        String line = dark ? "#94a3b8" : "#64748b";
        sp.setStyle(
                "-fx-background-color: transparent; -fx-box-border: "
                        + line
                        + "; -fx-background-insets: 0; -fx-padding: 0;");
    }

    /** テーマ適用後に境界線を太め・ハイコントラストにする（初期レイアウト後）。 */
    private static void attachSplitDividerGrip(SplitPane sp, DesktopTheme theme) {
        String bg = theme.isDarkUi() ? "rgba(148,163,184,0.98)" : "rgba(71,85,105,0.95)";
        sp.sceneProperty()
                .addListener(
                        (obs, o, scene) -> {
                            if (scene == null) {
                                return;
                            }
                            Platform.runLater(
                                    () -> {
                                        for (Node n : sp.lookupAll(".split-pane-divider")) {
                                            n.setStyle(
                                                    "-fx-background-color: "
                                                            + bg
                                                            + "; -fx-padding: 0 5 0 5; -fx-cursor: h-resize;");
                                        }
                                    });
                        });
    }

    private static void applySideHeaderStyle(
            Label lb, double colW, LayoutMetrics layout, GanttPalette palette) {
        lb.setMinWidth(colW);
        lb.setPrefWidth(colW);
        lb.setMaxWidth(colW);
        lb.setAlignment(Pos.CENTER_LEFT);
        lb.setPadding(new Insets(4 * layout.zoom, 6 * layout.zoom, 4 * layout.zoom, 6 * layout.zoom));
        lb.setFont(Font.font(layout.rowLabelFontSize * 1.05));
        lb.setStyle(palette.rowLabelCss());
        lb.setMinHeight(layout.headerHeight);
        lb.setPrefHeight(layout.headerHeight);
        lb.setMaxHeight(layout.headerHeight);
    }

    private static void applySideDataStyle(
            Label lb, double colW, LayoutMetrics layout, GanttPalette palette) {
        lb.setMinWidth(colW);
        lb.setPrefWidth(colW);
        lb.setMaxWidth(colW);
        lb.setAlignment(Pos.CENTER_LEFT);
        lb.setPadding(new Insets(2 * layout.zoom, 6 * layout.zoom, 2 * layout.zoom, 6 * layout.zoom));
        lb.setStyle(palette.rowLabelCss());
        lb.setMinHeight(layout.rowHeight);
        lb.setPrefHeight(layout.rowHeight);
        lb.setMaxHeight(layout.rowHeight);
    }

    private static double prefWrappedLabelHeight(String text, Font font, double wrapWidth) {
        String t = text != null ? text : "";
        Label tmp = new Label(t);
        tmp.setWrapText(true);
        tmp.setFont(font);
        tmp.setMaxWidth(Math.max(8, wrapWidth));
        MEASURE_ROOT.getChildren().setAll(tmp);
        tmp.applyCss();
        tmp.layout();
        return tmp.prefHeight(wrapWidth);
    }

    private static void fitFontIntoColumn(
            Label lb, String text, double innerWidth, double maxHeight, double maxFontPx) {
        String t = text != null ? text : "";
        lb.setText(t);
        if (t.isEmpty()) {
            lb.setFont(Font.font(Math.max(6, maxFontPx * 0.6)));
            return;
        }
        double lo = Math.max(6, maxFontPx * 0.35);
        double hi = maxFontPx;
        Font best = Font.font(lo);
        double w = Math.max(8, innerWidth);
        double maxH = Math.max(8, maxHeight);
        for (int iter = 0; iter < 22; iter++) {
            double mid = (lo + hi) / 2;
            Font f = Font.font(mid);
            double h = prefWrappedLabelHeight(t, f, w);
            if (h <= maxH) {
                best = f;
                lo = mid;
            } else {
                hi = mid;
            }
        }
        lb.setFont(best);
    }

    private record LayoutMetrics(
            double zoom,
            double slotWidth,
            double rowHeight,
            double sectionRowHeight,
            double headerHeight,
            double labelMinWidth,
            double labelMaxWidth,
            double rowLabelFontSize,
            double axisFontSize,
            double barFontSize,
            double progressFontSize,
            int progressCellWidth,
            int progressGap) {

        static LayoutMetrics fromZoom(double zoomIn) {
            double z = Math.clamp(zoomIn, 0.5, 2.0);
            return new LayoutMetrics(
                    z,
                    BASE_SLOT_WIDTH * z,
                    BASE_ROW_HEIGHT * z,
                    BASE_SECTION_ROW_HEIGHT * z,
                    BASE_HEADER_HEIGHT * z,
                    BASE_LABEL_MIN_WIDTH * z,
                    BASE_LABEL_MAX_WIDTH * z,
                    11 * z,
                    10 * z,
                    Math.max(8, 9 * z),
                    Math.max(8, 9 * z),
                    (int) Math.round(52 * z),
                    (int) Math.round(4 * z));
        }
    }

    private record GanttPalette(
            Color emptyLight,
            Color emptyBand,
            Color grid,
            Color headerAxis,
            Color axisLabel,
            Color barDefault,
            Color barBreak,
            Color barStartup,
            Color barDefaultText,
            Color barBreakText,
            Color barStartupText,
            Color barStroke,
            String rowLabelCss,
            String sectionBannerCss,
            String hintCss,
            String progressEmptyCss,
            String progressFilledCss) {

        static GanttPalette forTheme(DesktopTheme theme) {
            boolean dark = theme.isDarkUi();
            if (dark) {
                return new GanttPalette(
                        Color.web("#1e293b"),
                        Color.web("#0f172a"),
                        Color.web("#475569"),
                        Color.web("#334155"),
                        Color.web("#e2e8f0"),
                        Color.web("#2563eb"),
                        Color.web("#38bdf8"),
                        Color.web("#fb923c"),
                        Color.web("#f8fafc"),
                        Color.web("#0c1222"),
                        Color.web("#0c1222"),
                        Color.web("#93c5fd"),
                        "-fx-background-color: #334155; -fx-text-fill: #f8fafc; "
                                + "-fx-border-color: #64748b; -fx-border-width: 0 1 0 0;",
                        "-fx-background-color: #0f172a; -fx-font-weight: bold; -fx-font-size: 12px; "
                                + "-fx-text-fill: #f1f5f9;",
                        "-fx-text-fill: #94a3b8; -fx-font-size: 11px;",
                        "-fx-background-color: #1e293b; -fx-border-color: #475569; -fx-text-fill: #cbd5e1;",
                        "-fx-background-color: #422006; -fx-border-color: #d97706; -fx-text-fill: #ffedd5;");
            }
            return new GanttPalette(
                    Color.web("#ffffff"),
                    Color.web("#f1f5f9"),
                    Color.web("#cbd5e1"),
                    Color.web("#e2e8f0"),
                    Color.web("#0f172a"),
                    Color.web("#2563eb"),
                    Color.web("#7dd3fc"),
                    Color.web("#fdba74"),
                    Color.web("#ffffff"),
                    Color.web("#0f172a"),
                    Color.web("#9a3412"),
                    Color.web("#1e40af"),
                    "-fx-background-color: #e8eef7; -fx-text-fill: #0f172a; "
                            + "-fx-border-color: #cbd5e1; -fx-border-width: 0 1 0 0;",
                    "-fx-background-color: #1e3a5f; -fx-font-weight: bold; -fx-font-size: 12px; "
                            + "-fx-text-fill: #f8fafc;",
                    "-fx-text-fill: #64748b; -fx-font-size: 11px;",
                    "-fx-background-color: #fffbf0; -fx-border-color: #f0e1b7; -fx-text-fill: #334155;",
                    "-fx-background-color: #fff2cc; -fx-border-color: #d6b656; -fx-text-fill: #334155;");
        }
    }

    private static void drawTimeAxis(
            GraphicsContext gc,
            ParseResult parsed,
            double timelineWidth,
            LayoutMetrics layout,
            GanttPalette palette) {
        gc.setFill(palette.headerAxis());
        gc.fillRect(0, 0, timelineWidth, layout.headerHeight);
        gc.setStroke(palette.grid());
        gc.setLineWidth(0.5);
        gc.strokeRect(0, 0, timelineWidth, layout.headerHeight);

        List<Integer> slotCols = parsed.slotColumnIndices();
        int n = slotCols.size();
        int step = Math.max(1, 60 / Math.max(1, parsed.slotMinutes()));

        gc.setFill(palette.axisLabel());
        gc.setFont(Font.font(layout.axisFontSize));
        LocalTime t0 = parsed.slotBaseTime();
        for (int i = 0; i < n; i += step) {
            double x = i * layout.slotWidth;
            LocalTime tt = t0.plusMinutes((long) i * parsed.slotMinutes());
            String txt = tt.format(DateTimeFormatter.ofPattern("H:mm"));
            gc.fillText(txt, x + 2, layout.headerHeight - 8 * layout.zoom);
            gc.strokeLine(x, 0, x, layout.headerHeight);
        }
    }

    private static void drawTimelineRow(
            GraphicsContext gc,
            List<String> slotTexts,
            int stripeIndex,
            LayoutMetrics layout,
            GanttPalette palette) {
        int n = slotTexts.size();
        boolean stripe = (stripeIndex & 1) == 0;
        for (int i = 0; i < n; i++) {
            double x = i * layout.slotWidth;
            gc.setFill(stripe ? palette.emptyLight() : palette.emptyBand());
            gc.fillRect(x, 0, layout.slotWidth, layout.rowHeight);
        }
        gc.setStroke(palette.grid());
        gc.setLineWidth(0.3);
        for (int i = 0; i <= n; i++) {
            gc.strokeLine(i * layout.slotWidth, 0, i * layout.slotWidth, layout.rowHeight);
        }

        int runStart = -1;
        String runText = "";
        for (int i = 0; i < n; i++) {
            String t = slotTexts.get(i) != null ? slotTexts.get(i).strip() : "";
            boolean empty = t.isEmpty();
            if (empty) {
                if (runStart >= 0) {
                    paintBar(gc, runStart, i - 1, runText, layout, palette);
                    runStart = -1;
                    runText = "";
                }
                continue;
            }
            if (runStart < 0) {
                runStart = i;
                runText = t;
            } else if (!t.equals(runText)) {
                paintBar(gc, runStart, i - 1, runText, layout, palette);
                runStart = i;
                runText = t;
            }
        }
        if (runStart >= 0) {
            paintBar(gc, runStart, n - 1, runText, layout, palette);
        }
    }

    private static void paintBar(
            GraphicsContext gc,
            int fromSlot,
            int toSlot,
            String text,
            LayoutMetrics layout,
            GanttPalette palette) {
        double x = fromSlot * layout.slotWidth;
        double w = (toSlot - fromSlot + 1) * layout.slotWidth;
        BarKind kind = classifyBar(text);
        Color fill =
                switch (kind) {
                    case BREAK -> palette.barBreak();
                    case STARTUP -> palette.barStartup();
                    default -> palette.barDefault();
                };
        gc.setFill(fill);
        double arc = Math.max(2, 3 * layout.zoom);
        double inset = 0.5 * layout.zoom;
        double barTop = 3 * layout.zoom;
        double barH = layout.rowHeight - 2 * barTop;
        gc.fillRoundRect(x + inset, barTop, w - 2 * inset, barH, arc, arc);
        gc.setStroke(palette.barStroke());
        gc.setLineWidth(0.5 * layout.zoom);
        gc.strokeRoundRect(x + inset, barTop, w - 2 * inset, barH, arc, arc);

        Color label =
                switch (kind) {
                    case BREAK -> palette.barBreakText();
                    case STARTUP -> palette.barStartupText();
                    default -> palette.barDefaultText();
                };
        gc.setFill(label);
        gc.setFont(Font.font(layout.barFontSize));
        String shortTxt = text.replace('\n', ' ');
        if (shortTxt.length() > 80) {
            shortTxt = shortTxt.substring(0, 77) + "...";
        }
        double charPx = Math.max(4.0, 5.0 * layout.zoom);
        double maxChars = Math.max(4, (w - 6 * layout.zoom) / charPx);
        if (shortTxt.length() > maxChars) {
            shortTxt = shortTxt.substring(0, (int) maxChars - 2) + "..";
        }
        if (w > 28 * layout.zoom && !shortTxt.isEmpty()) {
            gc.fillText(shortTxt, x + 4 * layout.zoom, layout.rowHeight / 2 + 3 * layout.zoom);
        }
    }

    private enum BarKind {
        DEFAULT,
        BREAK,
        STARTUP
    }

    private static BarKind classifyBar(String t) {
        if (t.contains("休憩") || t.contains("（休憩）")) {
            return BarKind.BREAK;
        }
        if (t.contains("日次始業準備")) {
            return BarKind.STARTUP;
        }
        return BarKind.DEFAULT;
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

    private record RepairResult(
            List<String> columns, ObservableList<ObservableList<String>> rows) {}

    private static ParseResult parse(List<String> columns, ObservableList<ObservableList<String>> rows) {
        List<Integer> slots = new ArrayList<>();
        for (int c = 0; c < columns.size(); c++) {
            String h = columns.get(c);
            if (parseTimeHeader(h) != null) {
                slots.add(c);
            }
        }
        LocalTime slotBaseTime = LocalTime.of(8, 0);
        if (!slots.isEmpty()) {
            LocalTime t0 = parseTimeHeader(columns.get(slots.get(0)));
            if (t0 != null) {
                slotBaseTime = t0;
            }
        }
        int slotMinutes = 10;
        if (slots.size() >= 2) {
            LocalTime a = parseTimeHeader(columns.get(slots.get(0)));
            LocalTime b = parseTimeHeader(columns.get(slots.get(1)));
            if (a != null && b != null) {
                int delta =
                        (b.getHour() * 60 + b.getMinute()) - (a.getHour() * 60 + a.getMinute());
                if (delta > 0) {
                    slotMinutes = delta;
                }
            }
        }
        List<Integer> progressCols = new ArrayList<>();
        for (int c = 0; c < columns.size(); c++) {
            String h = columns.get(c);
            if (h != null && h.endsWith("進度")) {
                progressCols.add(c);
            }
        }

        List<DisplayRow> displayRows = new ArrayList<>();
        String carriedDate = "";
        for (int r = 0; r < rows.size(); r++) {
            ObservableList<String> row = rows.get(r);
            if (row == null || row.isEmpty()) {
                continue;
            }
            String c0 = row.size() > 0 && row.get(0) != null ? row.get(0).strip() : "";
            if (isSectionRow(row)) {
                String banner = !c0.isEmpty() ? c0 : sectionTitleFromRow(row);
                displayRows.add(new DisplayRow(banner, null, null, null, null, row));
                continue;
            }
            int dateCol = columns.indexOf("日付");
            if (dateCol >= 0 && row.size() > dateCol) {
                String dv = row.get(dateCol) != null ? row.get(dateCol).strip() : "";
                if (!dv.isEmpty()) {
                    carriedDate = dv;
                }
            }
            LeftParts lp = buildLeftParts(columns, row, carriedDate);
            List<String> slotCells = new ArrayList<>();
            for (int si : slots) {
                String v =
                        row.size() > si && row.get(si) != null ? row.get(si) : "";
                slotCells.add(v);
            }
            displayRows.add(
                    new DisplayRow(
                            null,
                            lp.machine(),
                            lp.process(),
                            lp.summary(),
                            slotCells,
                            row));
        }
        return new ParseResult(slots, progressCols, slotMinutes, slotBaseTime, displayRows);
    }

    private static boolean isSectionRow(ObservableList<String> row) {
        for (int i = 0; i < Math.min(4, row.size()); i++) {
            String s = row.get(i) != null ? row.get(i) : "";
            if (s.contains("■") || s.contains("▪")) {
                return true;
            }
            if (s.contains("【")) {
                if (BRACKETED_PLAIN_DATE_LABEL.matcher(s.strip()).matches()) {
                    continue;
                }
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
        return columns.indexOf(name);
    }

    private record LeftParts(String machine, String process, String summary) {}

    /**
     * 機械名列・工程名列に分割する。日付は工程名ではなく機械名側に表示する。
     * ツールチップ用の要約文字列も返す。
     */
    private static LeftParts buildLeftParts(
            List<String> columns, ObservableList<String> row, String carriedDate) {
        String mach = cellAt(columns, row, "機械名");
        String proc = cellAt(columns, row, "工程名");
        String task = cellAt(columns, row, "タスク概覝");
        String tb = cellAt(columns, row, "日時帯");
        String dateLine = compactDateLine(carriedDate);

        StringBuilder machCol = new StringBuilder();
        if (!mach.isEmpty()) {
            machCol.append(mach);
        }
        if (!dateLine.isEmpty()) {
            if (machCol.length() > 0) {
                machCol.append('\n');
            }
            machCol.append(dateLine);
        }

        StringBuilder procCol = new StringBuilder();
        if (!proc.isEmpty() && !proc.equals("—")) {
            procCol.append(proc);
        }
        if (!task.isEmpty() && !task.equals("—")) {
            String tk = task.length() > 120 ? task.substring(0, 117) + "..." : task;
            if (procCol.length() > 0) {
                procCol.append('\n');
            }
            procCol.append(tk);
        }
        if (!tb.isEmpty() && procCol.length() == 0 && machCol.length() == 0) {
            procCol.append(tb);
        }

        String machStr = machCol.toString();
        String procStr = procCol.toString();

        StringBuilder sum = new StringBuilder();
        if (!machStr.isEmpty()) {
            sum.append(machStr);
        }
        if (!procStr.isEmpty()) {
            if (sum.length() > 0) {
                sum.append('\n');
            }
            sum.append(procStr);
        }
        String summary = sum.toString();
        if (summary.isEmpty()) {
            return new LeftParts("", "", "（行）");
        }
        return new LeftParts(machStr, procStr, summary);
    }

    /** 「【2026/05/07】」等を {@code 2026/05/07} 形式に正規化する。 */
    private static String compactDateLine(String raw) {
        if (raw == null || raw.isBlank()) {
            return "";
        }
        String s = raw.strip();
        if (s.startsWith("【")) {
            int end = s.indexOf('】');
            if (end > 1) {
                s = s.substring(1, end).strip();
            }
        }
        Matcher m = LOOSE_YMD.matcher(s);
        if (m.find()) {
            int y = Integer.parseInt(m.group(1));
            int mo = Integer.parseInt(m.group(2));
            int d = Integer.parseInt(m.group(3));
            return String.format("%d/%02d/%02d", y, mo, d);
        }
        return s;
    }

    private static String cellAt(List<String> columns, ObservableList<String> row, String colName) {
        int i = columnIndex(columns, colName);
        if (i < 0 || row.size() <= i) {
            return "";
        }
        String v = row.get(i);
        return v != null ? v.strip() : "";
    }

    private record ParseResult(
            List<Integer> slotColumnIndices,
            List<Integer> progressColumnIndices,
            int slotMinutes,
            LocalTime slotBaseTime,
            List<DisplayRow> displayRows) {}

    /**
     * @param sectionBanner 非 null のときセクション行
     * @param machineLine データ行: 機械名列のテキスト
     * @param processBlock データ行: 工程名列（工程名・タスク等。日付は機械名列側）
     * @param rowSummary データ行: ツールチップ用
     */
    private record DisplayRow(
            String sectionBanner,
            String machineLine,
            String processBlock,
            String rowSummary,
            List<String> cellsInSlots,
            ObservableList<String> rawRow) {}
}
