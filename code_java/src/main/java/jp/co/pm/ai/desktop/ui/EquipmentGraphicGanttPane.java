package jp.co.pm.ai.desktop.ui;

import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.Locale;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
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
import javafx.scene.control.Tooltip;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Font;
import javafx.scene.text.Text;
import javafx.geometry.VPos;
import javafx.scene.input.ScrollEvent;

import javafx.scene.layout.Pane;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;

import jp.co.pm.ai.desktop.config.DesktopTheme;
import java.util.function.Function;

import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.io.gantt.PersonNameBadgeText;

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
    /** 時刻ヘッダを縦書き表示するため Excel 風にやや高め */
    private static final double BASE_HEADER_HEIGHT = 54;
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

    /** UI からの列幅上書き（px）。正の有限値のみ採用し、それ以外は計測幅を使う */
    private static double effectiveLeftColWidth(double measuredW, double overridePx) {
        if (Double.isFinite(overridePx) && overridePx > 0.5) {
            return Math.min(MAX_SIDE_COL_WIDTH, Math.max(MIN_SIDE_COL_WIDTH, overridePx));
        }
        return measuredW;
    }

    private record MeasuredLeftWidths(double dateW, double machW, double procW) {}

    private static MeasuredLeftWidths measureAutoLeftWidths(
            List<String> columns,
            ParseResult parsed,
            List<MachineColumnPlan> machPlans,
            LayoutMetrics layout) {
        Font headerFont = Font.font(layout.rowLabelFontSize * 1.05);
        Font cellFont = Font.font(layout.rowLabelFontSize);

        /* 日付列は見出しテキストなし。幅は縦書き日付（データ）のフォント＋左右パディング程度 */
        double dateBodyFontPx = layout.rowLabelFontSize * 0.92;
        double dateCol =
                Math.min(
                        MAX_SIDE_COL_WIDTH,
                        Math.max(
                                MIN_SIDE_COL_WIDTH,
                                Math.ceil(dateBodyFontPx + 12 * layout.zoom)));

        /* 見出し・セルは Label＋実際の Insets で測る（Text より実表示に近い）。列パディングは二重に足さない */
        double maxM = measureSideHeaderLabelPrefWidth("機械名", layout, headerFont);
        double maxP = measureSideHeaderLabelPrefWidth("工程名", layout, headerFont);

        List<DisplayRow> rows = parsed.displayRows();
        for (int i = 0; i < rows.size(); i++) {
            DisplayRow dr = rows.get(i);
            if (dr.sectionBanner() != null) {
                continue;
            }
            MachineColumnPlan plan = machPlans.get(i);
            String proc = dr.processBlock() != null ? dr.processBlock() : "";
            maxP = Math.max(maxP, measureMultilineMaxSideDataLabelWidth(proc, layout, cellFont));

            if (plan != null && !plan.continuation()) {
                String mach = plan.machineCellText() != null ? plan.machineCellText() : "";
                maxM = Math.max(maxM, measureMultilineMaxSideDataLabelWidth(mach, layout, cellFont));
            }
        }

        double machCol =
                Math.min(MAX_SIDE_COL_WIDTH, Math.max(MIN_SIDE_COL_WIDTH, maxM));
        double procCol =
                Math.min(MAX_SIDE_COL_WIDTH, Math.max(MIN_SIDE_COL_WIDTH, maxP));
        return new MeasuredLeftWidths(dateCol, machCol, procCol);
    }

    private static ColumnConstraints fixedPixelColumn(double w) {
        ColumnConstraints c = new ColumnConstraints(w, w, w);
        c.setHgrow(Priority.NEVER);
        return c;
    }

    private static double measureSideHeaderLabelPrefWidth(
            String text, LayoutMetrics layout, Font font) {
        String t = text != null ? text : "";
        Label lb = new Label(t);
        lb.setFont(font);
        lb.setPadding(
                new Insets(
                        4 * layout.zoom,
                        6 * layout.zoom,
                        4 * layout.zoom,
                        6 * layout.zoom));
        lb.setWrapText(false);
        MEASURE_ROOT.getChildren().setAll(lb);
        lb.applyCss();
        lb.layout();
        return Math.ceil(lb.prefWidth(-1));
    }

    private static double measureSideDataLabelPrefWidth(
            String text, LayoutMetrics layout, Font font) {
        String collapsed = collapseWhitespaceForColumnMeasure(text != null ? text : "");
        Label lb = new Label(collapsed.isEmpty() ? " " : collapsed);
        lb.setFont(font);
        lb.setPadding(
                new Insets(
                        2 * layout.zoom,
                        6 * layout.zoom,
                        2 * layout.zoom,
                        6 * layout.zoom));
        lb.setWrapText(false);
        MEASURE_ROOT.getChildren().setAll(lb);
        lb.applyCss();
        lb.layout();
        return Math.ceil(lb.prefWidth(-1));
    }

    private static double measureMultilineMaxSideDataLabelWidth(
            String text, LayoutMetrics layout, Font cellFont) {
        String t = text != null ? text : "";
        double m = 0;
        for (String line : t.split("\\R")) {
            String s = collapseWhitespaceForColumnMeasure(line);
            if (s.isEmpty()) {
                continue;
            }
            m = Math.max(m, measureSideDataLabelPrefWidth(s, layout, cellFont));
        }
        return m;
    }

    /**
     * 列幅計測用。データに含まれる全角スペース・連続空白がそのままだと {@link Text} の見た目より列だけ広がるため、
     * {@link Character#isWhitespace(int)} / {@code \\p{Zs}} に該当する連続を半角1つに畳む（表示文字列は変えない）。
     */
    private static String collapseWhitespaceForColumnMeasure(String text) {
        if (text == null || text.isEmpty()) {
            return "";
        }
        return text.strip().replaceAll("[\\p{Zs}]+", " ");
    }

    /**
     * @param columns シート列見出し
     * @param rows データ行（フィルタ行を含まない素データ）
     * @return 時刻列が検出できない場合は説明ラベルのみのペイン
     */
    public static BorderPane build(
            List<String> columns, ObservableList<ObservableList<String>> rows) {
        return build(columns, rows, DesktopTheme.LIGHT, 1.0, 100, 100, "", 100, 100);
    }

    public static BorderPane build(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            DesktopTheme theme,
            double zoom) {
        return build(columns, rows, theme, zoom, 100, 100, "", 100, 100);
    }

    /**
     * @param theme アプリの {@link DesktopTheme}（Canvas 帯の配色に反映）
     * @param zoom 表示倍率（0.5〜2.0、スライダー 100%÷100）
     * @param rowHeightPercent 行の高さ（50〜200、100＝既定）
     * @param slotWidthPercent 時刻列の幅スケール（50〜500）
     * @param barFontFamily バー上ラベル用フォントファミリ（null／空で既定）
     * @param barFontPercent バー内ラベル文字サイズ（50〜200、100＝既定）
     * @param headerHeightPercent 見出し行の高さ（50〜200、100＝既定）
     */
    public static BorderPane build(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            DesktopTheme theme,
            double zoom,
            double rowHeightPercent,
            double slotWidthPercent,
            String barFontFamily) {
        return build(
                columns,
                rows,
                theme,
                zoom,
                rowHeightPercent,
                slotWidthPercent,
                barFontFamily,
                100,
                100);
    }

    /** 見出し行高さは既定 100%（ヘッダ調整スライダーなしの呼び出し向け） */
    public static BorderPane build(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            DesktopTheme theme,
            double zoom,
            double rowHeightPercent,
            double slotWidthPercent,
            String barFontFamily,
            double barFontPercent) {
        return build(
                columns,
                rows,
                theme,
                zoom,
                rowHeightPercent,
                slotWidthPercent,
                barFontFamily,
                barFontPercent,
                100);
    }

    public static BorderPane build(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            DesktopTheme theme,
            double zoom,
            double rowHeightPercent,
            double slotWidthPercent,
            String barFontFamily,
            double barFontPercent,
            double headerHeightPercent) {
        return build(
                columns,
                rows,
                theme,
                zoom,
                rowHeightPercent,
                slotWidthPercent,
                barFontFamily,
                barFontPercent,
                headerHeightPercent,
                0d,
                0d,
                0d,
                0d,
                null,
                false,
                null);
    }

    /**
     * 担当バッジなし（後方互換）。
     *
     * @see #build(List, ObservableList, DesktopTheme, double, double, double, String, double, double, double, double,
     *     double, double, List, boolean, Function)
     */
    public static BorderPane build(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            DesktopTheme theme,
            double zoom,
            double rowHeightPercent,
            double slotWidthPercent,
            String barFontFamily,
            double barFontPercent,
            double headerHeightPercent,
            double dateColWidthOverridePx,
            double machineColWidthOverridePx,
            double processColWidthOverridePx,
            double shiftWheelHorizontalSensitivityPercent) {
        return build(
                columns,
                rows,
                theme,
                zoom,
                rowHeightPercent,
                slotWidthPercent,
                barFontFamily,
                barFontPercent,
                headerHeightPercent,
                dateColWidthOverridePx,
                machineColWidthOverridePx,
                processColWidthOverridePx,
                shiftWheelHorizontalSensitivityPercent,
                null,
                false,
                null);
    }

    /**
     * @param dateColWidthOverridePx 日付列幅（px）。{@code <= 0} または非有限は自動計測を使用
     * @param machineColWidthOverridePx 機械名列幅（px）。同上
     * @param processColWidthOverridePx 工程名列幅（px）。同上
     * @param shiftWheelHorizontalSensitivityPercent Shift+ホイール横スクロール感度（50〜1000、100＝従来の速さ、{@code <=0}
     *     は既定 200）
     * @param badgeSlotRowsRaw 行インデックスが {@code rows} と一致する担当バッジグリッド（各内側リストはスロット列数分）。Excel 連携のみのときは
     *     {@code null}
     * @param showPersonBadges 担当バッジオーバーレイを描画するか
     * @param personBadgeStyleResolver バッジ表示文字列ごとの見た目（{@code null} は常に {@link PersonBadgeStyle#defaultStyle()}）
     */
    public static BorderPane build(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            DesktopTheme theme,
            double zoom,
            double rowHeightPercent,
            double slotWidthPercent,
            String barFontFamily,
            double barFontPercent,
            double headerHeightPercent,
            double dateColWidthOverridePx,
            double machineColWidthOverridePx,
            double processColWidthOverridePx,
            double shiftWheelHorizontalSensitivityPercent,
            List<List<String>> badgeSlotRowsRaw,
            boolean showPersonBadges,
            Function<String, PersonBadgeStyle> personBadgeStyleResolver) {
        BorderPane root = new BorderPane();
        root.setCache(false);
        List<String> effCols = columns;
        ObservableList<ObservableList<String>> effRows = rows;
        RepairResult repaired = tryRepairPandasUnnamedEquipmentTimeline(effCols, effRows);
        List<List<String>> badgeEff = badgeSlotRowsRaw;
        if (repaired != null) {
            effCols = repaired.columns();
            effRows = repaired.rows();
            badgeEff = null;
        }
        Function<String, PersonBadgeStyle> badgeResolver =
                personBadgeStyleResolver != null
                        ? personBadgeStyleResolver
                        : (String __) -> PersonBadgeStyle.defaultStyle();
        ParseResult parsed = parse(effCols, effRows, badgeEff);
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

        LayoutMetrics layout =
                LayoutMetrics.fromScales(
                        zoom,
                        rowHeightPercent,
                        slotWidthPercent,
                        barFontPercent,
                        headerHeightPercent);
        GanttPalette palette = GanttPalette.forTheme(theme);
        Font barFont = resolveBarFont(barFontFamily, layout.barFontSize);

        List<MachineColumnPlan> machPlans =
                computeMachineColumnPlans(effCols, parsed.displayRows());
        List<DateColumnPlan> datePlans = computeDateColumnPlans(effCols, parsed.displayRows());
        MeasuredLeftWidths auto =
                measureAutoLeftWidths(effCols, parsed, machPlans, layout);
        double dateW = effectiveLeftColWidth(auto.dateW(), dateColWidthOverridePx);
        double machW = effectiveLeftColWidth(auto.machW(), machineColWidthOverridePx);
        double procW = effectiveLeftColWidth(auto.procW(), processColWidthOverridePx);
        double leftTotal = dateW + machW + procW;

        double timelineWidth = parsed.slotColumnIndices().size() * layout.slotWidth;
        /*
         * 幅・高さが 0 に近い／Prism の Canvas バックバッファ生成失敗時に NGCanvas で NPE になるのを避ける。
         * GPU ドライバ由来の不具合が残る場合は JVM に -Dprism.order=sw を試す。
         */
        final double canvasTimelineW = Math.max(1.0, timelineWidth);
        final double canvasHeaderH = Math.max(1.0, layout.headerHeight);

        Canvas headerCanvas = new Canvas(canvasTimelineW, canvasHeaderH);
        headerCanvas.setCache(false);
        drawTimeAxis(headerCanvas.getGraphicsContext2D(), parsed, canvasTimelineW, layout, palette);

        Label hDate = new Label("");
        Label hMach = new Label("機械名");
        Label hProc = new Label("工程名");
        applySideHeaderStyle(hDate, dateW, layout, palette);
        applySideHeaderStyle(hMach, machW, layout, palette);
        applySideHeaderStyle(hProc, procW, layout, palette);

        VBox wrapDate = new VBox(hDate);
        wrapDate.setAlignment(Pos.CENTER);
        wrapDate.setMinWidth(dateW);
        wrapDate.setPrefWidth(dateW);
        wrapDate.setMaxWidth(dateW);
        wrapDate.setMinHeight(layout.headerHeight);
        VBox wrapMach = new VBox(hMach);
        VBox wrapProc = new VBox(hProc);
        wrapMach.setAlignment(Pos.CENTER);
        wrapProc.setAlignment(Pos.CENTER);
        wrapMach.setMinWidth(machW);
        wrapMach.setPrefWidth(machW);
        wrapMach.setMaxWidth(machW);
        wrapProc.setMinWidth(procW);
        wrapProc.setPrefWidth(procW);
        wrapProc.setMaxWidth(procW);
        wrapMach.setMinHeight(layout.headerHeight);
        wrapProc.setMinHeight(layout.headerHeight);
        HBox leftHead = new HBox(0, wrapDate, wrapMach, wrapProc);
        leftHead.setMinWidth(leftTotal);
        leftHead.setPrefWidth(leftTotal);
        leftHead.setMaxHeight(layout.headerHeight);

        int progCell = layout.progressCellWidth;
        int gap = layout.progressGap;
        int progressTotal =
                parsed.progressColumnIndices().size() * progCell
                        + Math.max(0, parsed.progressColumnIndices().size() - 1) * gap;

        GridPane leftBodyGrid = new GridPane();
        leftBodyGrid.setMinWidth(leftTotal);
        leftBodyGrid.setMaxWidth(leftTotal);
        leftBodyGrid.setPrefWidth(leftTotal);
        leftBodyGrid.getColumnConstraints()
                .setAll(fixedPixelColumn(dateW), fixedPixelColumn(machW), fixedPixelColumn(procW));

        GridPane rightBodyGrid = new GridPane();
        rightBodyGrid.setMinWidth(timelineWidth + progressTotal);
        if (progressTotal > 0) {
            rightBodyGrid.getColumnConstraints()
                    .setAll(fixedPixelColumn(timelineWidth), fixedPixelColumn(progressTotal));
        } else {
            rightBodyGrid.getColumnConstraints().setAll(fixedPixelColumn(timelineWidth));
        }
        leftBodyGrid.setCache(false);
        rightBodyGrid.setCache(false);

        double timelineOuterPad =
                Math.min(
                        layout.rowHeight * 0.32,
                        Math.max(5 * layout.zoom, barFont.getSize() * 0.9));
        double cellBodyH = layout.rowHeight + 2 * timelineOuterPad;

        int machineColorSeq = -1;
        int gridR = 0;
        for (int ri = 0; ri < parsed.displayRows().size(); ri++) {
            DisplayRow dr = parsed.displayRows().get(ri);
            if (dr.sectionBanner() != null) {
                Label banL = new Label(dr.sectionBanner());
                banL.setPrefHeight(layout.sectionRowHeight);
                banL.setMinHeight(layout.sectionRowHeight);
                banL.setMaxHeight(layout.sectionRowHeight);
                banL.setMaxWidth(leftTotal);
                banL.setAlignment(Pos.CENTER_LEFT);
                banL.setPadding(new Insets(2 * layout.zoom, 8 * layout.zoom, 2 * layout.zoom, 8 * layout.zoom));
                banL.setStyle(palette.sectionBannerCss());
                banL.setWrapText(true);
                GridPane.setColumnSpan(banL, 3);
                leftBodyGrid.add(banL, 0, gridR);

                Region banR = new Region();
                banR.setPrefHeight(layout.sectionRowHeight);
                banR.setMinHeight(layout.sectionRowHeight);
                banR.setMaxHeight(layout.sectionRowHeight);
                banR.setMinWidth(timelineWidth + progressTotal);
                banR.setStyle(palette.sectionBannerCss());
                rightBodyGrid.add(banR, 0, gridR);
                if (progressTotal > 0) {
                    GridPane.setColumnSpan(banR, 2);
                }
                gridR++;
                continue;
            }

            MachineColumnPlan mplan = machPlans.get(ri);
            if (mplan == null) {
                mplan =
                        new MachineColumnPlan(
                                false, dr.machineLine() != null ? dr.machineLine() : "", 1);
            }

            if (!mplan.continuation()) {
                machineColorSeq++;
            }
            int machineGroupIndex = Math.max(0, machineColorSeq);

            String procTxt = dr.processBlock() != null ? dr.processBlock() : "";

            DateColumnPlan dplan = datePlans.get(ri);
            if (dplan == null) {
                dplan = new DateColumnPlan(false, "", 1);
            }

            if (!dplan.continuation()) {
                String dateTxt = dplan.dateText() != null ? dplan.dateText() : "";
                int dRows = Math.max(1, dplan.rowSpan());
                double spanDateH = dRows * cellBodyH;

                StackPane dateWrap = new StackPane();
                dateWrap.setMinWidth(dateW);
                dateWrap.setPrefWidth(dateW);
                dateWrap.setMaxWidth(dateW);
                dateWrap.setMinHeight(spanDateH);
                dateWrap.setPrefHeight(spanDateH);
                dateWrap.setMaxHeight(spanDateH);
                dateWrap.setPadding(
                        new Insets(
                                2 * layout.zoom,
                                6 * layout.zoom,
                                2 * layout.zoom,
                                6 * layout.zoom));
                dateWrap.setStyle(palette.machineSideCellCss(machineGroupIndex));
                if (dRows > 1) {
                    GridPane.setRowSpan(dateWrap, dRows);
                }
                GridPane.setValignment(dateWrap, VPos.TOP);

                Text dt = new Text(dateTxt != null ? dateTxt : "");
                dt.setFont(Font.font(layout.rowLabelFontSize * 0.92));
                dt.setFill(Color.web(palette.machineSideTextFill()));
                dt.setRotate(-90);
                StackPane.setAlignment(dt, Pos.CENTER);
                dateWrap.getChildren().add(dt);

                leftBodyGrid.add(dateWrap, 0, gridR);
            }

            if (!mplan.continuation()) {
                String machTxt = mplan.machineCellText() != null ? mplan.machineCellText() : "";
                Label ml = new Label(machTxt);
                applySideDataStyle(ml, machW, layout, palette, machineGroupIndex);
                ml.setWrapText(true);
                if (mplan.rowSpan() > 1) {
                    double spanH = mplan.rowSpan() * cellBodyH;
                    ml.setMinHeight(spanH);
                    ml.setPrefHeight(spanH);
                    ml.setMaxHeight(spanH);
                    ml.setAlignment(Pos.TOP_LEFT);
                    GridPane.setValignment(ml, VPos.TOP);
                    GridPane.setRowSpan(ml, mplan.rowSpan());
                    fitFontIntoColumn(
                            ml,
                            machTxt,
                            machW - 8,
                            spanH - 4,
                            layout.rowLabelFontSize);
                } else {
                    ml.setMinHeight(cellBodyH);
                    ml.setPrefHeight(cellBodyH);
                    ml.setMaxHeight(cellBodyH);
                    fitFontIntoColumn(
                            ml,
                            machTxt,
                            machW - 8,
                            cellBodyH - 4,
                            layout.rowLabelFontSize);
                }
                leftBodyGrid.add(ml, 1, gridR);
            }

            Label pl = new Label(procTxt);
            applySideDataStyle(pl, procW, layout, palette, machineGroupIndex);
            pl.setMinHeight(cellBodyH);
            pl.setPrefHeight(cellBodyH);
            pl.setMaxHeight(cellBodyH);
            pl.setWrapText(true);
            fitFontIntoColumn(pl, procTxt, procW - 8, cellBodyH - 4, layout.rowLabelFontSize);

            double rowCanvasH = Math.max(1.0, cellBodyH);
            Canvas rowCanvas = new Canvas(canvasTimelineW, rowCanvasH);
            rowCanvas.setCache(false);
            GraphicsContext gcx = rowCanvas.getGraphicsContext2D();
            gcx.translate(0, timelineOuterPad);
            drawTimelineRow(
                    gcx,
                    dr.cellsInSlots(),
                    machineGroupIndex,
                    layout,
                    palette,
                    barFont);

            Pane badgePane = new Pane();
            badgePane.setMouseTransparent(true);
            if (showPersonBadges) {
                layoutPersonBadgeOverlay(
                        badgePane,
                        dr.badgeCellsInSlots(),
                        dr.cellsInSlots(),
                        layout,
                        badgeResolver);
            }
            StackPane rowStack = new StackPane(rowCanvas, badgePane);

            String tip =
                    (dr.rowSummary() != null ? dr.rowSummary() : "")
                            + "\n（スロット "
                            + parsed.slotMinutes()
                            + " 分・設備ガント JSON と同一データ）";
            Tooltip.install(rowStack, new Tooltip(tip));

            HBox progBox = new HBox(gap);
            progBox.setMinHeight(cellBodyH);
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
                pLab.setStyle(palette.machineSideCellCss(machineGroupIndex));
                progBox.getChildren().add(pLab);
            }

            leftBodyGrid.add(pl, 2, gridR);
            rightBodyGrid.add(rowStack, 0, gridR);
            if (!progBox.getChildren().isEmpty()) {
                rightBodyGrid.add(progBox, 1, gridR);
            }
            gridR++;
        }

        ScrollPane leftBodyScroll = new ScrollPane(leftBodyGrid);
        leftBodyScroll.setHbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);
        leftBodyScroll.setVbarPolicy(ScrollPane.ScrollBarPolicy.AS_NEEDED);
        /* true のときビューポートより狭いグリッドが横に引き伸ばされ、機械名・工程名列が余白だらけになる */
        leftBodyScroll.setFitToWidth(false);
        leftBodyScroll.setMinViewportWidth(leftTotal);
        leftBodyScroll.setPrefViewportWidth(leftTotal);

        ScrollPane rightBodyScroll = new ScrollPane(rightBodyGrid);
        rightBodyScroll.setFitToWidth(false);
        rightBodyScroll.setPannable(true);
        HBox.setHgrow(rightBodyScroll, Priority.ALWAYS);

        leftHead.minWidthProperty().bind(leftBodyScroll.widthProperty());

        Region progressHeaderSpacer = new Region();
        progressHeaderSpacer.setMinWidth(progressTotal);
        progressHeaderSpacer.setPrefWidth(progressTotal);
        HBox headerRightContent = new HBox(0, headerCanvas, progressHeaderSpacer);
        headerRightContent.setMinHeight(layout.headerHeight);
        headerRightContent.setPrefHeight(layout.headerHeight);
        headerRightContent.setMaxHeight(layout.headerHeight);

        ScrollPane headerRightScroll = new ScrollPane(headerRightContent);
        headerRightScroll.setHbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);
        headerRightScroll.setVbarPolicy(ScrollPane.ScrollBarPolicy.NEVER);
        headerRightScroll.setPannable(false);
        headerRightScroll.setFitToHeight(true);
        HBox.setHgrow(headerRightScroll, Priority.ALWAYS);
        leftBodyScroll.setCache(false);
        rightBodyScroll.setCache(false);
        headerRightScroll.setCache(false);
        headerRightContent.setCache(false);

        HBox headRow = new HBox(0, leftHead, headerRightScroll);
        headRow.setMinHeight(layout.headerHeight);
        headRow.setPrefHeight(layout.headerHeight);
        headRow.setMaxHeight(layout.headerHeight);

        headerRightScroll.hvalueProperty().bindBidirectional(rightBodyScroll.hvalueProperty());
        leftBodyScroll.vvalueProperty().bindBidirectional(rightBodyScroll.vvalueProperty());

        double shiftSens =
                effectiveShiftWheelHorizontalSensitivity(shiftWheelHorizontalSensitivityPercent);
        installShiftWheelHorizontalScroll(rightBodyScroll, shiftSens);

        HBox bodySplit = new HBox(0, leftBodyScroll, rightBodyScroll);
        bodySplit.setAlignment(Pos.TOP_LEFT);
        HBox.setHgrow(rightBodyScroll, Priority.ALWAYS);
        VBox.setVgrow(bodySplit, Priority.ALWAYS);
        headRow.setCache(false);
        bodySplit.setCache(false);

        VBox mainColumn = new VBox(0, headRow, bodySplit);
        VBox.setVgrow(bodySplit, Priority.ALWAYS);
        mainColumn.setPadding(new Insets(4));
        mainColumn.setCache(false);
        root.setCenter(mainColumn);

        Label hint =
                new Label(
                        """
                        ヒント: 横スクロールで時刻軸を追えます（左3列は固定）。Shift+ホイールでも横（時刻軸方向）にスクロール（感度はツールバー「Shift横スクロール」）。Ctrl+ホイールで表示倍率。 \
                        左列は内容に応じ自動幅。日付列は見出しなし・データは反時計回り90°。同一暦日は日付列を縦結合、同一機械は機械名列を縦結合。 \
                        行の高さ・見出し行の高さ・時刻列幅・バー文字サイズはツールバーで調整できます。日付・機械名・工程名列幅もスライダーで指定できます（先頭は自動）。""");
        hint.setWrapText(true);
        hint.setStyle(palette.hintCss());
        hint.setPadding(new Insets(0, 8, 8, 8));
        root.setBottom(hint);
        return root;
    }

    /**
     * Shift+ホイール横スクロールの感度。{@code <= 0} または非有限は既定（速めにスクロール）。
     *
     * @param percentFromUi ツールバー「％」（50〜1000）。実効値は {@link Math#clamp(double, double, double)} で範囲内に収める
     */
    private static double effectiveShiftWheelHorizontalSensitivity(double percentFromUi) {
        if (!Double.isFinite(percentFromUi) || percentFromUi <= 0) {
            return 200.0;
        }
        return Math.clamp(percentFromUi, 50.0, 1000.0);
    }

    /** @param sensitivityPercent 100 で従来の {@code -delta/excess} と同等 */
    private static void installShiftWheelHorizontalScroll(
            ScrollPane scrollPane, double sensitivityPercent) {
        double factor = sensitivityPercent / 100.0;
        scrollPane.addEventFilter(
                ScrollEvent.SCROLL,
                e -> {
                    if (!e.isShiftDown()) {
                        return;
                    }
                    e.consume();
                    double deltaY = e.getDeltaY();
                    double deltaX = e.getDeltaX();
                    double delta = Math.abs(deltaY) >= Math.abs(deltaX) ? deltaY : deltaX;
                    if (delta == 0.0) {
                        return;
                    }
                    var content = scrollPane.getContent();
                    if (content == null) {
                        return;
                    }
                    var vp = scrollPane.getViewportBounds();
                    double viewportW = vp != null ? vp.getWidth() : 0.0;
                    if (!(viewportW > 0.0)) {
                        return;
                    }
                    double contentW = content.getLayoutBounds().getWidth();
                    double excess = contentW - viewportW;
                    if (!(excess > 1.0) || !Double.isFinite(excess)) {
                        return;
                    }
                    double deltaH = -delta / excess * factor;
                    scrollPane.setHvalue(Math.clamp(scrollPane.getHvalue() + deltaH, 0.0, 1.0));
                });
    }

    private static void applySideHeaderStyle(
            Label lb, double colW, LayoutMetrics layout, GanttPalette palette) {
        lb.setMinWidth(MIN_SIDE_COL_WIDTH);
        lb.setPrefWidth(colW);
        lb.setMaxWidth(colW);
        lb.setAlignment(Pos.CENTER);
        lb.setPadding(new Insets(4 * layout.zoom, 6 * layout.zoom, 4 * layout.zoom, 6 * layout.zoom));
        lb.setFont(Font.font(layout.rowLabelFontSize * 1.05));
        lb.setStyle(palette.rowLabelCss());
        lb.setMinHeight(layout.headerHeight);
        lb.setPrefHeight(layout.headerHeight);
        lb.setMaxHeight(layout.headerHeight);
    }

    private static void applySideDataStyle(
            Label lb,
            double colW,
            LayoutMetrics layout,
            GanttPalette palette,
            int machineGroupIndex) {
        lb.setMinWidth(colW);
        lb.setPrefWidth(colW);
        lb.setMaxWidth(colW);
        lb.setAlignment(Pos.CENTER_LEFT);
        lb.setPadding(new Insets(2 * layout.zoom, 6 * layout.zoom, 2 * layout.zoom, 6 * layout.zoom));
        lb.setStyle(palette.machineSideCellCss(machineGroupIndex));
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

    private static Font resolveBarFont(String family, double sizePx) {
        double s = Math.max(6, sizePx);
        if (family == null || family.isBlank()) {
            return Font.font(s);
        }
        return Font.font(family.strip(), s);
    }

    private record MachineColumnPlan(boolean continuation, String machineCellText, int rowSpan) {}

    /** 日付列: 同一暦日（正規化テキスト一致）の連続行をまとめて 1 セルに縦結合する。 */
    private record DateColumnPlan(boolean continuation, String dateText, int rowSpan) {}

    private static String[] computeCarriedDates(
            List<String> columns, List<DisplayRow> displayRows) {
        String[] out = new String[displayRows.size()];
        String cd = "";
        for (int i = 0; i < displayRows.size(); i++) {
            DisplayRow dr = displayRows.get(i);
            if (dr.sectionBanner() != null) {
                out[i] = cd;
                continue;
            }
            int dateCol = columns.indexOf("日付");
            if (dateCol >= 0 && dr.rawRow().size() > dateCol) {
                String dv =
                        dr.rawRow().get(dateCol) != null
                                ? dr.rawRow().get(dateCol).strip()
                                : "";
                if (!dv.isEmpty()) {
                    cd = dv;
                }
            }
            out[i] = cd;
        }
        return out;
    }

    private static List<MachineColumnPlan> computeMachineColumnPlans(
            List<String> columns, List<DisplayRow> displayRows) {
        List<MachineColumnPlan> plans = new ArrayList<>();
        for (int i = 0; i < displayRows.size(); i++) {
            plans.add(null);
        }
        String[] carriedAt = computeCarriedDates(columns, displayRows);
        int r = 0;
        while (r < displayRows.size()) {
            DisplayRow dr = displayRows.get(r);
            if (dr.sectionBanner() != null) {
                r++;
                continue;
            }
            String rawMach = cellAt(columns, dr.rawRow(), "機械名");
            if (rawMach.isEmpty()) {
                LeftParts lp = buildLeftParts(columns, dr.rawRow(), carriedAt[r]);
                plans.set(r, new MachineColumnPlan(false, lp.machine(), 1));
                r++;
                continue;
            }
            String mergeKey = machineMergeKey(rawMach);
            int j = r + 1;
            while (j < displayRows.size()) {
                DisplayRow drj = displayRows.get(j);
                if (drj.sectionBanner() != null) {
                    break;
                }
                String rawJ = cellAt(columns, drj.rawRow(), "機械名");
                if (rawJ.isEmpty()) {
                    break;
                }
                if (!mergeKey.equals(machineMergeKey(rawJ))) {
                    break;
                }
                j++;
            }
            plans.set(r, new MachineColumnPlan(false, mergeKey, j - r));
            for (int k = r + 1; k < j; k++) {
                plans.set(k, new MachineColumnPlan(true, "", 1));
            }
            r = j;
        }
        return plans;
    }

    private static List<DateColumnPlan> computeDateColumnPlans(
            List<String> columns, List<DisplayRow> displayRows) {
        List<DateColumnPlan> plans = new ArrayList<>();
        for (int i = 0; i < displayRows.size(); i++) {
            plans.add(null);
        }
        String[] carriedAt = computeCarriedDates(columns, displayRows);
        int r = 0;
        while (r < displayRows.size()) {
            DisplayRow dr = displayRows.get(r);
            if (dr.sectionBanner() != null) {
                r++;
                continue;
            }
            String dateKey = compactDateLine(carriedAt[r]).strip();
            if (dateKey.isEmpty()) {
                plans.set(r, new DateColumnPlan(false, "", 1));
                r++;
                continue;
            }
            int j = r + 1;
            while (j < displayRows.size()) {
                DisplayRow drj = displayRows.get(j);
                if (drj.sectionBanner() != null) {
                    break;
                }
                String keyJ = compactDateLine(carriedAt[j]).strip();
                if (!dateKey.equals(keyJ)) {
                    break;
                }
                j++;
            }
            plans.set(r, new DateColumnPlan(false, dateKey, j - r));
            for (int k = r + 1; k < j; k++) {
                plans.set(k, new DateColumnPlan(true, "", 1));
            }
            r = j;
        }
        return plans;
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
            return fromScales(zoomIn, 100, 100, 100, 100);
        }

        /**
         * @param zoomIn 表示倍率（0.5〜2.0、スライダー 100%＝1.0）
         * @param rowHeightPercent 行の高さ 50〜200（100＝基準、0／負は 100 扱い）
         * @param slotWidthPercent 時刻 1 スロット列幅 50〜500
         * @param barFontPercent バー内文字 50〜200（100＝基準）
         * @param headerHeightPercent 見出し行の高さ 50〜200（100＝基準）
         */
        static LayoutMetrics fromScales(
                double zoomIn,
                double rowHeightPercent,
                double slotWidthPercent,
                double barFontPercent,
                double headerHeightPercent) {
            double z = Math.clamp(zoomIn, 0.5, 2.0);
            double rPct =
                    rowHeightPercent <= 0 || rowHeightPercent > 500
                            ? 100
                            : Math.clamp(rowHeightPercent, 50, 200);
            double sPct =
                    slotWidthPercent <= 0 || slotWidthPercent > 500
                            ? 100
                            : Math.clamp(slotWidthPercent, 50, 500);
            double bfPct =
                    barFontPercent <= 0 || barFontPercent > 500
                            ? 100
                            : Math.clamp(barFontPercent, 50, 200);
            double hdrPct =
                    headerHeightPercent <= 0 || headerHeightPercent > 500
                            ? 100
                            : Math.clamp(headerHeightPercent, 50, 200);
            double rScale = rPct / 100.0;
            double sScale = sPct / 100.0;
            double bfScale = bfPct / 100.0;
            double hdrScale = hdrPct / 100.0;
            double barPx = Math.max(8, 9 * z * bfScale);
            return new LayoutMetrics(
                    z,
                    BASE_SLOT_WIDTH * z * sScale,
                    BASE_ROW_HEIGHT * z * rScale,
                    BASE_SECTION_ROW_HEIGHT * z * rScale,
                    BASE_HEADER_HEIGHT * z * hdrScale,
                    BASE_LABEL_MIN_WIDTH * z,
                    BASE_LABEL_MAX_WIDTH * z,
                    11 * z,
                    10 * z,
                    barPx,
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
            Color[] machineBands,
            String machineSideTextFill,
            String machineSideBorder,
            String rowLabelCss,
            String sectionBannerCss,
            String hintCss,
            String progressEmptyCss,
            String progressFilledCss) {

        Color machineBandFill(int machineGroupIndex) {
            Color[] b = machineBands();
            return b[Math.floorMod(machineGroupIndex, b.length)];
        }

        /** 機械名・工程名・進捗列の Excel 風パステル背景（機械ブロック単位で同色） */
        String machineSideCellCss(int machineGroupIndex) {
            return "-fx-background-color: "
                    + rgbHex(machineBandFill(machineGroupIndex))
                    + "; -fx-text-fill: "
                    + machineSideTextFill()
                    + "; -fx-border-color: "
                    + machineSideBorder()
                    + "; -fx-border-width: 0 1 1 0;";
        }

        static GanttPalette forTheme(DesktopTheme theme) {
            boolean dark = theme.isDarkUi();
            if (dark) {
                Color[] bandsDark =
                        new Color[] {
                            Color.web("#223449"),
                            Color.web("#3d2836"),
                            Color.web("#243d30"),
                            Color.web("#454018")
                        };
                return new GanttPalette(
                        Color.web("#1e293b"),
                        Color.web("#0f172a"),
                        Color.web("#64748b"),
                        Color.web("#334155"),
                        Color.web("#e2e8f0"),
                        Color.web("#60a5fa"),
                        Color.web("#93c5fd"),
                        Color.web("#fdba74"),
                        Color.web("#f8fafc"),
                        /* バー外ラベルは機械行の暗い帯の上に描くため明るい色にする */
                        Color.web("#f1f5f9"),
                        Color.web("#fffbeb"),
                        Color.web("#93c5fd"),
                        bandsDark,
                        "#e2e8f0",
                        "#64748b",
                        "-fx-background-color: #334155; -fx-text-fill: #f8fafc; "
                                + "-fx-border-color: #64748b; -fx-border-width: 0 1 0 0;",
                        "-fx-background-color: #1e293b; -fx-font-weight: bold; -fx-font-size: 12px; "
                                + "-fx-text-fill: #f1f5f9;",
                        "-fx-text-fill: #94a3b8; -fx-font-size: 11px;",
                        "-fx-background-color: #1e293b; -fx-border-color: #475569; -fx-text-fill: #cbd5e1;",
                        "-fx-background-color: #422006; -fx-border-color: #d97706; -fx-text-fill: #ffedd5;");
            }
            Color[] bandsLight =
                    new Color[] {
                        Color.web("#ddebf7"),
                        Color.web("#fce4d6"),
                        Color.web("#e2efda"),
                        Color.web("#fff2cc")
                    };
            return new GanttPalette(
                    Color.web("#ffffff"),
                    Color.web("#f3f3f3"),
                    Color.web("#d9d9d9"),
                    Color.web("#f8f9fa"),
                    Color.web("#333333"),
                    Color.web("#8faadc"),
                    Color.web("#c9daf8"),
                    Color.web("#ffd966"),
                    Color.web("#111827"),
                    Color.web("#1e3a5f"),
                    Color.web("#854d0e"),
                    Color.web("#2f5597"),
                    bandsLight,
                    "#111827",
                    "#bfbfbf",
                    "-fx-background-color: #4472c4; -fx-text-fill: #ffffff; -fx-font-weight: bold; "
                            + "-fx-border-color: #bfbfbf; -fx-border-width: 0 1 1 0;",
                    "-fx-background-color: #4472c4; -fx-font-weight: bold; -fx-font-size: 12px; "
                            + "-fx-text-fill: #ffffff;",
                    "-fx-text-fill: #595959; -fx-font-size: 11px;",
                    "-fx-background-color: #fffbf0; -fx-border-color: #d9d9d9; -fx-text-fill: #334155;",
                    "-fx-background-color: #fff2cc; -fx-border-color: #bfbfbf; -fx-text-fill: #334155;");
        }
    }

    private static String rgbHex(Color c) {
        int r = (int) Math.round(c.getRed() * 255);
        int g = (int) Math.round(c.getGreen() * 255);
        int b = (int) Math.round(c.getBlue() * 255);
        return String.format("#%02x%02x%02x", r, g, b);
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
        gc.setLineWidth(0.35);

        List<Integer> slotCols = parsed.slotColumnIndices();
        int n = slotCols.size();
        LocalTime t0 = parsed.slotBaseTime();
        int slotMin = Math.max(1, parsed.slotMinutes());

        for (int i = 0; i <= n; i++) {
            double x = i * layout.slotWidth;
            gc.strokeLine(x, 0, x, layout.headerHeight);
        }
        gc.strokeRect(0, 0, timelineWidth, layout.headerHeight);

        int labelStep = timeLabelStep(n, slotMin, layout.slotWidth);
        double labelFont = Math.min(layout.axisFontSize, Math.max(7, layout.slotWidth * 0.95));
        gc.setFill(palette.axisLabel());
        gc.setFont(Font.font(labelFont));
        DateTimeFormatter tf = DateTimeFormatter.ofPattern("H:mm");
        for (int i = 0; i < n; i += labelStep) {
            double cx = i * layout.slotWidth + layout.slotWidth * 0.5;
            LocalTime tt = t0.plusMinutes((long) i * slotMin);
            String txt = tt.format(tf);
            gc.save();
            double cy = layout.headerHeight * 0.38;
            gc.translate(cx, cy);
            gc.rotate(-90);
            double tw = approxLatinDigitTextWidth(txt, labelFont);
            gc.setTextBaseline(VPos.CENTER);
            gc.fillText(txt, -tw / 2, 0);
            gc.restore();
        }
    }

    /** スロットが細いときは時間単位で間引き、それ以外は Excel に近い密度でラベル */
    private static int timeLabelStep(int slotCount, int slotMinutes, double slotWidthPx) {
        if (slotCount <= 0) {
            return 1;
        }
        int perHour = Math.max(1, 60 / slotMinutes);
        double totalPx = slotCount * slotWidthPx;
        if (totalPx <= 960 || slotWidthPx >= 14) {
            return 1;
        }
        if (totalPx <= 1600) {
            return Math.min(perHour, 3);
        }
        return perHour;
    }

    private static double approxLatinDigitTextWidth(String txt, double fontPx) {
        String s = txt != null ? txt : "";
        return Math.max(8, s.length() * fontPx * 0.52);
    }

    private record BarRun(int fromSlot, int toSlot, String text, BarKind kind) {}

    private static void layoutPersonBadgeOverlay(
            Pane overlay,
            List<String> badgeSlotTexts,
            List<String> slotTexts,
            LayoutMetrics layout,
            Function<String, PersonBadgeStyle> styleForLabel) {
        if (overlay == null
                || styleForLabel == null
                || badgeSlotTexts == null
                || slotTexts == null
                || badgeSlotTexts.size() != slotTexts.size()) {
            return;
        }
        List<BarRun> runs = collectBarRuns(slotTexts);
        for (BarRun run : runs) {
            if (run.kind() == BarKind.BREAK) {
                continue;
            }
            String frag =
                    PersonNameBadgeText.firstNonEmptyInSlotRange(
                            badgeSlotTexts, run.fromSlot(), run.toSlot());
            if (frag.isEmpty()) {
                continue;
            }
            List<String> parts = PersonNameBadgeText.splitBadgeCell(frag);
            if (parts.isEmpty()) {
                continue;
            }
            HBox box = new HBox(2 * layout.zoom);
            box.setAlignment(Pos.CENTER);
            for (String p : parts) {
                box.getChildren()
                        .add(
                                PersonBadgeNodeFactory.createBadge(
                                        p,
                                        styleForLabel.apply(p),
                                        layout.zoom,
                                        layout.rowLabelFontSize));
            }
            double barTop = 3 * layout.zoom;
            double barH = layout.rowHeight - 2 * barTop;
            box.applyCss();
            box.layout();
            double w = Math.max(1, box.prefWidth(-1));
            double h = Math.max(1, box.prefHeight(-1));
            double cx = (run.fromSlot() + run.toSlot() + 1) * 0.5 * layout.slotWidth;
            double x0 = cx - w / 2;
            double y0 = barTop + (barH - h) / 2;
            box.setLayoutX(x0);
            box.setLayoutY(y0);
            overlay.getChildren().add(box);
        }
    }

    private static void drawTimelineRow(
            GraphicsContext gc,
            List<String> slotTexts,
            int machineGroupIndex,
            LayoutMetrics layout,
            GanttPalette palette,
            Font barFont) {
        int n = slotTexts.size();
        Color band = palette.machineBandFill(machineGroupIndex);
        for (int i = 0; i < n; i++) {
            double x = i * layout.slotWidth;
            gc.setFill(band);
            gc.fillRect(x, 0, layout.slotWidth, layout.rowHeight);
        }
        gc.setStroke(palette.grid());
        gc.setLineWidth(0.35);
        for (int i = 0; i <= n; i++) {
            gc.strokeLine(i * layout.slotWidth, 0, i * layout.slotWidth, layout.rowHeight);
        }

        List<BarRun> runs = collectBarRuns(slotTexts);
        for (BarRun run : runs) {
            fillBar(gc, run, layout, palette);
        }
        drawBarLabelsOutside(gc, runs, layout, palette, barFont);
    }

    private static List<BarRun> collectBarRuns(List<String> slotTexts) {
        int n = slotTexts.size();
        List<BarRun> runs = new ArrayList<>();
        int runStart = -1;
        String runText = "";
        for (int i = 0; i < n; i++) {
            String t = slotTexts.get(i) != null ? slotTexts.get(i).strip() : "";
            boolean empty = t.isEmpty();
            if (empty) {
                if (runStart >= 0) {
                    runs.add(
                            new BarRun(
                                    runStart,
                                    i - 1,
                                    runText,
                                    classifyBar(runText)));
                    runStart = -1;
                    runText = "";
                }
                continue;
            }
            if (runStart < 0) {
                runStart = i;
                runText = t;
            } else if (!t.equals(runText)) {
                runs.add(
                        new BarRun(
                                runStart,
                                i - 1,
                                runText,
                                classifyBar(runText)));
                runStart = i;
                runText = t;
            }
        }
        if (runStart >= 0) {
            runs.add(
                    new BarRun(
                            runStart,
                            n - 1,
                            runText,
                            classifyBar(runText)));
        }
        return runs;
    }

    private static void fillBar(
            GraphicsContext gc, BarRun run, LayoutMetrics layout, GanttPalette palette) {
        int fromSlot = run.fromSlot();
        int toSlot = run.toSlot();
        double x = fromSlot * layout.slotWidth;
        double w = (toSlot - fromSlot + 1) * layout.slotWidth;
        BarKind kind = run.kind();
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
    }

    private static void drawBarLabelsOutside(
            GraphicsContext gc,
            List<BarRun> runs,
            LayoutMetrics layout,
            GanttPalette palette,
            Font barFont) {
        if (runs.isEmpty()) {
            return;
        }
        List<BarRun> sorted = new ArrayList<>(runs);
        sorted.sort(Comparator.comparingInt(BarRun::fromSlot));
        double inset = 0.5 * layout.zoom;
        double barTop = 3 * layout.zoom;
        double barH = layout.rowHeight - 2 * barTop;
        double pad = 6 * layout.zoom;

        List<double[]> aboveRanges = new ArrayList<>();
        List<double[]> belowRanges = new ArrayList<>();

        gc.setFont(barFont);

        for (BarRun run : sorted) {
            String full = run.text().replace('\n', ' ');
            if (full.length() > 220) {
                full = full.substring(0, 217) + "...";
            }
            if (full.isEmpty()) {
                continue;
            }
            Color labelColor =
                    switch (run.kind()) {
                        case BREAK -> palette.barBreakText();
                        case STARTUP -> palette.barStartupText();
                        default -> palette.barDefaultText();
                    };

            double lx = run.fromSlot() * layout.slotWidth + inset + 3 * layout.zoom;
            double tw = measureTextWidth(full, barFont);
            double fh = measureTextHeight(full, barFont);
            double rx = lx + tw;

            boolean useAbove;
            if (!horizontalHits(aboveRanges, lx, rx, pad)) {
                useAbove = true;
            } else if (!horizontalHits(belowRanges, lx, rx, pad)) {
                useAbove = false;
            } else {
                useAbove = (run.fromSlot() & 1) == 0;
            }
            if (useAbove) {
                aboveRanges.add(new double[] {lx, rx});
            } else {
                belowRanges.add(new double[] {lx, rx});
            }

            double baseline =
                    useAbove
                            ? barTop - fh * 0.35
                            : barTop + barH + fh * 0.75;

            gc.setFill(labelColor);
            gc.fillText(full, lx, baseline);
        }
    }

    private static boolean horizontalHits(
            List<double[]> ranges, double lo, double hi, double pad) {
        for (double[] r : ranges) {
            if (!(hi + pad < r[0] || lo - pad > r[1])) {
                return true;
            }
        }
        return false;
    }

    private static double measureTextWidth(String s, Font f) {
        Text t = new Text(s != null ? s : "");
        t.setFont(f);
        return t.getLayoutBounds().getWidth();
    }

    private static double measureTextHeight(String s, Font f) {
        Text t = new Text(s != null ? s : "");
        t.setFont(f);
        return t.getLayoutBounds().getHeight();
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

    private static ParseResult parse(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            List<List<String>> badgeSlotRowsRaw) {
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
                displayRows.add(new DisplayRow(banner, null, null, null, null, null, null, row));
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
            String dateForCol = compactDateLine(carriedDate);
            List<String> slotCells = new ArrayList<>();
            for (int si : slots) {
                String v =
                        row.size() > si && row.get(si) != null ? row.get(si) : "";
                slotCells.add(v);
            }
            List<String> badgeSlotCells = new ArrayList<>();
            if (badgeSlotRowsRaw != null && r < badgeSlotRowsRaw.size()) {
                List<String> br = badgeSlotRowsRaw.get(r);
                for (int j = 0; j < slots.size(); j++) {
                    String b = (br != null && j < br.size() && br.get(j) != null) ? br.get(j) : "";
                    badgeSlotCells.add(b);
                }
            } else {
                for (int j = 0; j < slots.size(); j++) {
                    badgeSlotCells.add("");
                }
            }
            displayRows.add(
                    new DisplayRow(
                            null,
                            lp.machine(),
                            lp.process(),
                            lp.summary(),
                            dateForCol,
                            slotCells,
                            badgeSlotCells,
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
     * 機械名列・工程名列に分割する。日付は左端「日付」列で表示するため機械名からは付けない。
     * ツールチップ用の要約文字列も返す。
     */
    private static LeftParts buildLeftParts(
            List<String> columns, ObservableList<String> row, String carriedDate) {
        String machRaw = cellAt(columns, row, "機械名");
        String mach = machRaw.isEmpty() ? "" : machineMergeKey(machRaw);
        String proc = cellAt(columns, row, "工程名");
        String task = cellAt(columns, row, "タスク概覝");
        String tb = cellAt(columns, row, "日時帯");
        String dateLine = compactDateLine(carriedDate);

        StringBuilder machCol = new StringBuilder();
        if (!mach.isEmpty()) {
            machCol.append(mach);
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
        if (!dateLine.isEmpty()) {
            summary =
                    summary.isEmpty()
                            ? dateLine
                            : dateLine + "\n" + summary;
        }
        if (summary.isEmpty()) {
            return new LeftParts("", "", "（行）");
        }
        return new LeftParts(machStr, procStr, summary);
    }

    /**
     * 機械名列末尾の {@code (...)} および全角 {@code （…）}（工程名の重複表示など）を除いた表示・縦結合用キー。
     */
    private static String machineMergeKey(String rawMachineCell) {
        if (rawMachineCell == null || rawMachineCell.isBlank()) {
            return "";
        }
        String s = rawMachineCell.strip().split("\\R", 2)[0].strip();
        String prev;
        do {
            prev = s;
            s = s.replaceFirst("\\s*\\([^)]*\\)\\s*$", "").strip();
            s = s.replaceFirst("\\s*（[^）]*）\\s*$", "").strip();
        } while (!s.equals(prev));
        return s;
    }

    /**
     * 「【2026/05/07】」等から年月日を取り、{@code 2026年4月18日(土)} 形式に正規化する。
     */
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
            try {
                int y = Integer.parseInt(m.group(1));
                int mo = Integer.parseInt(m.group(2));
                int d = Integer.parseInt(m.group(3));
                LocalDate date = LocalDate.of(y, mo, d);
                String narrow =
                        date.getDayOfWeek().getDisplayName(TextStyle.NARROW, Locale.JAPAN);
                return String.format(
                        "%d年%d月%d日(%s)",
                        date.getYear(), date.getMonthValue(), date.getDayOfMonth(), narrow);
            } catch (Exception ignored) {
                return s;
            }
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
     * @param machineLine データ行: 機械名列のテキスト（日付は含めない）
     * @param processBlock データ行: 工程名列（工程名・タスク等）
     * @param rowSummary データ行: ツールチップ用
     * @param dateCompact 左「日付」列用（例 2026/05/07、空可）
     */
    private record DisplayRow(
            String sectionBanner,
            String machineLine,
            String processBlock,
            String rowSummary,
            String dateCompact,
            List<String> cellsInSlots,
            List<String> badgeCellsInSlots,
            ObservableList<String> rawRow) {}
}
