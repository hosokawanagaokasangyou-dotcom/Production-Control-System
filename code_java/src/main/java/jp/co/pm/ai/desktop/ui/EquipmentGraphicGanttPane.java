package jp.co.pm.ai.desktop.ui;

import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.Locale;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.function.BiConsumer;

import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Group;
import javafx.scene.Scene;
import javafx.scene.Cursor;
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
import javafx.scene.input.MouseButton;
import javafx.scene.input.MouseEvent;
import javafx.scene.input.ScrollEvent;

import javafx.scene.layout.Pane;
import javafx.scene.layout.Region;
import javafx.scene.layout.StackPane;
import javafx.scene.shape.Line;

import jp.co.pm.ai.desktop.config.DesktopSessionState;
import jp.co.pm.ai.desktop.config.DesktopTheme;
import jp.co.pm.ai.desktop.config.EquipmentGanttBadgeDragDelta;
import java.util.function.Function;

import javafx.animation.PauseTransition;
import javafx.application.Platform;
import javafx.geometry.BoundingBox;
import javafx.geometry.Bounds;
import javafx.geometry.Point2D;
import javafx.scene.Node;
import javafx.util.Duration;

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

    /**
     * 1 行あたりのタイムライン {@link Canvas} 幅（px）の上限。超過時はスロット幅を自動縮小し、GPU／ヒープ負荷を抑える。
     */
    private static final double MAX_TIMELINE_CANVAS_WIDTH_PX = 3072.0;

    /**
     * 行 Canvas 合計の RGBA ナイーブ見積（MiB）がこの値を超えると、面積比の平方根でスロット幅を追加縮小する。
     */
    private static final double MAX_NAIVE_ROW_CANVAS_RGBA_TOTAL_MIB = 512.0;

    /** 担当バッジの横方向固定間隔（px）。実体は {@link DesktopSessionState#DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_GAP_PX}。 */
    public static final double DEFAULT_PERSON_BADGE_GAP_PX =
            DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_GAP_PX;

    /** バッジワイヤーの不透明度（{@link #fillBar} のストローク色に乗算）。 */
    private static final double PERSON_BADGE_WIRE_OPACITY = 0.45;

    private static final Group MEASURE_ROOT = new Group();
    private static final Scene MEASURE_SCENE = new Scene(MEASURE_ROOT, 4000, 4000);

    private EquipmentGraphicGanttPane() {}

    /** {@link #build} のルート {@link BorderPane} から再構築前に保存するスクロール位置。 */
    public record EquipmentGanttScrollState(double hValue, double vValue) {}

    /** 時刻軸（右ペイン）の {@link ScrollPane}。縦スクロールは左ペインと双方向バインド済み。 */
    public record EquipmentGanttViewHandles(ScrollPane timelineScroll, Runnable scheduleViewportRepaint) {
        public EquipmentGanttViewHandles(ScrollPane timelineScroll) {
            this(timelineScroll, null);
        }
    }

    /** 横スクロールの見えている範囲の前後に確保するスロット数（部分描画のマージン）。 */
    private static final int VIEWPORT_SLOT_MARGIN = 48;

    private static final String PROFILE_PROP = "pm.ai.gantt.profile";

    /**
     * ビューポートとスロット幅から、背景・格子を描くスロットインデックス範囲（両端含む）を返す。
     */
    public static int[] visibleSlotRangeInclusive(
            ScrollPane sp, double slotWidthPx, int slotCount, int marginSlots) {
        if (slotCount <= 0 || slotWidthPx <= 1e-9) {
            return new int[] {0, Math.max(0, slotCount - 1)};
        }
        if (sp == null) {
            return new int[] {0, Math.max(0, slotCount - 1)};
        }
        Node content = sp.getContent();
        javafx.geometry.Bounds vp = sp.getViewportBounds();
        if (content == null || vp == null || vp.getWidth() <= 1.0) {
            return new int[] {0, slotCount - 1};
        }
        double contentW = content.getLayoutBounds().getWidth();
        double viewportW = vp.getWidth();
        double excess = contentW - viewportW;
        double scrollPx = excess > 1e-6 ? sp.getHvalue() * excess : 0.0;
        int from = (int) Math.floor(scrollPx / slotWidthPx) - marginSlots;
        int to = (int) Math.ceil((scrollPx + viewportW) / slotWidthPx) + marginSlots;
        from = Math.max(0, from);
        to = Math.min(slotCount - 1, to);
        if (from > to) {
            return new int[] {0, slotCount - 1};
        }
        return new int[] {from, to};
    }

    /**
     * Ctrl+ホイールで拡大率を変えるとき、マウス下の内容位置を維持するためのアンカー（コンテンツ座標の X とビューポート内 X）。
     */
    public record HorizontalZoomAnchor(double anchorContentX, double viewportMouseX) {}

    /**
     * 再構築前のガント {@link BorderPane} から横・縦スクロール位置を取得する。
     *
     * @param graphicRoot {@link #build} の戻り値。無効時は (0,0)
     */
    public static EquipmentGanttScrollState snapshotScroll(BorderPane graphicRoot) {
        if (graphicRoot == null) {
            return new EquipmentGanttScrollState(0.0, 0.0);
        }
        Object ud = graphicRoot.getUserData();
        if (!(ud instanceof EquipmentGanttViewHandles handles)) {
            return new EquipmentGanttScrollState(0.0, 0.0);
        }
        ScrollPane sp = handles.timelineScroll();
        return new EquipmentGanttScrollState(sp.getHvalue(), sp.getVvalue());
    }

    /**
     * 再構築直後のガントにスクロールを戻す。レイアウト確定後に合わせるため {@link Platform#runLater} を使う。
     *
     * @param zoomAnchorOrNull 非 null のときは横位置のみアンカー基準（マウス中心ズーム）。縦は常に {@code snap} を使う。
     */
    public static void restoreScrollAfterRebuild(
            BorderPane graphicRoot,
            EquipmentGanttScrollState snap,
            HorizontalZoomAnchor zoomAnchorOrNull) {
        if (graphicRoot == null || snap == null) {
            return;
        }
        Platform.runLater(
                () -> {
                    Object ud = graphicRoot.getUserData();
                    if (!(ud instanceof EquipmentGanttViewHandles handles)) {
                        return;
                    }
                    ScrollPane sp = handles.timelineScroll();
                    sp.setVvalue(Math.clamp(snap.vValue(), 0.0, 1.0));
                    if (zoomAnchorOrNull != null) {
                        Node content = sp.getContent();
                        Bounds vp = sp.getViewportBounds();
                        if (content == null || vp == null) {
                            sp.setHvalue(Math.clamp(snap.hValue(), 0.0, 1.0));
                            return;
                        }
                        double contentW = content.getLayoutBounds().getWidth();
                        double viewportW = vp.getWidth();
                        double excess = contentW - viewportW;
                        if (!(excess > 1.0) || !Double.isFinite(excess)) {
                            sp.setHvalue(0.0);
                            return;
                        }
                        double scrollNew =
                                zoomAnchorOrNull.anchorContentX() - zoomAnchorOrNull.viewportMouseX();
                        sp.setHvalue(Math.clamp(scrollNew / excess, 0.0, 1.0));
                    } else {
                        sp.setHvalue(Math.clamp(snap.hValue(), 0.0, 1.0));
                    }
                });
    }

    /**
     * マウス位置を基準にズームしたあと横スクロールを合わせるため、アンカーを求める。
     *
     * @return 取得できないときは null
     */
    public static HorizontalZoomAnchor computeHorizontalZoomAnchor(ScrollPane timelineScroll, ScrollEvent e) {
        if (timelineScroll == null || e == null) {
            return null;
        }
        Bounds vp = timelineScroll.getViewportBounds();
        if (vp == null || !(vp.getWidth() > 0.0)) {
            return null;
        }
        Point2D local = timelineScroll.sceneToLocal(e.getSceneX(), e.getSceneY());
        double mouseInVpX = local.getX() - vp.getMinX();
        Node content = timelineScroll.getContent();
        if (content == null) {
            return null;
        }
        double contentW = content.getLayoutBounds().getWidth();
        double viewportW = vp.getWidth();
        double excess = Math.max(0.0, contentW - viewportW);
        double scrollPx = timelineScroll.getHvalue() * excess;
        double anchorContentX = scrollPx + mouseInVpX;
        return new HorizontalZoomAnchor(anchorContentX, mouseInVpX);
    }

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
                null,
                DEFAULT_PERSON_BADGE_GAP_PX,
                DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX,
                false,
                null,
                null,
                DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_ENABLED);
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
                null,
                DEFAULT_PERSON_BADGE_GAP_PX,
                DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX,
                false,
                null,
                null,
                DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_WIRE_ENABLED);
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
     * @param personBadgeGapPx 担当バッジ横方向の固定間隔（px、隣接ピル左端間の追加距離、{@code <0} は既定）
     * @param personBadgeBandVerticalOffsetPx バッジブロックをタスク帯に対して縦にずらす量（px、正で下方向、非有限は既定）
     * @param personBadgeDragAdjustEnabled バッジをドラッグで移動する
     * @param personBadgeDragDeltas {@link #computeDataFingerprint} が同一のとき適用するドラッグずれ（{@code null} で空）
     * @param personBadgeDragDeltaSink ドラッグ確定時にずれを通知（{@code null} で保存しない）
     * @param showPersonBadgeWires 担当バッジとチャートバーをワイヤーで結ぶ（{@code showPersonBadges} が false のとき無効）
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
            Function<String, PersonBadgeStyle> personBadgeStyleResolver,
            double personBadgeGapPx,
            double personBadgeBandVerticalOffsetPx,
            boolean personBadgeDragAdjustEnabled,
            Map<String, EquipmentGanttBadgeDragDelta> personBadgeDragDeltas,
            BiConsumer<String, EquipmentGanttBadgeDragDelta> personBadgeDragDeltaSink,
            boolean showPersonBadgeWires) {
        BorderPane root = new BorderPane();
        root.setCache(false);
        RepairedGanttTable repairedTable = RepairedGanttTable.from(columns, rows, badgeSlotRowsRaw);
        List<String> effCols = repairedTable.effCols();
        ObservableList<ObservableList<String>> effRows = repairedTable.effRows();
        List<List<String>> badgeEff = repairedTable.badgeEff();
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
        Color personBadgeWireStroke =
                Color.color(
                        palette.barStroke().getRed(),
                        palette.barStroke().getGreen(),
                        palette.barStroke().getBlue(),
                        PERSON_BADGE_WIRE_OPACITY);
        boolean personBadgeWiresEffective = showPersonBadges && showPersonBadgeWires;
        Font barFont = resolveBarFont(barFontFamily, layout.barFontSize);

        double gapPxEff = personBadgeGapPx;
        if (!Double.isFinite(gapPxEff) || gapPxEff < 0d) {
            gapPxEff = DEFAULT_PERSON_BADGE_GAP_PX;
        }
        gapPxEff =
                Math.clamp(
                        gapPxEff,
                        0d,
                        DesktopSessionState.MAX_EQUIPMENT_GANTT_PERSON_BADGE_GAP_PX);

        double bandVertEff = personBadgeBandVerticalOffsetPx;
        if (!Double.isFinite(bandVertEff)) {
            bandVertEff =
                    DesktopSessionState.DEFAULT_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX;
        }
        bandVertEff =
                Math.clamp(
                        bandVertEff,
                        DesktopSessionState.MIN_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX,
                        DesktopSessionState.MAX_EQUIPMENT_GANTT_PERSON_BADGE_BAND_VERTICAL_OFFSET_PX);

        Map<String, EquipmentGanttBadgeDragDelta> dragEff =
                personBadgeDragDeltas != null ? personBadgeDragDeltas : Map.of();

        double timelineOuterPad =
                Math.min(
                        layout.rowHeight * 0.32,
                        Math.max(5 * layout.zoom, barFont.getSize() * 0.9));
        double cellBodyH = layout.rowHeight + 2 * timelineOuterPad;

        final int slotColCount = parsed.slotColumnIndices().size();
        int approxTimelineRows = countNonSectionDisplayRows(parsed.displayRows());
        double timelineWidthBeforeCap = slotColCount * layout.slotWidth;
        double naiveTotalRgbaMiB =
                (timelineWidthBeforeCap * cellBodyH * approxTimelineRows * 4.0)
                        / (1024.0 * 1024.0);
        double slotWidthScale = 1.0;
        if (timelineWidthBeforeCap > MAX_TIMELINE_CANVAS_WIDTH_PX) {
            slotWidthScale =
                    Math.min(
                            slotWidthScale,
                            MAX_TIMELINE_CANVAS_WIDTH_PX / timelineWidthBeforeCap);
        }
        if (naiveTotalRgbaMiB > MAX_NAIVE_ROW_CANVAS_RGBA_TOTAL_MIB && naiveTotalRgbaMiB > 0) {
            slotWidthScale =
                    Math.min(
                            slotWidthScale,
                            Math.sqrt(MAX_NAIVE_ROW_CANVAS_RGBA_TOTAL_MIB / naiveTotalRgbaMiB));
        }
        if (slotWidthScale < 1.0 - 1e-12) {
            layout = layout.scaleSlotWidth(slotWidthScale);
        }

        List<MachineColumnPlan> machPlans =
                computeMachineColumnPlans(effCols, parsed.displayRows());
        List<DateColumnPlan> datePlans = computeDateColumnPlans(effCols, parsed.displayRows());
        MeasuredLeftWidths auto =
                measureAutoLeftWidths(effCols, parsed, machPlans, layout);
        double dateW = effectiveLeftColWidth(auto.dateW(), dateColWidthOverridePx);
        double machW = effectiveLeftColWidth(auto.machW(), machineColWidthOverridePx);
        double procW = effectiveLeftColWidth(auto.procW(), processColWidthOverridePx);
        double leftTotal = dateW + machW + procW;

        double timelineWidth = slotColCount * layout.slotWidth;
        /*
         * 幅・高さが 0 に近い／Prism の Canvas バックバッファ生成失敗時に NGCanvas で NPE になるのを避ける。
         * GPU ドライバ由来の不具合が残る場合は JVM に -Dprism.order=sw を試す。
         */
        final double canvasTimelineW = Math.max(1.0, timelineWidth);
        final double canvasHeaderH = Math.max(1.0, layout.headerHeight);

        Canvas headerCanvas = new Canvas(canvasTimelineW, canvasHeaderH);
        headerCanvas.setCache(false);

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

        int machineColorSeq = -1;
        int gridR = 0;
        int timelineCanvasRowCount = 0;
        List<ViewportRowSpec> viewportRowSpecs = new ArrayList<>();
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
            timelineCanvasRowCount++;
            rowCanvas.setCache(false);
            viewportRowSpecs.add(
                    new ViewportRowSpec(
                            rowCanvas,
                            timelineOuterPad,
                            dr.cellsInSlots(),
                            machineGroupIndex,
                            rowCanvasH));

            Pane badgePane = new Pane();
            /*
             * 親 Pane が mouseTransparent のとき JavaFX は子もヒットしない（バッジが最前面でもドラッグ不可）。
             * ドラッグ調整 ON のときだけ非透過にし、余白は pickOnBounds(false) で Canvas 側へ透過させる。
             */
            badgePane.setPickOnBounds(false);
            badgePane.setMouseTransparent(!(showPersonBadges && personBadgeDragAdjustEnabled));
            if (showPersonBadges) {
                layoutPersonBadgeOverlay(
                        badgePane,
                        dr.badgeCellsInSlots(),
                        dr.cellsInSlots(),
                        layout,
                        badgeResolver,
                        ri,
                        gapPxEff,
                        personBadgeDragAdjustEnabled,
                        timelineOuterPad,
                        canvasTimelineW,
                        rowCanvasH,
                        dragEff,
                        personBadgeDragDeltaSink,
                        personBadgeWiresEffective,
                        personBadgeWireStroke);
                /*
                 * 帯内クランプ（clampBadgeLayoutYInBand）がスタック高≧帯のときオフセットを打ち消すため、
                 * レイアウト後にオーバーレイPaneへ縦移動だけ適用する。
                 */
                badgePane.setTranslateY(bandVertEff);
            }
            StackPane rowStack = new StackPane(rowCanvas, badgePane);
            /*
             * Canvas と Pane を並べるだけでも後勝ちだが、子追加や再有効化時に順序がずれたとき備えて明示的に最前面へ。
             */
            if (showPersonBadges && personBadgeDragAdjustEnabled) {
                badgePane.toFront();
            }

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
        /*
         * 既定の setPannable(true) は主ボタンドラッグでパンするため、担当バッジの左ドラッグと競合する。
         * ホイール（中）ボタンのみドラッグパンする。
         */
        rightBodyScroll.setPannable(false);
        installMiddleButtonPanScroll(rightBodyScroll);
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

        final LayoutMetrics layoutViewport = layout;

        PauseTransition viewportRepaintDebounce = new PauseTransition(Duration.millis(32));
        Runnable paintTimelineViewport =
                () -> {
                    int nSlots = slotColCount;
                    int[] vr =
                            visibleSlotRangeInclusive(
                                    rightBodyScroll,
                                    layoutViewport.slotWidth,
                                    nSlots,
                                    VIEWPORT_SLOT_MARGIN);
                    GraphicsContext hg = headerCanvas.getGraphicsContext2D();
                    hg.clearRect(0, 0, canvasTimelineW, canvasHeaderH);
                    drawTimeAxis(
                            hg,
                            parsed,
                            canvasTimelineW,
                            layoutViewport,
                            palette,
                            vr[0],
                            vr[1]);
                    for (ViewportRowSpec s : viewportRowSpecs) {
                        GraphicsContext gcx = s.canvas().getGraphicsContext2D();
                        gcx.clearRect(0, 0, canvasTimelineW, s.rowH());
                        gcx.translate(0, s.outerPad());
                        drawTimelineRow(
                                gcx,
                                s.slotTexts(),
                                s.machineGroupIndex(),
                                layoutViewport,
                                palette,
                                barFont,
                                vr[0],
                                vr[1]);
                        gcx.translate(0, -s.outerPad());
                    }
                    if (Boolean.getBoolean(PROFILE_PROP)) {
                        /* プロファイル用: -Dpm.ai.gantt.profile=true */
                    }
                };
        viewportRepaintDebounce.setOnFinished(e -> paintTimelineViewport.run());
        Runnable scheduleViewportRepaint =
                () -> {
                    viewportRepaintDebounce.stop();
                    viewportRepaintDebounce.playFromStart();
                };
        rightBodyScroll.hvalueProperty().addListener((o, a, b) -> scheduleViewportRepaint.run());
        rightBodyScroll.widthProperty().addListener((o, a, b) -> scheduleViewportRepaint.run());
        Platform.runLater(paintTimelineViewport);

        double shiftSens =
                effectiveShiftWheelHorizontalSensitivity(shiftWheelHorizontalSensitivityPercent);
        installShiftWheelHorizontalScroll(rightBodyScroll, shiftSens);

        HBox bodySplit = new HBox(0, leftBodyScroll, rightBodyScroll);
        bodySplit.setAlignment(Pos.TOP_LEFT);
        HBox.setHgrow(rightBodyScroll, Priority.ALWAYS);
        VBox.setVgrow(bodySplit, Priority.ALWAYS);
        headRow.setCache(false);
        bodySplit.setCache(false);

        int nonSectionRowCount = countNonSectionDisplayRows(parsed.displayRows());
        int dataRowsWithTimelineText = countDataRowsWithAnyNonEmptySlot(parsed.displayRows());
        boolean warnEmptyTimelineCells =
                nonSectionRowCount > 0 && dataRowsWithTimelineText == 0;
        VBox mainColumn;
        if (warnEmptyTimelineCells) {
            Label slotWarn =
                    new Label(
                            """
                            タイムライン（HH:MM 列）のセルに表示文言がありません（帯色のみになります）。ブック JSON 単体ではシェイプ由来の値が欠損することがあります。…設.json（設備ガント契約）を読んでいても、kwargs_packed.timeline_events が空配列 [] のときはバーを描画できません（段階2の配台でイベントが生成されていない／契約が古い可能性）。計画出力・段階2ログを確認してください。""");
            slotWarn.setWrapText(true);
            slotWarn.setPadding(new Insets(8));
            slotWarn.setStyle("-fx-background-color: rgba(255,165,0,0.22);");
            mainColumn = new VBox(0, slotWarn, headRow, bodySplit);
        } else {
            mainColumn = new VBox(0, headRow, bodySplit);
        }
        VBox.setVgrow(bodySplit, Priority.ALWAYS);
        mainColumn.setPadding(new Insets(4));
        mainColumn.setCache(false);
        root.setCenter(mainColumn);
        root.setUserData(new EquipmentGanttViewHandles(rightBodyScroll, scheduleViewportRepaint));

        Label hint =
                new Label(
                        """
                        ヒント: 横スクロールで時刻軸を追えます（左3列は固定）。ホイールボタンを押したままドラッグでもスクロールできます。Shift+ホイールでも横（時刻軸方向）にスクロール（感度はツールバー「Shift横スクロール」）。Ctrl+ホイールで表示倍率。 \
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

    /**
     * 右ボディ {@link ScrollPane} を、ホイール（中）ボタンドラッグでのみパン可能にする（左ドラッグはバッジ移動に使う）。
     */
    private static void installMiddleButtonPanScroll(ScrollPane scroll) {
        final boolean[] middlePanning = {false};
        final double[] lastScene = new double[2];
        scroll.addEventFilter(
                MouseEvent.MOUSE_PRESSED,
                e -> {
                    if (e.getButton() == MouseButton.MIDDLE) {
                        middlePanning[0] = true;
                        lastScene[0] = e.getSceneX();
                        lastScene[1] = e.getSceneY();
                        e.consume();
                    }
                });
        scroll.addEventFilter(
                MouseEvent.MOUSE_DRAGGED,
                e -> {
                    if (!middlePanning[0]) {
                        return;
                    }
                    double dx = e.getSceneX() - lastScene[0];
                    double dy = e.getSceneY() - lastScene[1];
                    lastScene[0] = e.getSceneX();
                    lastScene[1] = e.getSceneY();
                    panScrollPaneByPixelDelta(scroll, dx, dy);
                    e.consume();
                });
        scroll.addEventFilter(
                MouseEvent.MOUSE_RELEASED,
                e -> {
                    if (e.getButton() == MouseButton.MIDDLE) {
                        middlePanning[0] = false;
                    }
                });
    }

    private static void panScrollPaneByPixelDelta(ScrollPane scroll, double sceneDx, double sceneDy) {
        Node content = scroll.getContent();
        if (content == null) {
            return;
        }
        Bounds vp = scroll.getViewportBounds();
        if (vp == null || vp.getWidth() <= 0 || vp.getHeight() <= 0) {
            return;
        }
        double viewportW = vp.getWidth();
        double viewportH = vp.getHeight();
        double contentW = content.getLayoutBounds().getWidth();
        double contentH = content.getLayoutBounds().getHeight();
        double excessX = contentW - viewportW;
        double excessY = contentH - viewportH;
        if (excessX > 1.0 && Double.isFinite(excessX)) {
            double nh = scroll.getHvalue() - sceneDx / excessX;
            scroll.setHvalue(Math.clamp(nh, 0.0, 1.0));
        }
        if (excessY > 1.0 && Double.isFinite(excessY)) {
            double nv = scroll.getVvalue() - sceneDy / excessY;
            scroll.setVvalue(Math.clamp(nv, 0.0, 1.0));
        }
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

        /** スロット列幅のみ変更（OOM 対策の自動縮小）。行高・フォントは維持する。 */
        LayoutMetrics scaleSlotWidth(double factor) {
            if (!Double.isFinite(factor) || factor <= 0) {
                return this;
            }
            if (Math.abs(factor - 1.0) < 1e-15) {
                return this;
            }
            return new LayoutMetrics(
                    zoom,
                    slotWidth * factor,
                    rowHeight,
                    sectionRowHeight,
                    headerHeight,
                    labelMinWidth,
                    labelMaxWidth,
                    rowLabelFontSize,
                    axisFontSize,
                    barFontSize,
                    progressFontSize,
                    progressCellWidth,
                    progressGap);
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
            GanttPalette palette,
            int slotVisibleFrom,
            int slotVisibleToIncl) {
        gc.setFill(palette.headerAxis());
        gc.fillRect(0, 0, timelineWidth, layout.headerHeight);
        gc.setStroke(palette.grid());
        gc.setLineWidth(0.35);

        List<Integer> slotCols = parsed.slotColumnIndices();
        int n = slotCols.size();
        if (n <= 0) {
            return;
        }
        LocalTime t0 = parsed.slotBaseTime();
        int slotMin = Math.max(1, parsed.slotMinutes());

        int vf = Math.max(0, slotVisibleFrom);
        int vt = Math.min(n - 1, slotVisibleToIncl);
        if (vf > vt) {
            vf = 0;
            vt = Math.max(0, n - 1);
        }

        for (int i = vf; i <= vt + 1 && i <= n; i++) {
            double x = i * layout.slotWidth;
            gc.strokeLine(x, 0, x, layout.headerHeight);
        }
        gc.strokeRect(0, 0, timelineWidth, layout.headerHeight);

        int labelStep = timeLabelStep(n, slotMin, layout.slotWidth);
        double labelFont = Math.min(layout.axisFontSize, Math.max(7, layout.slotWidth * 0.95));
        gc.setFill(palette.axisLabel());
        gc.setFont(Font.font(labelFont));
        DateTimeFormatter tf = DateTimeFormatter.ofPattern("H:mm");
        int i0 = vf - (vf % Math.max(1, labelStep));
        if (i0 < vf) {
            i0 += labelStep;
        }
        for (int i = Math.max(i0, vf); i <= vt; i += labelStep) {
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

    /**
     * バッジドラッグ位置の保存キー（同一データ・同一レイアルゴリズム前提）。
     * 表示行インデックス・バー占据スロット・セグメント位置・担当名で一意化する。
     */
    private static String personBadgeDragKey(
            int displayRowIndex,
            BarRun run,
            int segmentIndex,
            int indexInSegment,
            String personLabel) {
        return displayRowIndex
                + "|"
                + run.fromSlot()
                + "|"
                + run.toSlot()
                + "|"
                + segmentIndex
                + "|"
                + indexInSegment
                + "|"
                + PersonBadgeStyle.normalizeLabelKey(personLabel);
    }

    /**
     * {@link #build} と同一の pandas 列補正のあと、表本体＋担当バッジ列から SHA-256 フィンガープリントを返す。
     */
    public static String computeDataFingerprint(
            List<String> columns,
            ObservableList<ObservableList<String>> rows,
            List<List<String>> badgeSlotRowsRaw) {
        if (columns == null || rows == null) {
            return "";
        }
        RepairedGanttTable rt = RepairedGanttTable.from(columns, rows, badgeSlotRowsRaw);
        return fingerprintFromRepairedSheets(rt.effCols(), rt.effRows(), rt.badgeEff());
    }

    private static String fingerprintFromRepairedSheets(
            List<String> effCols,
            ObservableList<ObservableList<String>> effRows,
            List<List<String>> badgeEff) {
        StringBuilder sb = new StringBuilder();
        for (String c : effCols) {
            sb.append('C').append(c != null ? c : "").append('\u001f');
        }
        sb.append('\u001e');
        for (int r = 0; r < effRows.size(); r++) {
            ObservableList<String> row = effRows.get(r);
            for (int c = 0; c < effCols.size(); c++) {
                String cell = row != null && c < row.size() ? row.get(c) : "";
                sb.append(cell != null ? cell : "").append('\u001f');
            }
            sb.append('\u001e');
        }
        if (badgeEff != null) {
            sb.append('\u001d');
            for (List<String> br : badgeEff) {
                if (br == null) {
                    sb.append('\u001e');
                    continue;
                }
                for (String x : br) {
                    sb.append(x != null ? x : "").append('\u001f');
                }
                sb.append('\u001e');
            }
        }
        return sha256Hex(sb.toString());
    }

    private static String sha256Hex(String text) {
        try {
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            byte[] digest = md.digest(text.getBytes(StandardCharsets.UTF_8));
            StringBuilder out = new StringBuilder(digest.length * 2);
            for (byte b : digest) {
                out.append(String.format(Locale.US, "%02x", b & 0xff));
            }
            return out.toString();
        } catch (java.security.NoSuchAlgorithmException e) {
            return Integer.toHexString(text.hashCode());
        }
    }

    private record BadgeWirePlacement(
            Line wireLineOrNull, StackPane badge, double anchorX, double anchorY) {}

    private static double barAnchorCenterX(LayoutMetrics layout, BarRun run) {
        return EquipmentGanttWireAnchorMath.barAnchorCenterX(
                layout.slotWidth(), layout.zoom(), run.fromSlot(), run.toSlot());
    }

    private static double barAnchorCenterY(double timelineOuterPad, LayoutMetrics layout) {
        return EquipmentGanttWireAnchorMath.barAnchorCenterY(
                timelineOuterPad, layout.rowHeight(), layout.zoom());
    }

    private static void installPersonBadgeWireListeners(
            Line line, StackPane sp, double anchorX, double anchorY) {
        if (line == null || sp == null) {
            return;
        }
        line.setStartX(anchorX);
        line.setStartY(anchorY);
        Runnable sync =
                () -> {
                    Bounds cb = badgeDragClampBounds(sp);
                    double halfW = cb.getWidth() <= 1e-9 ? 0 : cb.getWidth() / 2;
                    double halfH = cb.getHeight() <= 1e-9 ? 0 : cb.getHeight() / 2;
                    double ex = sp.getLayoutX() + cb.getMinX() + halfW;
                    double ey = sp.getLayoutY() + cb.getMinY() + halfH;
                    line.setEndX(ex);
                    line.setEndY(ey);
                };
        sp.layoutXProperty().addListener(obs -> sync.run());
        sp.layoutYProperty().addListener(obs -> sync.run());
        sync.run();
    }

    private static void layoutPersonBadgeOverlay(
            Pane overlay,
            List<String> badgeSlotTexts,
            List<String> slotTexts,
            LayoutMetrics layout,
            Function<String, PersonBadgeStyle> styleForLabel,
            int displayRowIndex,
            double personBadgeGapPx,
            boolean badgeDragAdjustEnabled,
            double timelineOuterPad,
            double timelinePaneWidth,
            double overlayPaneHeight,
            Map<String, EquipmentGanttBadgeDragDelta> dragDeltas,
            BiConsumer<String, EquipmentGanttBadgeDragDelta> dragDeltaSink,
            boolean personBadgeWiresEnabled,
            Color personBadgeWireColor) {
        if (overlay == null
                || styleForLabel == null
                || badgeSlotTexts == null
                || slotTexts == null
                || badgeSlotTexts.size() != slotTexts.size()) {
            return;
        }
        // スロット文言がスロット按分の m で微妙に異なると collectBarRuns が細切れになり、
        // バッジがスロット個数ぶん重複描画される。バッジ結合だけ末尾「◯◯m」を除いたキーでまとめる。
        List<BarRun> runs = collectBarRunsForPersonBadges(slotTexts);
        /*
         * 同一行に複数 BarRun（別タスクバー）があると、Run ごとのループでは座標が独立し
         * 別バーのバッジ同士が同じ Y 帯で重なる。論理矩形を蓄積して右へ押し出す。
         */
        List<BoundingBox> placedBadgeLogicalRects = new ArrayList<>();
        List<BadgeWirePlacement> wirePlacementBatch = new ArrayList<>();
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
            boolean wireThisRun =
                    personBadgeWiresEnabled
                            && personBadgeWireColor != null;
            double anchorX = 0d;
            double anchorY = 0d;
            if (wireThisRun) {
                anchorX = barAnchorCenterX(layout, run);
                anchorY = barAnchorCenterY(timelineOuterPad, layout);
            }
            List<String> parts = PersonNameBadgeText.splitBadgeCell(frag);
            if (parts.isEmpty()) {
                continue;
            }
            List<StackPane> nodes = new ArrayList<>();
            List<Bounds> locals = new ArrayList<>();
            /** 帯内の縦積み・中央寄せはピル寸法（pref）で行う。bounds は DropShadow で縦に肥大化し帯外へ押し出す。 */
            List<Double> stackHeights = new ArrayList<>();
            for (String p : parts) {
                PersonBadgeStyle st = styleForLabel.apply(p);
                StackPane sp =
                        PersonBadgeNodeFactory.createBadge(
                                p, st, layout.zoom, layout.rowLabelFontSize);
                sp.applyCss();
                sp.layout();
                Bounds lb = sp.getBoundsInLocal();
                if (lb == null
                        || !Double.isFinite(lb.getWidth())
                        || !Double.isFinite(lb.getHeight())
                        || lb.getWidth() < 0.5
                        || lb.getHeight() < 0.5) {
                    lb = computeBadgeFallbackBounds(sp);
                }
                locals.add(lb);
                nodes.add(sp);
                stackHeights.add(Math.max(1.0, sp.prefHeight(-1)));
            }

            double inset = 0.5 * layout.zoom;
            double innerBarTop = 3 * layout.zoom;
            double barTop = timelineOuterPad + innerBarTop;
            double barH = layout.rowHeight - 2 * innerBarTop;
            double xPad = 3 * layout.zoom;
            double bandRight = (run.toSlot() + 1) * layout.slotWidth - inset;
            double x0 = run.fromSlot() * layout.slotWidth + inset + xPad;

            /*
             * 横: personBadgeGapPx を隣接ピル左端間の追加距離として足す（論理幅 + gap）。
             * 縦: 間隔 0 のときは対角オフセット（diagDy）を使わない。
             */
            double gapPx = Math.max(0d, personBadgeGapPx);
            double diagDy =
                    Math.min(
                            Math.max(3 * layout.zoom, 4),
                            Math.max(6, barH / Math.max(8, nodes.size() + 4)));
            double diagDyEff = gapPx <= 1e-9 ? 0.0 : diagDy;
            double segmentGapEff = Math.max(2 * layout.zoom, diagDyEff * 0.45);

            List<List<Integer>> segments = new ArrayList<>();
            List<Integer> curSeg = new ArrayList<>();
            double xVisPack = x0;
            for (int i = 0; i < nodes.size(); i++) {
                StackPane spNode = nodes.get(i);
                Bounds b = locals.get(i);
                double stepW = badgeStepWidthLogical(spNode);
                if (!curSeg.isEmpty()
                        && xVisPack > x0
                        && xVisPack + stepW > bandRight + 1e-6) {
                    segments.add(curSeg);
                    curSeg = new ArrayList<>();
                    xVisPack = x0;
                }
                curSeg.add(i);
                double advance = stepW + gapPx;
                xVisPack += advance;
            }
            if (!curSeg.isEmpty()) {
                segments.add(curSeg);
            }

            List<Double> segHeights = new ArrayList<>();
            for (List<Integer> seg : segments) {
                double maxPref = 1.0;
                for (int ii : seg) {
                    maxPref = Math.max(maxPref, stackHeights.get(ii));
                }
                double span =
                        seg.size() <= 1
                                ? maxPref
                                : (seg.size() - 1) * diagDyEff + maxPref;
                segHeights.add(span);
            }

            double totalStackH = 0;
            for (int i = 0; i < segHeights.size(); i++) {
                totalStackH += segHeights.get(i);
                if (i + 1 < segHeights.size()) {
                    totalStackH += segmentGapEff;
                }
            }
            /*
             * 帯の描画高 barH がバッジより短いと、縦クランプでバッジが同じ Y に寄り横押し出しが増え、
             * UI の横間隔（gapPx）が潰れる。スタック高まで縦クランプ領域を広げる（行 Pane 高で上限）。
             */
            double insetBottom = Math.max(2.0, layout.zoom);
            double maxBottom =
                    Double.isFinite(overlayPaneHeight) && overlayPaneHeight > insetBottom + barTop
                            ? overlayPaneHeight - insetBottom
                            : barTop + barH;
            double desiredClampBottom = barTop + Math.max(barH, totalStackH);
            double badgeClampBottom = Math.min(desiredClampBottom, maxBottom);

            double ySegCursor =
                    barTop + Math.max(0, (barH - totalStackH) / 2);

            for (int si = 0; si < segments.size(); si++) {
                List<Integer> seg = segments.get(si);
                double segRowMax = 1.0;
                for (int ii : seg) {
                    segRowMax = Math.max(segRowMax, stackHeights.get(ii));
                }
                double xVis = x0;
                for (int k = 0; k < seg.size(); k++) {
                    int ii = seg.get(k);
                    StackPane sp = nodes.get(ii);
                    Bounds b = locals.get(ii);
                    double stackH = stackHeights.get(ii);
                    double stepW = badgeStepWidthLogical(sp);
                    double yTop =
                            ySegCursor
                                    + k * diagDyEff
                                    + (segRowMax - stackH) / 2;
                    double ly =
                            clampBadgeLayoutYInBand(
                                    yTop - b.getMinY(), b, barTop, badgeClampBottom);
                    double visualLeft = xVis;
                    Bounds logicalSize = badgeDragClampBounds(sp);
                    BoundingBox proposed =
                            new BoundingBox(
                                    visualLeft,
                                    ly,
                                    logicalSize.getWidth(),
                                    logicalSize.getHeight());
                    proposed =
                            nudgeBadgeLogicalRectToClearPlaced(
                                    proposed,
                                    placedBadgeLogicalRects,
                                    Math.max(0.5, layout.zoom));
                    visualLeft = proposed.getMinX();
                    placedBadgeLogicalRects.add(proposed);

                    String personLabel = ii < parts.size() ? parts.get(ii) : "";
                    String badgeKey =
                            personBadgeDragKey(displayRowIndex, run, si, k, personLabel);
                    Bounds cb = badgeDragClampBounds(sp);
                    double defaultLayoutX = visualLeft - b.getMinX();
                    double defaultLayoutY = ly;
                    double lx = defaultLayoutX;
                    double lyUse = defaultLayoutY;
                    EquipmentGanttBadgeDragDelta sv =
                            dragDeltas != null ? dragDeltas.get(badgeKey) : null;
                    if (sv != null) {
                        lx += sv.dx();
                        lyUse += sv.dy();
                    }
                    lx = clampBadgeLayoutX(lx, cb, timelinePaneWidth);
                    lyUse = clampBadgeLayoutYInBand(lyUse, cb, barTop, badgeClampBottom);
                    sp.setLayoutX(lx);
                    sp.setLayoutY(lyUse);
                    Line wireLine = null;
                    if (wireThisRun) {
                        wireLine = new Line(anchorX, anchorY, anchorX, anchorY);
                        wireLine.setStroke(personBadgeWireColor);
                        wireLine.setStrokeWidth(Math.max(0.75, 0.65 * layout.zoom));
                        wireLine.setMouseTransparent(true);
                        wireLine.setPickOnBounds(false);
                    }
                    wirePlacementBatch.add(
                            new BadgeWirePlacement(wireLine, sp, anchorX, anchorY));
                    if (badgeDragAdjustEnabled) {
                        sp.setMouseTransparent(false);
                        sp.setCursor(Cursor.DEFAULT);
                        /*
                         * ドラッグのクランプは DropShadow を含む getBoundsInLocal() を使わない。
                         * グローの広がりがバッジごとに異なり論理高さが帯より大きくなると可動域がほぼゼロになる。
                         */
                        installBadgeDragHandlers(
                                sp,
                                badgeDragClampBounds(sp),
                                barTop,
                                badgeClampBottom,
                                timelinePaneWidth,
                                defaultLayoutX,
                                defaultLayoutY,
                                badgeKey,
                                dragDeltaSink);
                    }
                    double advance = stepW + gapPx;
                    xVis = visualLeft + advance;
                }
                ySegCursor += segHeights.get(si) + segmentGapEff;
            }
        }
        for (BadgeWirePlacement wp : wirePlacementBatch) {
            if (wp.wireLineOrNull() != null) {
                overlay.getChildren().add(wp.wireLineOrNull());
            }
        }
        for (BadgeWirePlacement wp : wirePlacementBatch) {
            overlay.getChildren().add(wp.badge());
        }
        for (BadgeWirePlacement wp : wirePlacementBatch) {
            if (wp.wireLineOrNull() != null) {
                installPersonBadgeWireListeners(
                        wp.wireLineOrNull(), wp.badge(), wp.anchorX(), wp.anchorY());
            }
        }
    }

    /** 既配置バッジの論理矩形と重ならないよう、水平方向にだけ押し出す（同一行・複数 BarRun 対策）。 */
    private static BoundingBox nudgeBadgeLogicalRectToClearPlaced(
            BoundingBox proposed,
            List<BoundingBox> placed,
            double gap) {
        BoundingBox cur = proposed;
        for (int guard = 0; guard < 512; guard++) {
            boolean moved = false;
            for (BoundingBox o : placed) {
                if (!badgeLogicalRectsOverlapWithSeparation(cur, o, gap)) {
                    continue;
                }
                double shift = o.getMaxX() + gap - cur.getMinX();
                if (shift > 1e-6) {
                    cur =
                            new BoundingBox(
                                    cur.getMinX() + shift,
                                    cur.getMinY(),
                                    cur.getWidth(),
                                    cur.getHeight());
                    moved = true;
                    break;
                }
            }
            if (!moved) {
                break;
            }
        }
        return cur;
    }

    private static boolean badgeLogicalRectsOverlapWithSeparation(
            BoundingBox a, BoundingBox b, double gap) {
        return !(a.getMaxX() + gap <= b.getMinX()
                || b.getMaxX() + gap <= a.getMinX()
                || a.getMaxY() + gap <= b.getMinY()
                || b.getMaxY() + gap <= a.getMinY());
    }

    /**
     * バッジドラッグの「つかみ」判定を、見た目矩形の中央に限定する（幅・高さそれぞれこの割合の矩形）。
     * {@link StackPane#getBoundsInLocal()} は DropShadow 等を含み、{@code sceneToLocal} の座標系と一致する。
     * クランプ用の論理ピル（pref のみ）とは寸法が異なるため、掴み判定だけ実ヒット矩形を使う。
     */
    private static final double BADGE_DRAG_GRAB_ZONE_FRAC = 0.5;

    /** {@link #BADGE_DRAG_GRAB_ZONE_FRAC} に基づき、{@link StackPane#getBoundsInLocal()} 内の中央域にあるか。 */
    private static boolean isWithinBadgeDragGrabZone(StackPane sp, double localX, double localY) {
        if (sp == null) {
            return true;
        }
        Bounds b = sp.getBoundsInLocal();
        if (b == null) {
            return true;
        }
        double w = b.getWidth();
        double h = b.getHeight();
        if (!Double.isFinite(w) || !Double.isFinite(h) || w <= 1e-6 || h <= 1e-6) {
            return true;
        }
        double cx = (b.getMinX() + b.getMaxX()) / 2;
        double cy = (b.getMinY() + b.getMaxY()) / 2;
        double halfW = (w * BADGE_DRAG_GRAB_ZONE_FRAC) / 2;
        double halfH = (h * BADGE_DRAG_GRAB_ZONE_FRAC) / 2;
        return Math.abs(localX - cx) <= halfW && Math.abs(localY - cy) <= halfH;
    }

    /** 掴み可能ホバー時のみ MOVE、それ以外は DEFAULT（バッジ全域が MOVE に見える問題の回避）。 */
    private static void updateBadgeDragHoverCursor(StackPane sp, double localX, double localY) {
        sp.setCursor(
                isWithinBadgeDragGrabZone(sp, localX, localY)
                        ? Cursor.MOVE
                        : Cursor.DEFAULT);
    }

    /**
     * 子 {@link Label} がイベントソースのとき {@code e.getX/Y} は Label 座標になるため、シーン座標から
     * {@link StackPane} ローカルへ変換して掴み判定・カーソル同期に使う。
     */
    private static Point2D badgeMouseLocalInStackPane(StackPane sp, MouseEvent e) {
        return sp.sceneToLocal(e.getSceneX(), e.getSceneY());
    }

    private static void installBadgeDragHandlers(
            StackPane sp,
            Bounds local,
            double bandTop,
            double bandBottom,
            double paneWidth,
            double defaultLayoutX,
            double defaultLayoutY,
            String badgeKey,
            BiConsumer<String, EquipmentGanttBadgeDragDelta> dragDeltaSink) {
        final double[] press = new double[4];
        final boolean[] dragged = {false};
        final boolean[] armed = {false};
        sp.addEventHandler(
                MouseEvent.MOUSE_MOVED,
                e -> {
                    if (armed[0]) {
                        sp.setCursor(Cursor.MOVE);
                        return;
                    }
                    Point2D lp = badgeMouseLocalInStackPane(sp, e);
                    if (lp != null
                            && Double.isFinite(lp.getX())
                            && Double.isFinite(lp.getY())) {
                        updateBadgeDragHoverCursor(sp, lp.getX(), lp.getY());
                    }
                });
        sp.addEventHandler(
                MouseEvent.MOUSE_ENTERED,
                e -> {
                    if (armed[0]) {
                        sp.setCursor(Cursor.MOVE);
                        return;
                    }
                    Point2D lp = badgeMouseLocalInStackPane(sp, e);
                    if (lp != null
                            && Double.isFinite(lp.getX())
                            && Double.isFinite(lp.getY())) {
                        updateBadgeDragHoverCursor(sp, lp.getX(), lp.getY());
                    }
                });
        sp.addEventHandler(
                MouseEvent.MOUSE_EXITED,
                e -> {
                    if (armed[0]) {
                        sp.setCursor(Cursor.MOVE);
                    } else {
                        sp.setCursor(Cursor.DEFAULT);
                    }
                });
        sp.setOnMousePressed(
                e -> {
                    Point2D lp = badgeMouseLocalInStackPane(sp, e);
                    boolean lpFinite =
                            lp != null
                                    && Double.isFinite(lp.getX())
                                    && Double.isFinite(lp.getY());
                    boolean withinGrab =
                            lpFinite && isWithinBadgeDragGrabZone(sp, lp.getX(), lp.getY());
                    armed[0] =
                            e.getButton() == MouseButton.PRIMARY && withinGrab;
                    if (!armed[0]) {
                        if (lpFinite) {
                            updateBadgeDragHoverCursor(sp, lp.getX(), lp.getY());
                        }
                        return;
                    }
                    dragged[0] = false;
                    press[0] = e.getSceneX();
                    press[1] = e.getSceneY();
                    press[2] = sp.getLayoutX();
                    press[3] = sp.getLayoutY();
                    sp.setCursor(Cursor.MOVE);
                    e.consume();
                });
        sp.setOnMouseDragged(
                e -> {
                    if (!armed[0]) {
                        return;
                    }
                    dragged[0] = true;
                    sp.setCursor(Cursor.MOVE);
                    double dx = e.getSceneX() - press[0];
                    double dy = e.getSceneY() - press[1];
                    double nx = clampBadgeLayoutX(press[2] + dx, local, paneWidth);
                    double ny =
                            clampBadgeLayoutYInBand(press[3] + dy, local, bandTop, bandBottom);
                    sp.setLayoutX(nx);
                    sp.setLayoutY(ny);
                    e.consume();
                });
        sp.setOnMouseReleased(
                e -> {
                    if (e.getButton() != MouseButton.PRIMARY) {
                        return;
                    }
                    boolean sink =
                            armed[0]
                                    && dragged[0]
                                    && dragDeltaSink != null
                                    && badgeKey != null
                                    && !badgeKey.isEmpty();
                    if (sink) {
                        double tdx = sp.getLayoutX() - defaultLayoutX;
                        double tdy = sp.getLayoutY() - defaultLayoutY;
                        dragDeltaSink.accept(
                                badgeKey, new EquipmentGanttBadgeDragDelta(tdx, tdy));
                        e.consume();
                    }
                    armed[0] = false;
                    dragged[0] = false;
                    Point2D lp = badgeMouseLocalInStackPane(sp, e);
                    if (lp != null
                            && Double.isFinite(lp.getX())
                            && Double.isFinite(lp.getY())) {
                        updateBadgeDragHoverCursor(sp, lp.getX(), lp.getY());
                    }
                });
    }

    private static double clampBadgeLayoutX(double layoutX, Bounds local, double paneWidth) {
        if (local == null || paneWidth <= 1.0) {
            return layoutX;
        }
        double lx = layoutX;
        double left = lx + local.getMinX();
        double right = lx + local.getMaxX();
        if (left < 0) {
            lx -= left;
        }
        right = lx + local.getMaxX();
        if (right > paneWidth) {
            lx -= right - paneWidth;
        }
        left = lx + local.getMinX();
        if (left < 0) {
            lx -= left;
        }
        return lx;
    }

    private static Bounds computeBadgeFallbackBounds(StackPane sp) {
        double w = Math.max(1.0, sp.prefWidth(-1));
        double h = Math.max(1.0, sp.prefHeight(-1));
        return new BoundingBox(0, 0, w, h);
    }

    /**
     * ドラッグ時の左右・上下クランプに使うピル本体の論理矩形（原点付き）。
     * {@link StackPane#getBoundsInLocal()} はエフェクトで縦横に肥大するため可動域がバッジごとに潰れる。
     */
    private static Bounds badgeDragClampBounds(StackPane sp) {
        return computeBadgeFallbackBounds(sp);
    }

    private static double visualWidth(Bounds b) {
        if (b == null) {
            return 1.0;
        }
        double w = b.getMaxX() - b.getMinX();
        return Math.max(1.0, Double.isFinite(w) ? w : b.getWidth());
    }

    /** 横方向の 1 バッジ分のステップ（グロー除く論理幅）。 */
    private static double badgeStepWidthLogical(StackPane sp) {
        return Math.max(1.0, badgeDragClampBounds(sp).getWidth());
    }

    /**
     * 親座標系での見かけ矩形が {@code [bandTop, bandBottom]} に収まるよう layoutY を調整する（グローのはみ出し対策）。
     */
    private static double clampBadgeLayoutYInBand(
            double layoutY, Bounds local, double bandTop, double bandBottom) {
        if (local == null) {
            return layoutY;
        }
        double ly = layoutY;
        double top = ly + local.getMinY();
        double bot = ly + local.getMaxY();
        if (top < bandTop) {
            ly += bandTop - top;
        }
        bot = ly + local.getMaxY();
        if (bot > bandBottom) {
            ly -= bot - bandBottom;
        }
        top = ly + local.getMinY();
        if (top < bandTop) {
            ly += bandTop - top;
        }
        return ly;
    }

    private static void drawTimelineRow(
            GraphicsContext gc,
            List<String> slotTexts,
            int machineGroupIndex,
            LayoutMetrics layout,
            GanttPalette palette,
            Font barFont,
            int slotVisibleFrom,
            int slotVisibleToIncl) {
        int n = slotTexts.size();
        int vf = Math.max(0, slotVisibleFrom);
        int vt = Math.min(n - 1, slotVisibleToIncl);
        if (vf > vt) {
            vf = 0;
            vt = Math.max(0, n - 1);
        }

        Color band = palette.machineBandFill(machineGroupIndex);
        for (int i = vf; i <= vt; i++) {
            double x = i * layout.slotWidth;
            gc.setFill(band);
            gc.fillRect(x, 0, layout.slotWidth, layout.rowHeight);
        }
        gc.setStroke(palette.grid());
        gc.setLineWidth(0.35);
        for (int i = vf; i <= vt + 1 && i <= n; i++) {
            gc.strokeLine(i * layout.slotWidth, 0, i * layout.slotWidth, layout.rowHeight);
        }

        List<BarRun> runs = collectBarRuns(slotTexts);
        for (BarRun run : runs) {
            if (run.toSlot() < vf || run.fromSlot() > vt) {
                continue;
            }
            fillBar(gc, run, layout, palette);
        }
        drawBarLabelsOutside(gc, runs, layout, palette, barFont, vf, vt);
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

    /**
     * 担当バッジの横並び位置は連続スロットを1ランにまとめる必要がある。
     * タイムライン文言が「依頼NO 123.4m」のようにスロットごとに変わる場合でも、依頼NO単位で結合するためのキー。
     */
    private static String personBadgeRunMergeKey(String slotText) {
        String t = slotText != null ? slotText.strip() : "";
        if (t.isEmpty()) {
            return "";
        }
        BarKind k = classifyBar(t);
        if (k == BarKind.BREAK || k == BarKind.STARTUP) {
            return k + "\u0001" + t;
        }
        String base = t.replaceFirst("\\s+\\d+(?:\\.\\d+)?m\\s*$", "").strip();
        String identity = base.isEmpty() ? t : base;
        return k + "\u0001" + identity;
    }

    /**
     * {@link #collectBarRuns} と同様だが、連続判定に {@link #personBadgeRunMergeKey} を使う（バッジ重複抑制用）。
     */
    private static List<BarRun> collectBarRunsForPersonBadges(List<String> slotTexts) {
        int n = slotTexts.size();
        List<BarRun> runs = new ArrayList<>();
        int runStart = -1;
        String runKey = "";
        String headText = "";
        for (int i = 0; i < n; i++) {
            String t = slotTexts.get(i) != null ? slotTexts.get(i).strip() : "";
            boolean empty = t.isEmpty();
            if (empty) {
                if (runStart >= 0) {
                    runs.add(
                            new BarRun(
                                    runStart,
                                    i - 1,
                                    headText,
                                    classifyBar(headText)));
                    runStart = -1;
                    runKey = "";
                    headText = "";
                }
                continue;
            }
            String key = personBadgeRunMergeKey(t);
            if (runStart < 0) {
                runStart = i;
                runKey = key;
                headText = t;
            } else if (!key.equals(runKey)) {
                runs.add(
                        new BarRun(
                                runStart,
                                i - 1,
                                headText,
                                classifyBar(headText)));
                runStart = i;
                runKey = key;
                headText = t;
            }
        }
        if (runStart >= 0) {
            runs.add(
                    new BarRun(
                            runStart,
                            n - 1,
                            headText,
                            classifyBar(headText)));
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
            Font barFont,
            int slotVisibleFrom,
            int slotVisibleToIncl) {
        List<BarRun> sorted = new ArrayList<>();
        for (BarRun run : runs) {
            if (run.toSlot() < slotVisibleFrom || run.fromSlot() > slotVisibleToIncl) {
                continue;
            }
            sorted.add(run);
        }
        if (sorted.isEmpty()) {
            return;
        }
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
        return switch (GanttScheduleSlotBarKind.fromTimelineCell(t)) {
            case BREAK -> BarKind.BREAK;
            case STARTUP -> BarKind.STARTUP;
            default -> BarKind.DEFAULT;
        };
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

    private record ViewportRowSpec(
            Canvas canvas,
            double outerPad,
            List<String> slotTexts,
            int machineGroupIndex,
            double rowH) {}

    /**
     * {@link #build} と {@link #computeDataFingerprint} で同一の pandas 救済を適用した列・行・担当バッジ列。
     */
    private record RepairedGanttTable(
            List<String> effCols,
            ObservableList<ObservableList<String>> effRows,
            List<List<String>> badgeEff) {
        static RepairedGanttTable from(
                List<String> columns,
                ObservableList<ObservableList<String>> rows,
                List<List<String>> badgeSlotRowsRaw) {
            List<String> effCols = columns;
            ObservableList<ObservableList<String>> effRows = rows;
            List<List<String>> badgeEff = badgeSlotRowsRaw;
            RepairResult repaired = tryRepairPandasUnnamedEquipmentTimeline(effCols, effRows);
            if (repaired != null) {
                return new RepairedGanttTable(repaired.columns(), repaired.rows(), null);
            }
            return new RepairedGanttTable(effCols, effRows, badgeEff);
        }
    }

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

    private static int countNonSectionDisplayRows(List<DisplayRow> displayRows) {
        int n = 0;
        for (DisplayRow dr : displayRows) {
            if (dr.sectionBanner() == null) {
                n++;
            }
        }
        return n;
    }

    /** いずれかのタイムスロット列に非空白があるデータ行数（バー描画可否の集計） */
    private static int countDataRowsWithAnyNonEmptySlot(List<DisplayRow> displayRows) {
        int n = 0;
        for (DisplayRow dr : displayRows) {
            if (dr.sectionBanner() != null) {
                continue;
            }
            List<String> cells = dr.cellsInSlots();
            if (cells == null) {
                continue;
            }
            boolean any = false;
            for (String s : cells) {
                if (s != null && !s.isBlank()) {
                    any = true;
                    break;
                }
            }
            if (any) {
                n++;
            }
        }
        return n;
    }

    /** 先頭のデータ行（セクション行以外）における非空スロットセル数。バー非表示調査用 */
    private static int countNonEmptySlotsFirstDataRow(List<DisplayRow> displayRows) {
        for (DisplayRow dr : displayRows) {
            if (dr.sectionBanner() != null) {
                continue;
            }
            List<String> cells = dr.cellsInSlots();
            if (cells == null) {
                return 0;
            }
            int c = 0;
            for (String s : cells) {
                if (s != null && !s.isBlank()) {
                    c++;
                }
            }
            return c;
        }
        return 0;
    }

    private static List<String> firstDataRowSlotSample(List<DisplayRow> displayRows, int maxCells) {
        List<String> out = new ArrayList<>();
        for (DisplayRow dr : displayRows) {
            if (dr.sectionBanner() != null) {
                continue;
            }
            List<String> cells = dr.cellsInSlots();
            if (cells == null) {
                return out;
            }
            int lim = Math.min(maxCells, cells.size());
            for (int i = 0; i < lim; i++) {
                String s = cells.get(i);
                String t = s == null ? "" : s.strip();
                if (t.length() > 40) {
                    t = t.substring(0, 40) + "…";
                }
                out.add(t);
            }
            return out;
        }
        return out;
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
