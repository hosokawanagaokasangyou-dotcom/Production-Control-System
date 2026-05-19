package jp.co.pm.ai.desktop.print;

import java.time.LocalTime;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

import javafx.collections.ObservableList;

import jp.co.pm.ai.desktop.config.PersonBadgeStyle;
import jp.co.pm.ai.desktop.config.EquipmentGanttBadgeDragDelta;

/**
 * 設備ガント 1 印刷ページ分の入力（画面用ノードツリーとは分離して {@link EquipmentGanttPrintCompositor} が組み立てる）。
 */
public record EquipmentGanttPrintPageSpec(
        List<String> columns,
        ObservableList<ObservableList<String>> rows,
        List<List<String>> badgeSlotRows,
        double zoom,
        double rowHeightPercent,
        double slotWidthPercent,
        String barFontFamily,
        double barFontPercent,
        double headerHeightPercent,
        double dateColWidthOverridePx,
        double machineColWidthOverridePx,
        double processColWidthOverridePx,
        boolean showPersonBadges,
        Function<String, PersonBadgeStyle> personBadgeStyleResolver,
        double personBadgeGapPx,
        double personBadgeBandVerticalOffsetPx,
        Map<String, EquipmentGanttBadgeDragDelta> personBadgeDragDeltas,
        String personBadgeWireStrokeHex,
        double personBadgeWireWidthPx,
        String personBadgeWireDashStyleKey,
        double personBadgeWireMaxLengthPx,
        boolean showPersonBadgeWires,
        /** 印刷タイムライン列の半開区間開始（{@code null} は全スロット列）。 */
        LocalTime printTimeRangeStartInclusive,
        /** 印刷タイムライン列の半開区間終了（{@code null} は全スロット列）。 */
        LocalTime printTimeRangeEndExclusive) {}
