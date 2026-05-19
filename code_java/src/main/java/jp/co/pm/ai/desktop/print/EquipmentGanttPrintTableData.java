package jp.co.pm.ai.desktop.print;

import java.util.List;

import javafx.collections.ObservableList;

/**
 * 設備ガント印刷用に列（タイムライン絞り込み後）を束ねた表データ。
 */
public record EquipmentGanttPrintTableData(
        List<String> columns,
        ObservableList<ObservableList<String>> rows,
        List<List<String>> badgeSlotRows) {}
