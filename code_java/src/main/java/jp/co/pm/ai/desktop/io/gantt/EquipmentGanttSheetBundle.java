package jp.co.pm.ai.desktop.io.gantt;

import java.util.List;

import jp.co.pm.ai.desktop.io.JsonTableIo;

/**
 * 設備ガント契約から組み立てた表と、タイムスロット列と同形状の担当バッジグリッド（行・列対応）。
 */
public record EquipmentGanttSheetBundle(
        JsonTableIo.SheetTable table, List<List<String>> badgeSlotRows) {}
