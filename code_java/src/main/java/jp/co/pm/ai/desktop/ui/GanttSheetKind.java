package jp.co.pm.ai.desktop.ui;

/**
 * ガント風スプレッド用の表現モード。
 */
public enum GanttSheetKind {
    /** ハッシュ色・語義色（メンバー労務など） */
    DEFAULT,
    /**
     * UI参照用の「結果_設備ガント」風（濃納見出し、天の灰軸、青ブロック、進度列の黄系）。
     */
    EQUIPMENT_TIMELINE
}
