package jp.co.pm.ai.desktop.config;

/**
 * 設備ガント・担当バッジの見た目（セッション保存・デザインタブと共有）。
 */
public record PersonBadgeStyle(
        String fontFamily,
        /** 行ラベル基準に対するパーセント（例: 85） */
        double fontPercent,
        String fillHex,
        String textHex,
        String strokeHex,
        double strokeWidth,
        double cornerRadius,
        /** {@code true} のとき高さに合わせたカプセル角 */
        boolean pill,
        String glowColorHex,
        double glowRadius,
        /** DropShadow の spread（0〜1 付近） */
        double glowSpread) {

    public static PersonBadgeStyle defaultStyle() {
        return new PersonBadgeStyle(
                "",
                85,
                "#2563eb",
                "#f8fafc",
                "#1d4ed8",
                1.0,
                6.0,
                false,
                "#38bdf8",
                14.0,
                0.28);
    }
}
