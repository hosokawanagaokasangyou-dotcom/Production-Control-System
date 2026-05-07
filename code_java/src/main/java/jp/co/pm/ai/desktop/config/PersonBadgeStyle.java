package jp.co.pm.ai.desktop.config;

import java.text.Normalizer;

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
        double glowSpread,
        /** バッジ全体の不透明度（0〜1） */
        double opacity) {

    /**
     * 担当者ごとのスタイルマップのキーに使う（NFKC・前後空白除去）。
     *
     * <p>ガント上のバッジ表示文字列と一致させる。
     */
    public static String normalizeLabelKey(String raw) {
        if (raw == null) {
            return "";
        }
        return Normalizer.normalize(raw.strip(), Normalizer.Form.NFKC);
    }

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
                0.28,
                1.0);
    }

    /** 実行タブ「ソースキャッシュ」バッジの初期配色（ネットワーク不可フォールバックの視認性）。 */
    public static PersonBadgeStyle networkSourceCacheBadgeDefault() {
        return new PersonBadgeStyle(
                "",
                92,
                "#ea580c",
                "#fffbeb",
                "#c2410c",
                1.2,
                8.0,
                false,
                "#fdba74",
                12.0,
                0.22,
                1.0);
    }
}
