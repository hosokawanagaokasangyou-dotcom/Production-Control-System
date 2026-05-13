package jp.co.pm.ai.desktop;

/**
 * {@link PmAiFxApp#main} で Prism（GPU / sw）を決めた結果を保持し、実行・ログタブ表示に使う。
 */
public final class PrismGpuBootstrapStatus {

    public enum Mode {
        /** 起動時プローブで GPU パイプラインを採用 */
        GPU_AFTER_PROBE,
        /** 起動時プローブ失敗・タイムアウト等でソフトウェア描画 */
        SOFTWARE_AFTER_PROBE,
        /**
         * {@code javafx:run} 等で OpenJFX が CLASSPATH（無名モジュール）に載る構成。子 JVM の GPU 試験は合格しても本体では
         * Canvas＋HW Prism で RTTexture NPE になる再現があるため {@code prism.order=sw} に固定した。
         */
        SOFTWARE_CLASSPATH_OPENJFX,
        /** {@code pm.ai.javafx.prism.gpu} 等で GPU を強制（プローブ省略） */
        GPU_OPT_IN,
        /** プローブ省略時の従来ロジック（既定 sw または JVM の prism.order） */
        LEGACY_NO_PROBE,
        /** 初期値（通常は記録されるまで使わない） */
        UNKNOWN
    }

    private static volatile Mode mode = Mode.UNKNOWN;
    private static volatile String detail = "";

    private PrismGpuBootstrapStatus() {}

    static void recordGpuAfterProbe() {
        mode = Mode.GPU_AFTER_PROBE;
        detail = "";
    }

    static void recordSoftwareAfterProbe(String reason) {
        mode = Mode.SOFTWARE_AFTER_PROBE;
        detail = reason != null ? reason : "";
    }

    static void recordSoftwareClasspathOpenJfx() {
        mode = Mode.SOFTWARE_CLASSPATH_OPENJFX;
        detail = "";
    }

    static void recordGpuOptIn() {
        mode = Mode.GPU_OPT_IN;
        detail = "";
    }

    static void recordLegacyNoProbe() {
        mode = Mode.LEGACY_NO_PROBE;
        detail = "";
    }

    /** 実行・ログタブ用の一行（日本語）。 */
    public static String runTabSummary() {
        String order = System.getProperty("prism.order", "");
        String ordShort = order.isEmpty() ? "（既定）" : order;
        return switch (mode) {
            case GPU_AFTER_PROBE ->
                    "JavaFX Prism: GPU 有効（起動時テスト合格） order=" + ordShort;
            case SOFTWARE_AFTER_PROBE ->
                    "JavaFX Prism: ソフトウェア描画（GPU テスト不合格） order="
                            + ordShort
                            + (detail.isBlank() ? "" : " — " + detail);
            case SOFTWARE_CLASSPATH_OPENJFX ->
                    "JavaFX Prism: ソフトウェア描画（CLASSPATH の OpenJFX／Canvas 安定化。GPU 試験は子 JVM） order="
                            + ordShort;
            case GPU_OPT_IN ->
                    "JavaFX Prism: GPU 強制（opt-in） order=" + ordShort;
            case LEGACY_NO_PROBE ->
                    "JavaFX Prism: プローブ省略 order=" + ordShort;
            case UNKNOWN -> "JavaFX Prism: （未記録） order=" + ordShort;
        };
    }

    public static Mode getMode() {
        return mode;
    }
}
