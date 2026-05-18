package jp.co.pm.ai.desktop.config;

/**
 * ポータブルバージョンアップの進捗（共有 ZIP 取得 → 展開 → pm-ai-data 同期）。バックグラウンドスレッドから報告し、UI は
 * {@link javafx.application.Platform#runLater(Runnable)} で反映する。
 */
public final class PortableBundleUpgradeProgress {

    /** バージョンアップ処理の段階。 */
    public enum Phase {
        /** 正本 ZIP を共有フォルダからローカルへコピー（取得）。 */
        DOWNLOAD,
        /** ローカル ZIP を一時フォルダへ展開。 */
        EXTRACT,
        /** 展開済み（または正本フォルダ）から pm-ai-data へファイル同期。 */
        SYNC_PM_AI_DATA,
        /** {@code PMD.exe} / {@code app} / {@code runtime} をインストール先へ同期（またはステージング）。 */
        SYNC_DESKTOP
    }

    /** 進捗コールバック。{@code totalUnits <= 0} のときはその段階を不定（indeterminate）扱いにできる。 */
    @FunctionalInterface
    public interface Listener {

        void onProgress(Phase phase, long doneUnits, long totalUnits, String detail);

        /** 段階の開始（合計が未定のときは {@code totalUnits <= 0}）。 */
        default void onPhaseStarted(Phase phase, long totalUnits) {
            onProgress(phase, 0, totalUnits, null);
        }

        /** 段階の完了。 */
        default void onPhaseFinished(Phase phase, long totalUnits) {
            if (totalUnits > 0) {
                onProgress(phase, totalUnits, totalUnits, null);
            } else {
                onProgress(phase, 1, 1, null);
            }
        }
    }

    private PortableBundleUpgradeProgress() {}
}
