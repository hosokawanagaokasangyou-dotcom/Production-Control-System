package jp.co.pm.ai.desktop.dispatch;

/**
 * 配台計画手動修正タブのワイド表・日別表と同一の、メモリ上 {@link ResultDispatchDocument} の整形手順。
 *
 * <p>{@link jp.co.pm.ai.desktop.DispatchInteractiveTabController} のグリッド再構築の先頭で適用している処理と同一。
 */
public final class ResultDispatchInteractiveGridModel {

    private ResultDispatchInteractiveGridModel() {}

    /**
     * 手動修正タブ表示用に、長形式行をワイド同一視キーでマージし、セル値を正規化する。
     */
    public static void applyWideMergeAndNormalize(ResultDispatchDocument doc) {
        if (doc == null) {
            return;
        }
        ResultDispatchPivot.mergeDispatchRowsByWideIdentity(
                doc.columns(),
                doc.rows(),
                ResultDispatchPivot.DISPATCH_INTERACTIVE_WIDE_MERGE_IDENTITY_HEADERS);
        ResultDispatchNormalizer.normalizeInPlace(doc.columns(), doc.rows());
    }
}
