package jp.co.pm.ai.desktop.reconciliation;

/**
 * 依頼書入力の TPI PDF 再読込で、parse キャッシュを使うか再抽出するかを決める。
 *
 * <p>「データを再読込」はキャッシュを使わず再抽出する。画像スキャン PDF はそのとき再 OCR になる。
 * 起動時や工場切替などの自動再読込は、Excel 原本と重複する PDF を飛ばし、有効な parse キャッシュを使う。
 */
final class RequestFormTpiPdfReload {

    enum Action {
        /** ファイル名の依頼Ｎｏが Excel 原本に既にある。抽出も OCR もしない。 */
        SKIP_EXCEL_DUPLICATE,
        /** parse キャッシュの entries を使う。 */
        USE_CACHE,
        /** PDF を再抽出する。画像スキャンは OCR。 */
        REEXTRACT
    }

    private RequestFormTpiPdfReload() {}

    static Action decide(
            boolean explicitDataReload, boolean excelOriginalHasIrai, boolean parseCacheHit) {
        if (explicitDataReload) {
            return Action.REEXTRACT;
        }
        if (excelOriginalHasIrai) {
            return Action.SKIP_EXCEL_DUPLICATE;
        }
        return parseCacheHit ? Action.USE_CACHE : Action.REEXTRACT;
    }
}
