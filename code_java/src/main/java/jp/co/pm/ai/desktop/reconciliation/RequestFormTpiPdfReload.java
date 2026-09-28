package jp.co.pm.ai.desktop.reconciliation;

/**
 * 依頼書入力の TPI PDF 再読込で、parse キャッシュを使うか再抽出するかを決める。
 *
 * <p>再 OCR が必要なのは新規ファイルと、更新日時またはサイズが変わったファイルだけ。
 * parse キャッシュの指紋が一致する未変更ファイルは再抽出しない。
 * 未変更かつ依頼Ｎｏが Excel 原本にある PDF は、一覧へ足さず OCR もしない。
 */
final class RequestFormTpiPdfReload {

    enum Action {
        /** 未変更で、ファイル名の依頼Ｎｏが Excel 原本に既にある。抽出も OCR もしない。 */
        SKIP_EXCEL_DUPLICATE,
        /** 未変更。parse キャッシュの entries を使う。 */
        USE_CACHE,
        /** 新規、または更新日時・サイズが変わった PDF を再抽出する。画像スキャンは OCR。 */
        REEXTRACT
    }

    private RequestFormTpiPdfReload() {}

    /**
     * @param excelOriginalHasIrai ファイル名から読んだ依頼Ｎｏが Excel 原本にある
     * @param parseCacheHit 更新日時とサイズが parse キャッシュの指紋と一致する
     */
    static Action decide(boolean excelOriginalHasIrai, boolean parseCacheHit) {
        if (!parseCacheHit) {
            return Action.REEXTRACT;
        }
        if (excelOriginalHasIrai) {
            return Action.SKIP_EXCEL_DUPLICATE;
        }
        return Action.USE_CACHE;
    }
}
