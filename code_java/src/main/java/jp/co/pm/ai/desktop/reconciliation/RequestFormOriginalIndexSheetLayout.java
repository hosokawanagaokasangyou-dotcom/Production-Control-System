package jp.co.pm.ai.desktop.reconciliation;

/**
 * 加工依頼書原本ブック内「目次」シートの列座標正本。
 * POI は 0-based（Excel 列 A → index 0）。
 */
public final class RequestFormOriginalIndexSheetLayout {

    public static final String SHEET_NAME = "目次";

    /** 加工依頼NO（シート名 T6-20 等と対応）。 */
    public static final int COL_IRAI_NO = columnLetterToIndex("A");

    /** 発注依頼日（将来拡張用・今回 rawMap へは未反映）。 */
    public static final int COL_ORDER_REQUEST_DATE = columnLetterToIndex("H");

    /** 回答日（将来拡張用）。 */
    public static final int COL_RESPONSE_DATE = columnLetterToIndex("I");

    /** 投入日 → rawMap {@code 投入日}。 */
    public static final int COL_INPUT_DATE = columnLetterToIndex("J");

    /** 納期 → rawMap {@code 納期回答}（依頼シート U20 と照合）。 */
    public static final int COL_DELIVERY_DATE = columnLetterToIndex("K");

    /** 納期回答・備考（将来拡張用）。 */
    public static final int COL_DELIVERY_REMARKS = columnLetterToIndex("L");

    /** 契約日（将来拡張用）。 */
    public static final int COL_CONTRACT_DATE = columnLetterToIndex("M");

    /** 契約NO → rawMap {@code 契約Ｎｏ}。 */
    public static final int COL_CONTRACT_NO = columnLetterToIndex("N");

    /** 契約・備考（将来拡張用）。 */
    public static final int COL_CONTRACT_REMARKS = columnLetterToIndex("O");

    /** ヘッダ行探索の最大行数（0-based）。 */
    public static final int HEADER_SCAN_MAX_ROW = 30;

    private RequestFormOriginalIndexSheetLayout() {}

    public static int columnLetterToIndex(String letters) {
        return RequestFormOriginalCellLayout.columnLetterToIndex(letters);
    }
}
