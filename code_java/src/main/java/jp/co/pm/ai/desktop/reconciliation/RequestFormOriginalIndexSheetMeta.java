package jp.co.pm.ai.desktop.reconciliation;

import java.util.Map;

/**
 * 依頼書原本「目次」シートの列値を rawMap メタキーとして保持する（UI 表示用・依頼シートマージ値とは別）。
 */
public final class RequestFormOriginalIndexSheetMeta {

    public static final String KEY_RESPONSE_DATE = "_indexSheetResponseDate";
    public static final String KEY_INPUT_DATE = "_indexSheetInputDate";
    public static final String KEY_DELIVERY_DATE = "_indexSheetDeliveryDate";
    public static final String KEY_DELIVERY_REMARKS = "_indexSheetDeliveryRemarks";
    public static final String KEY_CONTRACT_NO = "_indexSheetContractNo";
    public static final String KEY_CONTRACT_REMARKS = "_indexSheetContractRemarks";

    /** 目次マージ前の依頼シート単体の投入日（原反投入日4ソース照合用）。 */
    public static final String KEY_SHEET_INPUT_DATE = "_sheetInputDate";

    /** 目次 1 行分の表示用値（strip 済み、空は {@code ""}）。 */
    public record IndexSheetDisplay(
            String responseDate,
            String inputDate,
            String deliveryDate,
            String deliveryRemarks,
            String contractNo,
            String contractRemarks) {

        public static IndexSheetDisplay empty() {
            return new IndexSheetDisplay("", "", "", "", "", "");
        }

        public static IndexSheetDisplay fromRaw(Map<String, String> raw) {
            if (raw == null || raw.isEmpty()) {
                return empty();
            }
            return new IndexSheetDisplay(
                    nz(raw.get(KEY_RESPONSE_DATE)),
                    nz(raw.get(KEY_INPUT_DATE)),
                    nz(raw.get(KEY_DELIVERY_DATE)),
                    nz(raw.get(KEY_DELIVERY_REMARKS)),
                    nz(raw.get(KEY_CONTRACT_NO)),
                    nz(raw.get(KEY_CONTRACT_REMARKS)));
        }

        public static IndexSheetDisplay fromIndexEntry(
                RequestFormOriginalIndexSheetReader.IndexEntry entry) {
            if (entry == null) {
                return empty();
            }
            return new IndexSheetDisplay(
                    nz(entry.responseDate()),
                    nz(entry.inputDate()),
                    nz(entry.deliveryDate()),
                    nz(entry.deliveryRemarks()),
                    nz(entry.contractNo()),
                    nz(entry.contractRemarks()));
        }

        private static String nz(String value) {
            return value != null ? value.strip() : "";
        }
    }

    private RequestFormOriginalIndexSheetMeta() {}

    static void writeIndexMetaToRawMap(
            Map<String, String> rawMap, RequestFormOriginalIndexSheetReader.IndexEntry indexEntry) {
        if (rawMap == null || indexEntry == null) {
            return;
        }
        IndexSheetDisplay display = IndexSheetDisplay.fromIndexEntry(indexEntry);
        putIfNonBlank(rawMap, KEY_RESPONSE_DATE, display.responseDate());
        putIfNonBlank(rawMap, KEY_INPUT_DATE, display.inputDate());
        putIfNonBlank(rawMap, KEY_DELIVERY_DATE, display.deliveryDate());
        putIfNonBlank(rawMap, KEY_DELIVERY_REMARKS, display.deliveryRemarks());
        putIfNonBlank(rawMap, KEY_CONTRACT_NO, display.contractNo());
        putIfNonBlank(rawMap, KEY_CONTRACT_REMARKS, display.contractRemarks());
    }

    private static void putIfNonBlank(Map<String, String> rawMap, String key, String value) {
        if (value == null || value.isBlank()) {
            rawMap.remove(key);
            return;
        }
        rawMap.put(key, value);
    }
}
