package jp.co.pm.ai.desktop.reconciliation;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.desktop.io.PoiWorkbookOpener;

/** 依頼書原本 xlsm のワークブック単位解析（目次優先 + 依頼シート抽出）。 */
public final class RequestFormOriginalWorkbookParser {

    public static final Pattern ORIGINAL_SHEET_NAME =
            Pattern.compile("^[A-Z]+\\d+-\\d+$|^[A-Z]\\d+-\\d+-\\d+$");

    private RequestFormOriginalWorkbookParser() {}

    public static List<Map<String, String>> parse(File file) throws Exception {
        List<Map<String, String>> parsed = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(file);
                Workbook wb = PoiWorkbookOpener.open(fis)) {
            Map<String, RequestFormOriginalIndexSheetReader.IndexEntry> indexByIrai =
                    readIndexMap(wb);
            for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                String sheetName = wb.getSheetName(s);
                String sheetNameForMatch =
                        JuchuTransferValueNormalizer.toHalfWidthAscii(sheetName)
                                .toUpperCase(java.util.Locale.ROOT);
                if (!ORIGINAL_SHEET_NAME.matcher(sheetNameForMatch).matches()) {
                    continue;
                }
                Sheet rawSheet = wb.getSheetAt(s);
                Map<String, String> raw =
                        RequestFormOriginalExtractor.buildRawMapFromSheet(file, sheetName, rawSheet);
                RequestFormOriginalIndexSheetReader.IndexEntry indexEntry =
                        lookupIndexEntry(indexByIrai, sheetName, raw);
                if (indexEntry != null) {
                    RequestFormOriginalIndexSheetMerger.applyIndexOverrides(raw, indexEntry);
                }
                parsed.add(raw);
            }
        }
        return parsed;
    }

    private static Map<String, RequestFormOriginalIndexSheetReader.IndexEntry> readIndexMap(
            Workbook wb) {
        for (int s = 0; s < wb.getNumberOfSheets(); s++) {
            if (RequestFormOriginalIndexSheetLayout.SHEET_NAME.equals(wb.getSheetName(s))) {
                return RequestFormOriginalIndexSheetReader.read(wb.getSheetAt(s));
            }
        }
        return Map.of();
    }

    private static RequestFormOriginalIndexSheetReader.IndexEntry lookupIndexEntry(
            Map<String, RequestFormOriginalIndexSheetReader.IndexEntry> indexByIrai,
            String sheetName,
            Map<String, String> raw) {
        if (indexByIrai == null || indexByIrai.isEmpty()) {
            return null;
        }
        RequestFormOriginalIndexSheetReader.IndexEntry bySheet =
                indexByIrai.get(JuchuTransferValueNormalizer.normalizeKey(sheetName));
        if (bySheet != null) {
            return bySheet;
        }
        String iraiNo = raw != null ? raw.get("依頼Ｎｏ") : null;
        if (iraiNo == null || iraiNo.isBlank()) {
            return null;
        }
        return indexByIrai.get(JuchuTransferValueNormalizer.normalizeKey(iraiNo));
    }
}
