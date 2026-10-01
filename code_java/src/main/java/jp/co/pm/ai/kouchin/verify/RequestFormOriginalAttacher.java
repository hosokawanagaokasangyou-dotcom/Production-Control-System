package jp.co.pm.ai.kouchin.verify;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jp.co.pm.ai.desktop.config.AppPaths;
import jp.co.pm.ai.desktop.config.NetworkSourceDirResolver;
import jp.co.pm.ai.desktop.io.PoiWorkbookOpener;
import jp.co.pm.ai.desktop.reconciliation.JuchuTransferValueNormalizer;
import jp.co.pm.ai.desktop.reconciliation.RequestFormOriginalFee;
import jp.co.pm.ai.desktop.reconciliation.RequestFormOriginalFileFinder;

/**
 * 検証A・検証Dの異常行について、依頼書原本の依頼シートを結果ブックへコピーする。
 * 原本フォルダは読み取りのみ。
 */
public final class RequestFormOriginalAttacher {

    public static final String INDEX_SHEET = "依頼書原本";
    private static final Pattern MONTH_PREFIX = Pattern.compile("^[（(]\\d{1,2}月[）)]");
    private static final int MAX_COPY_ROWS = 500;
    private static final int MAX_COPY_COLS = 80;
    private static final DataFormatter FORMATTER = new DataFormatter();

    /** 異常依頼1件。{@code destSheet} はコピーできたときだけ入る。 */
    public record Item(
            String iraiKey,
            String iraiLabel,
            String source,
            Path file,
            String sourceSheet,
            String destSheet,
            String missing,
            RequestFormOriginalFee.Result fee) {}

    public record Plan(List<Item> items, List<String> warnings) {
        public Item byKey(String key) {
            if (key == null || key.isEmpty()) {
                return null;
            }
            for (Item item : items) {
                if (key.equals(item.iraiKey())) {
                    return item;
                }
            }
            return null;
        }

        public int attachedCount() {
            int n = 0;
            for (Item item : items) {
                if (item.destSheet() != null && !item.destSheet().isEmpty()) {
                    n++;
                }
            }
            return n;
        }

        public int missingCount() {
            return items.size() - attachedCount();
        }

        /** 行に載っている依頼NOの原本加工賃（円）合計。1件も計算できなければ null。 */
        public Double amountOf(List<String> irais) {
            if (irais == null || irais.isEmpty()) {
                return null;
            }
            double sum = 0.0;
            boolean any = false;
            for (String irai : irais) {
                Item item = byKey(JuchuTransferValueNormalizer.normalizeKey(irai));
                if (item == null || item.fee() == null) {
                    continue;
                }
                sum += item.fee().amountYen();
                any = true;
            }
            return any ? sum : null;
        }
    }

    private RequestFormOriginalAttacher() {}

    public static Plan prepare(VerifyResult result, Map<String, String> ui) {
        if (result == null) {
            return new Plan(List.of(), List.of());
        }
        Map<String, String> labels = new LinkedHashMap<>();
        Map<String, String> sources = new LinkedHashMap<>();
        Set<String> anomalyKeys = new LinkedHashSet<>();
        collectA(result, labels, sources, anomalyKeys);
        collectD(result, labels, sources, anomalyKeys);
        if (labels.isEmpty()) {
            return new Plan(List.of(), List.of());
        }
        List<String> warnings = new ArrayList<>();
        Map<String, String> env = ui == null ? Map.of() : ui;
        Path dir = AppPaths.resolveRequestFormOriginalDir(env);
        Map<String, RequestFormOriginalFileFinder.Found> found = Map.of();
        if (!NetworkSourceDirResolver.isRequestFormOriginalDirReachable(env)) {
            warnings.add("依頼書原本フォルダにアクセスできません: " + dir);
        } else {
            found = RequestFormOriginalFileFinder.find(dir, labels.keySet(), warnings);
        }
        List<Item> items = new ArrayList<>();
        Set<String> usedNames = new LinkedHashSet<>();
        for (Map.Entry<String, String> e : labels.entrySet()) {
            RequestFormOriginalFileFinder.Found hit = found.get(e.getKey());
            if (hit == null) {
                items.add(new Item(
                        e.getKey(), e.getValue(), sources.get(e.getKey()), null, "", "", "原本なし", null));
                continue;
            }
            if (hit.sheetName() == null || hit.sheetName().isBlank()) {
                items.add(new Item(
                        e.getKey(), e.getValue(), sources.get(e.getKey()), hit.file(), "", "", "依頼シートなし", null));
                continue;
            }
            String dest = anomalyKeys.contains(e.getKey()) ? destSheetName(e.getValue(), usedNames) : "";
            items.add(new Item(
                    e.getKey(), e.getValue(), sources.get(e.getKey()), hit.file(), hit.sheetName(), dest, "",
                    hit.fee()));
        }
        return new Plan(List.copyOf(items), List.copyOf(warnings));
    }

    /** 原本シートを結果ブックの末尾へ値コピーする。原本は変更しない。 */
    public static void copyInto(XSSFWorkbook dest, Plan plan) {
        if (dest == null || plan == null) {
            return;
        }
        Map<Path, List<Item>> byFile = new LinkedHashMap<>();
        for (Item item : plan.items()) {
            if (item.file() == null || item.destSheet() == null || item.destSheet().isEmpty()) {
                continue;
            }
            byFile.computeIfAbsent(item.file(), k -> new ArrayList<>()).add(item);
        }
        for (Map.Entry<Path, List<Item>> e : byFile.entrySet()) {
            try (Workbook src = PoiWorkbookOpener.open(e.getKey().toFile())) {
                for (Item item : e.getValue()) {
                    Sheet from = src.getSheet(item.sourceSheet());
                    if (from == null || dest.getSheet(item.destSheet()) != null) {
                        continue;
                    }
                    copyValues(from, dest.createSheet(item.destSheet()));
                }
            } catch (Exception ex) {
                // 原本が開けなくても検証結果自体は残す。行のリンク側は missing にできないのでシートを作らない。
            }
        }
    }

    static boolean anomalyA(String judge) {
        if (judge == null) {
            return false;
        }
        return switch (judge) {
            case Judge.MATCH, Judge.RESOLVED, Judge.BRANCH_MERGED, Judge.MANUAL_1, Judge.MANUAL_2,
                    Judge.PREV_ADJUST -> false;
            default -> true;
        };
    }

    static List<String> splitIrai(String raw) {
        if (raw == null || raw.isBlank()) {
            return List.of();
        }
        List<String> out = new ArrayList<>();
        for (String part : raw.split("[,、]")) {
            String text = MONTH_PREFIX.matcher(part.strip()).replaceFirst("").strip();
            if (!text.isEmpty()) {
                out.add(text);
            }
        }
        return List.copyOf(out);
    }

    private static void collectA(
            VerifyResult result,
            Map<String, String> labels,
            Map<String, String> sources,
            Set<String> anomalyKeys) {
        if (result.recordsA() == null) {
            return;
        }
        for (RecordA rec : result.recordsA()) {
            if (rec == null) {
                continue;
            }
            addAll(splitIrai(rec.iraiNo()), "検証A", anomalyA(rec.judge()), labels, sources, anomalyKeys);
        }
    }

    private static void collectD(
            VerifyResult result,
            Map<String, String> labels,
            Map<String, String> sources,
            Set<String> anomalyKeys) {
        MatomeCheckResult d = result.checkD();
        if (d == null || d.isSkipped() || d.rows() == null) {
            return;
        }
        for (MatomeCheckResult.MatomeRow row : d.rows()) {
            if (row == null || row.judge() == null || row.judge().isBlank()) {
                continue;
            }
            addAll(splitIrai(row.irai()), "検証D", true, labels, sources, anomalyKeys);
        }
    }

    private static void addAll(
            List<String> irais,
            String source,
            boolean anomaly,
            Map<String, String> labels,
            Map<String, String> sources,
            Set<String> anomalyKeys) {
        for (String irai : irais) {
            String key = JuchuTransferValueNormalizer.normalizeKey(irai);
            if (key.isEmpty()) {
                continue;
            }
            labels.putIfAbsent(key, irai);
            if (anomaly) {
                anomalyKeys.add(key);
            }
            String prev = sources.get(key);
            if (prev == null) {
                sources.put(key, source);
            } else if (!prev.contains(source)) {
                sources.put(key, prev + "・" + source);
            }
        }
    }

    static String destSheetName(String irai, Set<String> used) {
        String cleaned = irai == null ? "依頼" : irai.replaceAll("[\\\\/*?:\\[\\]]", "_").strip();
        if (cleaned.isEmpty()) {
            cleaned = "依頼";
        }
        String base = "原本_" + cleaned;
        if (base.length() > 31) {
            base = base.substring(0, 31);
        }
        String name = base;
        int n = 2;
        while (used.contains(name)) {
            String suffix = "_" + n++;
            int keep = Math.max(1, 31 - suffix.length());
            name = base.substring(0, Math.min(base.length(), keep)) + suffix;
        }
        used.add(name);
        return name;
    }

    private static void copyValues(Sheet src, XSSFSheet dest) {
        int cols = 0;
        int last = Math.min(src.getLastRowNum(), MAX_COPY_ROWS - 1);
        for (int r = 0; r <= last; r++) {
            Row srcRow = src.getRow(r);
            if (srcRow == null) {
                continue;
            }
            Row destRow = dest.createRow(r);
            if (srcRow.getHeight() > 0) {
                destRow.setHeight(srcRow.getHeight());
            }
            int lastCell = Math.min(srcRow.getLastCellNum(), MAX_COPY_COLS);
            cols = Math.max(cols, lastCell);
            for (short c = 0; c < lastCell; c++) {
                Cell srcCell = srcRow.getCell(c);
                if (srcCell == null || srcCell.getCellType() == CellType.BLANK) {
                    continue;
                }
                String text;
                try {
                    text = FORMATTER.formatCellValue(srcCell);
                } catch (RuntimeException ex) {
                    text = "";
                }
                if (!text.isEmpty()) {
                    destRow.createCell(c).setCellValue(text);
                }
            }
        }
        for (int i = 0; i < src.getNumMergedRegions(); i++) {
            CellRangeAddress region = src.getMergedRegion(i);
            if (region.getFirstRow() < MAX_COPY_ROWS && region.getFirstColumn() < MAX_COPY_COLS) {
                int lastRow = Math.min(region.getLastRow(), MAX_COPY_ROWS - 1);
                int lastCol = Math.min(region.getLastColumn(), MAX_COPY_COLS - 1);
                if (lastRow >= region.getFirstRow() && lastCol >= region.getFirstColumn()) {
                    dest.addMergedRegion(new CellRangeAddress(
                            region.getFirstRow(), lastRow, region.getFirstColumn(), lastCol));
                }
            }
        }
        for (int c = 0; c < cols; c++) {
            dest.setColumnWidth(c, src.getColumnWidth(c));
        }
    }
}
