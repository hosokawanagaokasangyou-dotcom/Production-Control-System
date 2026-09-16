package jp.co.pm.ai.kouchin.trend;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;

import jp.co.pm.ai.kouchin.verify.ExcelValues;
import jp.co.pm.ai.kouchin.verify.FactoryProfile;
import jp.co.pm.ai.kouchin.verify.VerifyException;
import jp.co.pm.ai.kouchin.verify.YearMonthKey;

/**
 * 湖南② 加工賃試算（東レ3シート）の種類別・シート区分別集計。
 */
public final class TrendKonanReader {

    private TrendKonanReader() {}

    public static TrendMonthData read(Path path) {
        List<String> warnings = new ArrayList<>();
        Map<String, TrendMetric> kinds = new LinkedHashMap<>();
        Map<String, TrendMetric> sheets = new LinkedHashMap<>();
        for (String s : FactoryProfile.SHISAN_SHEETS) {
            sheets.put(s, new TrendMetric(0, 0));
        }
        YearMonthKey ym = null;
        try (Workbook wb = ExcelValues.open(path)) {
            int found = 0;
            for (String sname : FactoryProfile.SHISAN_SHEETS) {
                if (wb.getSheet(sname) == null) {
                    warnings.add(path.getFileName() + ": シート「" + sname + "」がありません");
                    continue;
                }
                found++;
                List<List<Object>> rows = ExcelValues.readSheet(wb, sname);
                if (ym == null) {
                    ym = findYm(rows);
                }
                Integer hdr = findHeader(rows);
                if (hdr == null) {
                    warnings.add(path.getFileName() + "「" + sname + "」: ヘッダー行が見つかりません");
                    continue;
                }
                List<Object> header = rows.get(hdr);
                List<Object> nameRow = hdr + 1 < rows.size() ? rows.get(hdr + 1) : List.of();
                Integer amtCol = findHeaderCol(header, "加工賃", "加工種類別");
                Integer qtyBlock = findHeaderCol(header, "加工量", "加工種類別");
                if (amtCol == null) {
                    warnings.add(path.getFileName() + "「" + sname + "」: 加工種類別 加工賃列がありません");
                    continue;
                }
                Integer qcol = findHeaderCol(header, "入庫", "数量");
                if (qcol == null) {
                    qcol = findHeaderCol(header, "入庫");
                }
                List<ColName> wageKindCols = new ArrayList<>();
                int endW = qtyBlock != null ? qtyBlock : nameRow.size();
                for (int j = amtCol + 1; j < endW; j++) {
                    String name = TrendText.disp(at(nameRow, j));
                    if (!name.isEmpty() && !"合計".equals(name)) {
                        wageKindCols.add(new ColName(j, name));
                    }
                }
                List<ColName> qtyKindCols = new ArrayList<>();
                if (qtyBlock != null) {
                    int endQ = Math.max(nameRow.size(), header.size());
                    for (int j = qtyBlock + 1; j < endQ; j++) {
                        String name = TrendText.disp(at(nameRow, j));
                        if (!name.isEmpty() && !"合計".equals(name)) {
                            qtyKindCols.add(new ColName(j, name));
                        }
                    }
                }
                int dataStart = hdr + 4;
                for (int r = dataStart; r < rows.size(); r++) {
                    List<Object> row = rows.get(r);
                    if (row == null || row.isEmpty()) {
                        continue;
                    }
                    String irai = TrendText.nfkc(at(row, 0));
                    String keiyaku = TrendText.nfkc(at(row, 2)).replace("-", "");
                    if (irai.isEmpty() && keiyaku.isEmpty()) {
                        continue;
                    }
                    addSheet(sheets, sname, softNum(row, amtCol), qcol == null ? 0.0 : TrendText.num(softNumObj(row, qcol)));
                    for (ColName col : wageKindCols) {
                        addKindWage(kinds, col.name, softNum(row, col.col));
                    }
                    for (ColName col : qtyKindCols) {
                        addKindQty(kinds, col.name, softNum(row, col.col));
                    }
                }
            }
            if (found == 0) {
                warnings.add(path.getFileName() + ": 東レシートが1つもありません");
            }
        } catch (IOException e) {
            throw new VerifyException("湖南試算を閉じられませんでした: " + path + " (" + e.getMessage() + ")", e);
        }
        return new TrendMonthData(ym, path, kinds, sheets, warnings);
    }

    private static YearMonthKey findYm(List<List<Object>> rows) {
        int limit = Math.min(15, rows.size());
        for (int i = 0; i < limit; i++) {
            for (Object v : rows.get(i)) {
                if (v instanceof String s) {
                    var ym = YearMonthKey.parseGatsudo(s);
                    if (ym.isPresent()) {
                        return ym.get();
                    }
                }
            }
        }
        return null;
    }

    private static Integer findHeader(List<List<Object>> rows) {
        int limit = Math.min(40, rows.size());
        for (int i = 0; i < limit; i++) {
            if (!rows.get(i).isEmpty() && TrendText.nfkc(ExcelValues.at(rows, i, 0)).contains("加工依頼")) {
                return i;
            }
        }
        return null;
    }

    private static Integer findHeaderCol(List<Object> header, String... needles) {
        for (int j = 0; j < header.size(); j++) {
            String v = TrendText.nfkc(header.get(j));
            boolean ok = true;
            for (String n : needles) {
                if (!v.contains(n)) {
                    ok = false;
                    break;
                }
            }
            if (ok) {
                return j;
            }
        }
        return null;
    }

    private static Object at(List<Object> row, int idx) {
        return row != null && idx >= 0 && idx < row.size() ? row.get(idx) : null;
    }

    private static Object softNumObj(List<Object> row, int idx) {
        Object v = at(row, idx);
        if (v instanceof String s && s.isBlank()) {
            return 0;
        }
        return v;
    }

    private static double softNum(List<Object> row, int idx) {
        return TrendText.num(softNumObj(row, idx));
    }

    private static void addSheet(Map<String, TrendMetric> sheets, String name, double wage, double qty) {
        TrendMetric cur = sheets.getOrDefault(name, new TrendMetric(0, 0));
        sheets.put(name, new TrendMetric(cur.wage() + wage, cur.qty() + qty));
    }

    private static void addKindWage(Map<String, TrendMetric> kinds, String name, double wage) {
        TrendMetric cur = kinds.getOrDefault(name, new TrendMetric(0, 0));
        kinds.put(name, new TrendMetric(cur.wage() + wage, cur.qty()));
    }

    private static void addKindQty(Map<String, TrendMetric> kinds, String name, double qty) {
        TrendMetric cur = kinds.getOrDefault(name, new TrendMetric(0, 0));
        kinds.put(name, new TrendMetric(cur.wage(), cur.qty() + qty));
    }

    private record ColName(int col, String name) {}
}
